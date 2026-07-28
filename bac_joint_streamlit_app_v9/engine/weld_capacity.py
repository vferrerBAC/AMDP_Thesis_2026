"""Welded connection capacity — AISC Chapter J (Block 4, welded path).

Pure-Python port of the "Welded Connections" sheet in
``Capacity Analysis Template.xlsx`` (S. Rosario), and the exact counterpart of
``bolted_capacity.py``. Every line below maps to a column in that sheet so
results are traceable back to the validated workbook:

    E Branch thickness     (branch gauge -> thickness lookup)
    K Chord thickness      (chord gauge  -> thickness lookup)
    M Branch nominal shear (branch material -> Fu lookup, ksi)
    N Chord nominal shear  (chord material  -> Fu lookup, ksi)
    O Base metal area      = F*G - (F - 2E)*(G - 2E)      branch tube wall area
    P Weld length          = F + G
    Q Fexx                 = VALUE(MID(electrode, 2, 1)) * 10
    R Weld leg size        (min(E, K) -> AISC Table J2.4 minimum fillet)
    S Base material str.   = MIN(M, N) * O
    T Weld shear str.      = 0.75 * 0.6 * Q * 0.707 * R * P
    U MaxShearStrength     = MIN(S, T)                    <-- governing capacity
    V Value Used           = "Base Material Strength" if S < T else "Weld Shear Strength"

This is a TUBE-TO-TUBE (HSS branch-to-chord) model: the template's M/N lookups
only resolve "Carbon Steel Tube" and "Stainless Steel Tube". It therefore lines
up directly with Block 1's ``Joint.member_roles`` (branch/chord) and
``Joint.geom_descriptor`` (T/Y/K/X), which is what the joint recommendation
consumes. Any non-tube material routes to review rather than guessing.

Capacities come out of the template in KIPS; this module returns lbf (x1000) so
they line up with the lbf demand side of the pipeline, exactly as the bolted
module does. ``weld_shear_capacity_lbf`` is a drop-in twin of
``bolted_shear_capacity_lbf`` so the recommendation engine can compare the two
connection types without special-casing either.

Notes inherited verbatim from the template (flagged, not changed):
  * Column S applies NEITHER the 0.75 resistance factor NOR the 0.60 shear
    coefficient that column T applies. The two limit states are therefore not
    factored consistently. Reproduced faithfully; confirm with the template
    author before relying on the balance between them.
  * The Fu table values (17-27 ksi) are low for structural steel. Same table as
    the bolted sheet; reproduced faithfully, flagged there too.
  * Column P takes the weld length as F + G (half the perimeter of a
    rectangular branch, not the full 2*(F+G) all-around weld).
  * Column Q reads a SINGLE digit: MID(electrode, 2, 1) * 10. "E70XX" -> 70 is
    correct, but a 3-digit electrode such as "E100XX" would yield 10 ksi. This
    module refuses 3-digit electrodes instead of reproducing that error.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, Optional

PHI = 0.75  # resistance factor, baked into template col T only (see note above)

# --- Lookup tables (verbatim from "Welded Connection Data") -----------------

# Gauge -> thickness (in), sheet P2:Q11. Note this table carries odd gauges
# (9, 11, 13, 15) that the bolted sheet's table does not.
_THICKNESS_BY_GAUGE: Dict[int, float] = {
    8: 0.153,
    9: 0.148,
    10: 0.127,
    11: 0.12,
    12: 0.095,
    13: 0.095,
    14: 0.067,
    15: 0.072,
    16: 0.054,
    18: 0.044,
}

# Material -> ultimate strength Fu (ksi), sheet M2:N9. This is the FULL table
# from the welded sheet; the bolted module's copy omits the stainless rows.
_FU_KSI: Dict[str, float] = {
    "Galvanized Sheet Steel": 19,
    "Stainless Sheet Steel": 20,
    "Carbon Steel Plate": 21,
    "Carbon Steel Pipe": 20,
    "Carbon Steel Tube": 27,
    "Stainless Steel Plate": 17,
    "Stainless Steel Pipe": 17,
    "Stainless Steel Tube": 17,
}

# Columns M and N are IFS() over exactly two materials. Anything else is #N/A in
# the sheet, which we surface as needs_review rather than a silent zero.
_TEMPLATE_TUBE_MATERIALS = ("Carbon Steel Tube", "Stainless Steel Tube")

# Column R: minimum fillet leg size from min(branch t, chord t).
# Matches AISC 360-22 Table J2.4.
_LEG_SIZE_BY_MIN_THICKNESS = (
    (0.25, 0.125),
    (0.50, 0.1875),
    (0.75, 0.25),
    (float("inf"), 0.3125),
)

DEFAULT_ELECTRODE = "E70XX"  # template default; also the practical floor


def thickness_from_gauge(gauge: Any) -> Optional[float]:
    """Column E/K: gauge -> thickness (in). None for an unrecognized gauge;
    callers route that to review rather than guessing a thickness."""
    try:
        return _THICKNESS_BY_GAUGE.get(int(round(float(gauge))))
    except (TypeError, ValueError):
        return None


def weld_leg_size_in(min_thickness_in: float) -> float:
    """Column R: auto-selected fillet leg size. The template DERIVES this from
    the thinner connected part -- it is not a user input."""
    for upper, leg in _LEG_SIZE_BY_MIN_THICKNESS:
        if min_thickness_in <= upper:
            return leg
    return _LEG_SIZE_BY_MIN_THICKNESS[-1][1]


def fexx_from_electrode(electrode: Any) -> Optional[float]:
    """Column Q: VALUE(MID(electrode, 2, 1)) * 10.

    The template reads one digit, so it is only correct for 2-digit electrode
    classifications (E60XX..E90XX). A 3-digit spec like "E100XX" would silently
    become 10 ksi in the sheet; we return None so it routes to review.
    """
    if electrode is None:
        return None
    text = str(electrode).strip().upper()
    match = re.fullmatch(r"E(\d{2,3})X*", text)
    if not match:
        return None
    digits = match.group(1)
    if len(digits) != 2:
        return None  # refuse to reproduce the template's 3-digit truncation
    return float(digits[0]) * 10.0


def map_material(raw: Any) -> Optional[str]:
    """Resolve a material label to a key in the template's Fu table."""
    if not raw:
        return None
    s = str(raw).strip().lower()
    for canonical in _FU_KSI:
        if s == canonical.lower():
            return canonical
    stainless = "stainless" in s
    if "tube" in s:
        return "Stainless Steel Tube" if stainless else "Carbon Steel Tube"
    if "pipe" in s:
        return "Stainless Steel Pipe" if stainless else "Carbon Steel Pipe"
    if "plate" in s:
        return "Stainless Steel Plate" if stainless else "Carbon Steel Plate"
    if "sheet" in s:
        return "Stainless Sheet Steel" if stainless else "Galvanized Sheet Steel"
    if "glv" in s or "galvan" in s or "hdg" in s:
        return "Galvanized Sheet Steel"
    return None


def _num(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except Exception:
        return default


@dataclass
class WeldCapacity:
    """Field-compatible with ``BoltedCapacity`` where the recommendation engine
    reads it, so ``.shear_lbf`` / ``.governing_mode`` / ``.needs_review`` /
    ``.basis`` can be compared across connection types without special-casing.
    """
    shear_lbf: Optional[float]        # col U * 1000 (governing)
    tension_lbf: Optional[float]      # not a limit state in this template
    governing_mode: str               # col V
    needs_review: bool
    basis: str
    # intermediates for the design record / detailed report
    branch_thickness_in: Optional[float] = None
    chord_thickness_in: Optional[float] = None
    base_metal_area_in2: Optional[float] = None   # col O
    weld_length_in: Optional[float] = None        # col P
    weld_leg_size_in: Optional[float] = None      # col R
    fexx_ksi: Optional[float] = None              # col Q
    fu_governing_ksi: Optional[float] = None      # MIN(col M, col N)
    electrode: str = ""


# --- Core: faithful column O..V evaluation ---------------------------------

def _core(branch_t: float, chord_t: float, fu_branch: float, fu_chord: float,
          width: float, height: float, fexx: float):
    """Return (capacity_lbf, mode, area, weld_len, leg, fu_governing)."""
    # O: branch tube wall cross-sectional area
    area = width * height - (width - 2 * branch_t) * (height - 2 * branch_t)
    weld_len = width + height                                   # P
    leg = weld_leg_size_in(min(branch_t, chord_t))              # R

    fu_governing = min(fu_branch, fu_chord)                     # MIN(M, N)
    base_kips = fu_governing * area                             # S  (unfactored!)
    weld_kips = PHI * 0.6 * fexx * 0.707 * leg * weld_len       # T

    mode = "Base Material Strength" if base_kips < weld_kips else "Weld Shear Strength"  # V
    governing = min(base_kips, weld_kips)                       # U

    return governing * 1000.0, mode, area, weld_len, leg, fu_governing


def capacity_from_template_inputs(branch_material: str, branch_gauge: int,
                                  branch_width: float, branch_height: float,
                                  chord_material: str, chord_gauge: int,
                                  electrode: str = DEFAULT_ELECTRODE) -> WeldCapacity:
    """Faithful spreadsheet path: resolves thicknesses from gauges exactly as the
    template does. Used by the golden test. Raises KeyError on unknown inputs
    (mirrors the #N/A the sheet would produce)."""
    branch_t = _THICKNESS_BY_GAUGE[int(branch_gauge)]
    chord_t = _THICKNESS_BY_GAUGE[int(chord_gauge)]
    fu_branch = _FU_KSI[branch_material]
    fu_chord = _FU_KSI[chord_material]
    fexx = fexx_from_electrode(electrode)
    if fexx is None:
        raise KeyError(f"Unrecognized electrode specification: {electrode!r}")

    cap, mode, area, wlen, leg, fu_gov = _core(
        branch_t, chord_t, fu_branch, fu_chord, branch_width, branch_height, fexx
    )
    return WeldCapacity(
        cap, None, mode, False,
        f"AISC Ch. J via template: {branch_material} ga {branch_gauge} branch to "
        f"{chord_material} ga {chord_gauge} chord, {electrode}, "
        f"{leg:g} in fillet x {wlen:g} in. Governing: {mode}.",
        branch_t, chord_t, area, wlen, leg, fexx, fu_gov, str(electrode),
    )


# --- Live app path: reads a connection-schedule row -------------------------

def weld_capacity(conn: Dict[str, Any]) -> WeldCapacity:
    """Compute capacity from a connection-schedule row.

    Expects the template's branch/chord shape, which Block 1 can populate
    directly from ``Joint.member_roles`` + each member's ``cross_section``:

        branch_material, branch_gauge, branch_width_in, branch_height_in,
        chord_material,  chord_gauge,  fastener_type (electrode)

    Thicknesses may be supplied directly as ``branch_t_in`` / ``chord_t_in`` to
    bypass the gauge lookup. Anything ambiguous routes to needs_review rather
    than guessing.
    """
    notes: list[str] = []
    needs_review = False

    # --- materials (template IFS resolves TUBES only) ---------------------
    branch_material = map_material(conn.get("branch_material") or conn.get("material_grade"))
    chord_material = map_material(conn.get("chord_material") or conn.get("material_grade"))
    if branch_material is None or chord_material is None:
        return WeldCapacity(None, None, "", True,
                            f"Branch/chord material not recognized "
                            f"(branch={conn.get('branch_material')!r}, "
                            f"chord={conn.get('chord_material')!r}). Routed to review.")

    for role, mat in (("branch", branch_material), ("chord", chord_material)):
        if mat not in _TEMPLATE_TUBE_MATERIALS:
            needs_review = True
            notes.append(f"{role} material '{mat}' is outside the welded template's "
                         f"IFS lookup (which resolves only "
                         f"{' and '.join(_TEMPLATE_TUBE_MATERIALS)}); the sheet would "
                         f"return #N/A here. Fu taken from the material table instead")

    fu_branch = _FU_KSI[branch_material]
    fu_chord = _FU_KSI[chord_material]

    # --- thicknesses: explicit override, else gauge lookup ----------------
    branch_t = _num(conn.get("branch_t_in"))
    if branch_t <= 0:
        branch_t = thickness_from_gauge(conn.get("branch_gauge")) or 0.0
    chord_t = _num(conn.get("chord_t_in"))
    if chord_t <= 0:
        chord_t = thickness_from_gauge(conn.get("chord_gauge")) or 0.0

    width = _num(conn.get("branch_width_in"))
    height = _num(conn.get("branch_height_in"))

    if branch_t <= 0 or chord_t <= 0 or width <= 0 or height <= 0:
        return WeldCapacity(None, None, "", True,
                            "Missing/zero geometry input (need branch gauge or "
                            "thickness, chord gauge or thickness, branch width, "
                            "branch height). Routed to review.")

    # Column O goes negative/nonsensical if the wall is not thinner than half the
    # tube. Guard rather than propagate a garbage area.
    if 2 * branch_t >= min(width, height):
        return WeldCapacity(None, None, "", True,
                            f"Branch wall (2 x {branch_t:.3f} in) is not thinner than "
                            f"the branch section ({width:g} x {height:g} in); the base "
                            f"metal area formula is not physical here. Routed to review.")

    # --- electrode --------------------------------------------------------
    raw_electrode = conn.get("fastener_type") or conn.get("electrode")
    fexx = fexx_from_electrode(raw_electrode)
    electrode = str(raw_electrode or "")
    if fexx is None:
        fexx = fexx_from_electrode(DEFAULT_ELECTRODE)
        electrode = DEFAULT_ELECTRODE
        needs_review = True
        notes.append(f"electrode {raw_electrode!r} unrecognized; "
                     f"assumed {DEFAULT_ELECTRODE} (conservative)")

    cap, mode, area, wlen, leg, fu_gov = _core(
        branch_t, chord_t, fu_branch, fu_chord, width, height, fexx
    )

    basis = (f"AISC Ch. J welded template ({branch_material} branch t={branch_t:.3f} in "
             f"to {chord_material} chord t={chord_t:.3f} in, Fu={fu_gov} ksi, {electrode} "
             f"(Fexx={fexx:g} ksi), auto leg {leg:g} in x weld length {wlen:g} in, "
             f"base metal area {area:.4f} in^2, governing: {mode}).")
    if notes:
        basis += " NEEDS REVIEW: " + "; ".join(notes) + "."

    return WeldCapacity(cap, None, mode, needs_review, basis,
                        branch_t, chord_t, area, wlen, leg, fexx, fu_gov, electrode)


def weld_shear_capacity_lbf(conn: Dict[str, Any]):
    """Drop-in for the welded branch of ``estimate_screening_capacity_lbf``.
    Returns ``(capacity_lbf, basis)`` like the bolted twin. A 0.0 capacity
    forces an Incomplete status when the joint needs review."""
    r = weld_capacity(conn)
    return (r.shear_lbf if r.shear_lbf is not None else 0.0), r.basis


# --- Public constants and helpers for the Streamlit UI ----------------------

MATERIALS: list = list(_FU_KSI.keys())
TUBE_MATERIALS: list = list(_TEMPLATE_TUBE_MATERIALS)
ELECTRODES: list = ["E60XX", "E70XX", "E80XX", "E90XX"]
GAUGES: list = sorted(_THICKNESS_BY_GAUGE)


def default_welded_row(joint_id: Any) -> Dict[str, Any]:
    """Default connection-schedule row for a single welded tube-to-tube joint."""
    return {
        "joint_id": joint_id,
        "branch_material": "Carbon Steel Tube",
        "branch_gauge": 12,
        "branch_width_in": 0.0,
        "branch_height_in": 0.0,
        "chord_material": "Carbon Steel Tube",
        "chord_gauge": 12,
        "fastener_type": DEFAULT_ELECTRODE,
        "notes": "",
    }
