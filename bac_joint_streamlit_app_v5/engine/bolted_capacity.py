"""Bolted connection capacity — AISC Chapter J (Block 4).

Pure-Python port of the "Bolted Connections" sheet in
``Capacity Analysis Template.xlsx`` (S. Rosario). Every line below maps to a
column in that sheet so results are traceable back to the validated workbook:

    H Thickness            (gauge -> min thickness lookup)
    I Ultimate strength Fu (material -> Fu lookup, ksi)
    J Bearing deformation  = 2.4 * d * t * Fu
    K Hole diameter h      = d + 1/16 if d < 1 else d
    L Clear distance Lc    = Le - h/2
    M Tear-out             = 1.2 * Lc * t * Fu
    N Bolt bearing         = min(J, M)
    O AISC bolt category   (bolt grade -> category)
    P Nominal shear stress (category lookup, ksi, threads NOT excluded)
    Q Bolt gross area      = pi * d**2 / 4
    R Bolt shear           = P * Q
    S MaxShearStrength     = min(R, N) * n * 0.75   <-- factored group shear
    T Governing mode       = "Bolt Shear" if R < N else "Bolt Bearing"
    U Nominal tensile str. (category lookup, ksi)
    V MaxBoltTension       = 0.75 * U * Q * n       <-- factored, group

Both S and V are group capacities (x n fasteners). Capacities come out of the
template in KIPS; this module returns lbf (x1000) so they line up with the lbf
demand side of the pipeline.

One note inherited verbatim from the template (flagged, not changed):
  * The Fu table values (17-27 ksi) are low for structural steel; reproduced
    faithfully. Confirm with the template author before relying on magnitudes.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Dict, Optional

PHI = 0.75  # resistance factor baked into template cols S and V

# --- Lookup tables (verbatim from "Bolted Connection Data") -----------------

# Gauge -> minimum thickness (in). Carbon/Galvanized share one table;
# Stainless families use the other (sheet A3:B8 vs E3:F8).
_THICKNESS_CARBON_GALV = {8: 0.153, 10: 0.127, 12: 0.095, 14: 0.067, 16: 0.054, 18: 0.044}

# Material -> ultimate strength Fu (ksi), column M/N table.
_FU_KSI = {
    "Galvanized Sheet Steel": 19,
    "Carbon Steel Plate": 21,
    "Carbon Steel Pipe": 20,
    "Carbon Steel Tube": 27,
}
_STAINLESS_MATERIALS = {"Stainless Sheet Steel", "Stainless Steel Plate", "Stainless Steel Pipe"}

# AISC category -> (nominal tensile ksi, nominal shear ksi threads-NOT-excluded),
# table A44:D49. P uses column 3 (shear, not excluded); U uses column 2 (tensile).
_CATEGORY_STRESS = {
    "A307": (45, 27),
    "Group 120": (90, 54),
    "Group 144": (108, 65),
    "Group 150": (113, 68),
    "Group 200": (150, 90),
}

# Canonical bolt designation -> AISC category, column O logic.
_BOLT_CATEGORY = {
    "ASTM A307, Grade A": "A307",
    "SAE J429, Grade 2": "A307",
    "SAE J429, Grade 5": "Group 120",
    "SAE J429, Grade 8": "Group 150",
    "ASTM F593, Type 304/316, Condition A, annealed": "A307",
    "ASTM A193, Grade B8, Class 2, Type 304": "Group 120",
    "ASTM F468, Ni 500": "Group 150",
    "ASTM F3125, Grade A325, Type 1": "Group 120",
}

DEFAULT_BOLT = "ASTM A307, Grade A"  # template/VB default; also the conservative floor


@dataclass
class BoltedCapacity:
    shear_lbf: Optional[float]      # col S * 1000 (factored, group)
    tension_lbf: Optional[float]    # col V * 1000 (factored, per bolt)
    governing_mode: str             # col T
    needs_review: bool
    basis: str
    # intermediate values for the design record / a future detailed report
    thickness_in: Optional[float] = None
    fu_ksi: Optional[float] = None
    category: str = ""


# --- Core: faithful column J..V evaluation ----------------------------------

def _core(thickness_in: float, fu_ksi: float, n_fasteners: float,
          diameter_in: float, le_in: float, category: str):
    tensile_ksi, shear_ksi = _CATEGORY_STRESS[category]
    bearing_def = 2.4 * diameter_in * thickness_in * fu_ksi          # J
    hole = diameter_in + (1.0 / 16.0) if diameter_in < 1 else diameter_in  # K
    lc = le_in - hole / 2.0                                          # L
    tear_out = 1.2 * lc * thickness_in * fu_ksi                      # M
    bearing = min(bearing_def, tear_out)                            # N
    area = math.pi * diameter_in ** 2 / 4.0                          # Q
    bolt_shear = shear_ksi * area                                   # R
    shear_kips = min(bolt_shear, bearing) * n_fasteners * PHI        # S
    mode = "Bolt Shear" if bolt_shear < bearing else "Bolt Bearing"  # T
    tension_kips = PHI * tensile_ksi * area * n_fasteners           # V (group)
    return shear_kips * 1000.0, mode, tension_kips * 1000.0, lc


def capacity_from_template_inputs(material: str, gauge: int, n_fasteners: float,
                                  bolt_used: str, diameter_in: float,
                                  le_in: float) -> BoltedCapacity:
    """Faithful spreadsheet path: resolves thickness from gauge exactly as the
    template does. Used by the golden test. Raises KeyError on unknown inputs
    (mirrors the #N/A the sheet would produce)."""
    table = _THICKNESS_STAINLESS if material in _STAINLESS_MATERIALS else _THICKNESS_CARBON_GALV
    thickness = table[int(gauge)]
    fu = _FU_KSI[material]
    category = _BOLT_CATEGORY[bolt_used]
    shear, mode, tension, _lc = _core(thickness, fu, n_fasteners, diameter_in, le_in, category)
    return BoltedCapacity(shear, tension, mode, False,
                          f"AISC Ch. J via template: {material}, ga {gauge}, "
                          f"{category}, {n_fasteners:g} x {diameter_in:g} in.",
                          thickness, fu, category)


# --- Mapping helpers for the live connection schedule -----------------------

def map_material(raw: Any) -> Optional[str]:
    if not raw:
        return None
    s = str(raw).strip().lower()
    if s in {m.lower() for m in _FU_KSI}:
        return next(m for m in _FU_KSI if m.lower() == s)
    if "glv" in s or "galvan" in s:
        return "Galvanized Sheet Steel"
    if "stainless" in s:
        if "plate" in s:
            return "Stainless Steel Plate"
        if "pipe" in s:
            return "Stainless Steel Pipe"
        return "Stainless Sheet Steel"
    if "carbon" in s or "steel" in s:
        if "tube" in s:
            return "Carbon Steel Tube"
        if "pipe" in s:
            return "Carbon Steel Pipe"
        if "plate" in s:
            return "Carbon Steel Plate"
    return None


def map_bolt(raw: Any) -> Optional[str]:
    if not raw:
        return None
    s = str(raw).strip().lower()
    for canonical in _BOLT_CATEGORY:
        if s == canonical.lower():
            return canonical
    if "a325" in s or "f3125" in s:
        return "ASTM F3125, Grade A325, Type 1"
    if "grade 8" in s or "j429" in s and "8" in s:
        return "SAE J429, Grade 8"
    if "grade 5" in s:
        return "SAE J429, Grade 5"
    if "grade 2" in s:
        return "SAE J429, Grade 2"
    if "f468" in s:
        return "ASTM F468, Ni 500"
    if "b8" in s or "a193" in s:
        return "ASTM A193, Grade B8, Class 2, Type 304"
    if "f593" in s:
        return "ASTM F593, Type 304/316, Condition A, annealed"
    if "a307" in s:
        return "ASTM A307, Grade A"
    return None


def _num(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except Exception:
        return default


# --- Live app path: reads a connection-schedule row -------------------------

def bolted_capacity(conn: Dict[str, Any]) -> BoltedCapacity:
    """Compute capacity from a connection-schedule row. Thickness is taken
    directly from ``sheet_t_in`` (no gauge lookup needed in the live model).
    Anything ambiguous routes to needs_review rather than guessing."""
    notes = []
    needs_review = False

    n = _num(conn.get("n_fasteners"))
    d = _num(conn.get("diameter_in"))
    t = _num(conn.get("sheet_t_in"))
    le = _num(conn.get("edge_dist_in"))

    material = map_material(conn.get("material_grade"))
    if material is None:
        return BoltedCapacity(None, None, "", True,
                              f"Material '{conn.get('material_grade')}' not recognized "
                              f"(expected one of {', '.join(_FU_KSI)}). Routed to review.")
    fu = _FU_KSI[material]

    bolt = map_bolt(conn.get("fastener_type"))
    if bolt is None:
        bolt = DEFAULT_BOLT
        needs_review = True
        notes.append(f"bolt grade '{conn.get('fastener_type')}' unrecognized; "
                     f"assumed {DEFAULT_BOLT} (conservative)")
    category = _BOLT_CATEGORY[bolt]

    if n <= 0 or d <= 0 or t <= 0 or le <= 0:
        return BoltedCapacity(None, None, "", True,
                              "Missing/zero geometry input (need n, diameter, "
                              "thickness, edge distance). Routed to review.")

    shear, mode, tension, lc = _core(t, fu, n, d, le, category)
    if lc <= 0:
        needs_review = True
        notes.append(f"clear distance Lc = {lc:.3f} in <= 0; edge distance too "
                     f"small for this hole — tear-out result is not physical")

    basis = (f"AISC Ch. J ({material}, t={t:.3f} in, Fu={fu} ksi, {category}, "
             f"{n:g} x {d:g} in, governing: {mode}). "
             f"Factored group tension = {tension:,.0f} lbf.")
    if notes:
        basis += " NEEDS REVIEW: " + "; ".join(notes) + "."
    return BoltedCapacity(shear, tension, mode, needs_review, basis, t, fu, category)


def bolted_shear_capacity_lbf(conn: Dict[str, Any]):
    """Drop-in for the bolted branch of estimate_screening_capacity_lbf.
    Returns (capacity_lbf, basis) like the other branches. Capacity is the
    factored group SHEAR (template col S); 0.0 capacity forces an Incomplete
    status when the joint needs review."""
    r = bolted_capacity(conn)
    return (r.shear_lbf if r.shear_lbf is not None else 0.0), r.basis


# --- Public constants and helpers for the Streamlit UI ----------------------

MATERIALS: list = list(_FU_KSI.keys())
BOLTS: list = list(_BOLT_CATEGORY.keys())


def default_bolted_row(joint_id: Any) -> Dict[str, Any]:
    """Return a default connection-schedule row for a single joint."""
    return {
        "joint_id": joint_id,
        "material_grade": MATERIALS[0],
        "sheet_t_in": 0.0,
        "n_fasteners": 1,
        "diameter_in": 0.3125,
        "edge_dist_in": 1.5*0.3125,
        "fastener_type": DEFAULT_BOLT,
        "notes": "",
    }
