"""
Block 1 - Joint & Section Recognition pipeline (UNIT-SAFE)
==========================================================
Single-module extraction of the frozen Block-1 notebook
(Joint_Analysis_Block1_FINAL_FR.ipynb), assembled in strict cell order.

Public surface used by the app:
    run(...)            -> AnalysisResult     (requires Windows + live Inventor)
    AnalysisResult      dataclass with .to_json()
    InventorError       raised when no live Inventor session is reachable
    DEFAULT_TOL_TOUCH_IN, DEFAULT_CLUSTER_CAP_IN, DEFAULT_SUPPORT_Z_TOL_IN

Import is platform-safe: win32 is guarded, so this module imports on any host.
Only run() needs Inventor. Canonical output units: in, in^2, in^4, lbm, lbf, deg.

Source of truth is the notebook; this file is a faithful concatenation, not a
rewrite. Section banners are preserved so it maps back to the thesis sections.
"""

from __future__ import annotations



# ############################################################################
# >>> notebook cell [02]
# ############################################################################
"""
BLOCK 1 - COMPLETE PIPELINE
Single Jupyter Notebook
All code in one place
"""


import math
from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Optional, Sequence
import json

try:
    import win32com.client as _w32
    import pythoncom
    _HAVE_WIN32 = True
except Exception:
    _HAVE_WIN32 = False


# ############################################################################
# >>> notebook cell [04]
# ############################################################################
# ============================================================================
# SECTION 0: UNIT GUARD (unit_guard.py)   [NEW: prevents unit mixing]
# ============================================================================
# Canonical units exported by Block 1 and consumed by Block 2:
#   length: in
#   area: in^2
#   inertia/J: in^4
#   mass: lbm
#   force/weight: lbf
#   angle: deg
#
# Inventor API/database values are treated as:
#   length: cm
#   mass: kg
#   angle: rad, when raw angles are returned by the API
# ============================================================================

CM_TO_IN = 1.0 / 2.54
KG_TO_LBM = 2.20462262185
KG_TO_LBF_WEIGHT = 2.20462262185  # kg mass under standard gravity -> lbf weight
RAD_TO_DEG = 180.0 / math.pi

CANONICAL_UNITS = {
    "unit_policy": "convert_inventor_api_boundary_to_block2_canonical_units",
    "length": "in",
    "area": "in^2",
    "inertia": "in^4",
    "mass": "lbm",
    "force": "lbf",
    "angle": "deg",
    "source_api_length": "cm",
    "source_api_mass": "kg",
    "source_api_angle": "rad",
}

DEFAULT_TOL_TOUCH_IN = 0.20
DEFAULT_SUPPORT_Z_TOL_IN = 1.00
DEFAULT_CLUSTER_CAP_IN = None


def cm_to_in(x: float) -> float:
    return float(x) * CM_TO_IN


def cm2_to_in2(x: float) -> float:
    return float(x) * CM_TO_IN ** 2


def cm4_to_in4(x: float) -> float:
    return float(x) * CM_TO_IN ** 4


def kg_to_lbm(x: float) -> float:
    return float(x) * KG_TO_LBM


def kg_to_lbf_weight(x: float) -> float:
    return float(x) * KG_TO_LBF_WEIGHT


def rad_to_deg(x: float) -> float:
    return float(x) * RAD_TO_DEG


def pt_cm_to_in(p):
    """Convert a 3D point from Inventor API/database cm to canonical inches."""
    return tuple(cm_to_in(v) for v in p)


def loops_cm_to_in(loops):
    """Convert section-profile loops from Inventor API/database cm to inches."""
    return [[(cm_to_in(x), cm_to_in(y)) for x, y in loop] for loop in loops]


def require_positive_length_in(value: float, name: str):
    if value is None or value <= 0:
        raise ValueError(f"{name} must be a positive length in inches. Got {value}")
    return value


def require_reasonable_member_length_in(value: float, name: str = "member length"):
    """
    Catch obvious cm/in/mm mistakes early. Adjust the upper bound if your
    products include members longer than 2000 in.
    """
    if value <= 0 or value > 2000:
        raise ValueError(
            f"{name}={value:.3f} in looks unreasonable. "
            "Possible unit conversion error."
        )
    return value


def assert_block2_canonical_units(result=None):
    """
    Lightweight audit hook. Use after run(result) if you want to fail fast
    before handing JSON to Block 2.
    """
    expected = {
        "length": "in",
        "area": "in^2",
        "inertia": "in^4",
        "mass": "lbm",
        "force": "lbf",
        "angle": "deg",
    }
    units = CANONICAL_UNITS if result is None else result.units
    for k, v in expected.items():
        if units.get(k) != v:
            raise ValueError(f"Unit audit failed: expected {k}={v}, got {units.get(k)!r}")
    return True


def unit_guard_smoke_test():
    assert abs(cm_to_in(2.54) - 1.0) < 1e-9
    assert abs(cm2_to_in2(2.54 ** 2) - 1.0) < 1e-9
    assert abs(cm4_to_in4(2.54 ** 4) - 1.0) < 1e-9
    assert abs(kg_to_lbm(1.0) - 2.20462262185) < 1e-9
    assert abs(kg_to_lbf_weight(1.0) - 2.20462262185) < 1e-9
    assert abs(rad_to_deg(math.pi) - 180.0) < 1e-9
    assert assert_block2_canonical_units() is True
    print("✓ Unit guard smoke test passed: canonical units are in, lbm, lbf, deg")


unit_guard_smoke_test()


# ############################################################################
# >>> notebook cell [06]
# ############################################################################
# ============================================================================
# SECTION 1: DATA MODEL (data_model.py)   [MODIFIED for Block 2 hand-off]
# ============================================================================

class SectionType(str, Enum):
    """Cross-section families per BAC Section 1.1 nomenclature."""
    C_UNLIPPED = "c_unlipped"
    C_LIPPED = "c_lipped"
    H_UNLIPPED = "h_unlipped"
    H_LIPPED = "h_lipped"
    L_UNLIPPED = "l_unlipped"
    L_LIPPED = "l_lipped"
    Z_UNLIPPED = "z_unlipped"
    Z_LIPPED = "z_lipped"
    ROUND_BAR = "round_bar"
    ROUND_HSS = "round_hss"
    RECT_HSS = "rect_hss"
    UNKNOWN = "unknown"


class SectionFamily(str, Enum):
    OPEN = "open"
    HOLLOW = "hollow"


HOLLOW_TYPES = {SectionType.RECT_HSS, SectionType.ROUND_HSS}


class JointType(str, Enum):
    """Connection configuration taxonomy."""
    CORNER = "corner"
    TEE_CONNECTION = "tee_connection"
    SPLICE = "splice"
    GUSSET = "gusset"
    CROSSING = "crossing"
    UNKNOWN = "unknown"


@dataclass
class CrossSection:
    """Profile classification and geometry."""
    section_type: SectionType = SectionType.UNKNOWN
    family: SectionFamily = SectionFamily.OPEN
    gauge: Optional[int] = None
    depth: Optional[float] = None
    width: Optional[float] = None
    wall_thickness: Optional[float] = None
    is_lipped: Optional[bool] = None
    n_loops: Optional[int] = None
    occupancy_signature: Optional[str] = None
    detection_method: str = "geometry"
    confidence: float = 0.0
    # --- solver-ready section properties (canonical Block-2 units) ---
    A: Optional[float] = None          # area, in^2
    Iy: Optional[float] = None         # 2nd moment about local y, in^4
    Iz: Optional[float] = None         # 2nd moment about local z, in^4
    Iyz: Optional[float] = None        # product of inertia, in^4
    J: Optional[float] = None          # torsion constant, in^4
    props_method: str = ""             # how A/Iy/Iz/J were obtained
    needs_review: bool = False         # True => properties approximate / missing
    review_reason: str = ""
    length_unit: str = "in"
    area_unit: str = "in^2"
    inertia_unit: str = "in^4"


@dataclass
class Member:
    occurrence_name: str
    bom_description: str = ""
    part_number: str = ""
    material: str = ""
    is_glv_or_hdg: bool = False
    start_point: tuple[float, float, float] = (0.0, 0.0, 0.0)
    end_point: tuple[float, float, float] = (0.0, 0.0, 0.0)
    length: float = 0.0
    cross_section: CrossSection = field(default_factory=CrossSection)
    # --- NEW ---
    occurrence_path: str = ""          # UNIQUE per-instance id (full occurrence path)
    dry_mass: Optional[float] = None       # dry mass, lbm
    mass_unit: str = "lbm"
    self_weight_lbf: Optional[float] = None # force from dry mass under standard gravity, lbf
    force_unit: str = "lbf"
    length_unit: str = "in"


@dataclass
class Joint:
    """A joint node where members meet."""
    joint_id: str
    location: tuple[float, float, float]
    member_names: list[str] = field(default_factory=list)
    member_roles: dict[str, str] = field(default_factory=dict)
    angles_deg: list[float] = field(default_factory=list)
    gap: Optional[float] = None
    joint_type: JointType = JointType.UNKNOWN
    taxonomy: str = "open"
    is_inferred: bool = False
    confidence: float = 0.0
    # --- NEW ---
    geom_descriptor: str = ""          # HSS geometry T/Y/K/X/KT (hollow only)
    is_support_candidate: bool = False # base/ground node candidate (confirm!)
    needs_review: bool = False
    review_reason: str = ""


@dataclass
class AnalysisResult:
    source_document: str = ""
    units: dict = field(default_factory=dict)   # NEW: section-0 unit self-check
    members: list[Member] = field(default_factory=list)
    joints: list[Joint] = field(default_factory=list)

    def to_json(self, indent: int = 2) -> str:
        def encode(o):
            if isinstance(o, Enum):
                return o.value
            raise TypeError(f"not serializable: {type(o)}")
        return json.dumps(asdict(self), indent=indent, default=encode)

print("\u2713 Data model loaded (Block 2 fields added, canonical units added)")


# ############################################################################
# >>> notebook cell [08]
# ############################################################################
Loop = Sequence[tuple[float, float]]

_CANONICAL = {
    # OPEN sections only. Hollow sections are caught earlier.
    SectionType.C_UNLIPPED: ("111", "100", "111"),
    SectionType.C_LIPPED:   ("111", "110", "111"),
    SectionType.H_UNLIPPED: ("111", "010", "010"),
    SectionType.H_LIPPED:   ("111", "110", "110"),
    SectionType.L_UNLIPPED: ("100", "100", "111"),
    SectionType.L_LIPPED:   ("100", "110", "111"),
    SectionType.Z_UNLIPPED: ("110", "010", "011"),
    SectionType.Z_LIPPED:   ("110", "110", "011"),
}

# Approximate sharp-corner baselines. A lipped version adds about two extra
# return-leg corners. This is used only as a lip refinement, not as the main
# C/H/L/Z classifier.
_LIP_BASE_CORNERS = {
    SectionType.C_UNLIPPED: 8,
    SectionType.H_UNLIPPED: 8,
    SectionType.L_UNLIPPED: 6,
    SectionType.Z_UNLIPPED: 8,
}

_LIPPED_OPEN_TYPES = {
    SectionType.C_LIPPED,
    SectionType.H_LIPPED,
    SectionType.L_LIPPED,
    SectionType.Z_LIPPED,
}

_UNLIPPED_OPEN_TYPES = {
    SectionType.C_UNLIPPED,
    SectionType.H_UNLIPPED,
    SectionType.L_UNLIPPED,
    SectionType.Z_UNLIPPED,
}


def _bbox(points: Loop) -> tuple[float, float, float, float]:
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def _point_in_loop(x: float, y: float, loop: Loop) -> bool:
    inside = False
    n = len(loop)
    j = n - 1
    for i in range(n):
        xi, yi = loop[i]
        xj, yj = loop[j]
        if ((yi > y) != (yj > y)) and (
            x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi
        ):
            inside = not inside
        j = i
    return inside


def _point_in_region(x: float, y: float, outer: Loop, holes: list[Loop]) -> bool:
    if not _point_in_loop(x, y, outer):
        return False
    return not any(_point_in_loop(x, y, h) for h in holes)


def _rotate(points: Loop, ang: float) -> list[tuple[float, float]]:
    c, s = math.cos(ang), math.sin(ang)
    return [(p[0] * c + p[1] * s, -p[0] * s + p[1] * c) for p in points]


def _min_area_angle(outer: Loop) -> float:
    best_ang, best_area = 0.0, float("inf")
    n = len(outer)
    for i in range(n):
        x1, y1 = outer[i]
        x2, y2 = outer[(i + 1) % n]
        dx, dy = x2 - x1, y2 - y1
        if math.hypot(dx, dy) < 1e-9:
            continue
        ang = math.atan2(dy, dx)
        minx, miny, maxx, maxy = _bbox(_rotate(outer, ang))
        area = (maxx - minx) * (maxy - miny)
        if area < best_area:
            best_area, best_ang = area, ang
    return best_ang


def _fit_circle(loop: Loop) -> tuple[float, float, float, float]:
    n = len(loop)
    cx = sum(p[0] for p in loop) / n
    cy = sum(p[1] for p in loop) / n
    r = sum(math.hypot(p[0] - cx, p[1] - cy) for p in loop) / n
    if r < 1e-9:
        return cx, cy, r, 1.0
    err = math.sqrt(sum((math.hypot(p[0] - cx, p[1] - cy) - r) ** 2 for p in loop) / n)
    return cx, cy, r, err / r


def _polygon_area(loop: Loop) -> float:
    a = 0.0
    n = len(loop)
    for i in range(n):
        x1, y1 = loop[i]
        x2, y2 = loop[(i + 1) % n]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def _loop_moments(loop):
    """One closed loop [(x,y),...] -> (A, Cx, Cy, Ixx, Iyy, Ixy), centroidal, as a positive solid."""
    n = len(loop)
    A = Cx = Cy = Ixx = Iyy = Ixy = 0.0
    for i in range(n):
        x0, y0 = loop[i]
        x1, y1 = loop[(i + 1) % n]
        cr = x0 * y1 - x1 * y0
        A   += cr
        Cx  += (x0 + x1) * cr
        Cy  += (y0 + y1) * cr
        Ixx += (y0 * y0 + y0 * y1 + y1 * y1) * cr
        Iyy += (x0 * x0 + x0 * x1 + x1 * x1) * cr
        Ixy += (x0 * y1 + 2 * x0 * y0 + 2 * x1 * y1 + x1 * y0) * cr
    A *= 0.5
    if abs(A) < 1e-12:
        return (0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    Cx /= (6 * A); Cy /= (6 * A)
    Ixx = Ixx / 12.0 - A * Cy * Cy      # centroidal
    Iyy = Iyy / 12.0 - A * Cx * Cx
    Ixy = Ixy / 24.0 - A * Cx * Cy
    return (abs(A), Cx, Cy, abs(Ixx), abs(Iyy), Ixy)


def section_polygon_props(outer, holes=()):
    """Exact area, centroid, and centroidal 2nd moments of (outer minus holes),
    in the loop's own x-y frame. Returns (A, Cx, Cy, Ixx, Iyy, Ixy)."""
    parts = [(+1.0,) + _loop_moments(outer)] + [(-1.0,) + _loop_moments(h) for h in holes]
    A = sum(s * a for s, a, *_ in parts)
    if abs(A) < 1e-12:
        return (0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    Cx = sum(s * a * cx for s, a, cx, *_ in parts) / A
    Cy = sum(s * a * cy for s, a, _, cy, *_ in parts) / A
    Ixx = Iyy = Ixy = 0.0
    for s, a, cx, cy, ixx, iyy, ixy in parts:
        Ixx += s * (ixx + a * (cy - Cy) ** 2)
        Iyy += s * (iyy + a * (cx - Cx) ** 2)
        Ixy += s * (ixy + a * (cx - Cx) * (cy - Cy))
    return (A, Cx, Cy, Ixx, Iyy, Ixy)


def _occupancy(outer: Loop, holes: list[Loop]) -> tuple[tuple[str, str, str], float]:
    """Coarse 3x3 occupancy grid used only for base C/H/L/Z matching."""
    ang = _min_area_angle(outer)
    ro = _rotate(outer, ang)
    rh = [_rotate(h, ang) for h in holes]
    minx, miny, maxx, maxy = _bbox(ro)
    w = (maxx - minx) / 3.0
    h = (maxy - miny) / 3.0
    sub = 7
    rows = []
    total_probe = 0
    filled_probe = 0

    for r in range(3):
        row = ""
        for c in range(3):
            cx0 = minx + c * w
            cy0 = miny + (2 - r) * h
            hit = False
            for i in range(sub):
                for jj in range(sub):
                    x = cx0 + (i + 0.5) / sub * w
                    y = cy0 + (jj + 0.5) / sub * h
                    total_probe += 1
                    if _point_in_region(x, y, ro, rh):
                        filled_probe += 1
                        hit = True
            row += "1" if hit else "0"
        rows.append(row)

    fill_ratio = filled_probe / total_probe if total_probe else 0.0
    return (rows[0], rows[1], rows[2]), fill_ratio


def _dihedral_orbit(sig: tuple[str, str, str]) -> set[tuple[str, str, str]]:
    """All rotations/reflections of a 3x3 signature; makes matching orientation-safe."""
    def as_grid(s):
        return [[int(s[r][c]) for c in range(3)] for r in range(3)]

    def to_sig(g):
        return tuple("".join(str(g[r][c]) for c in range(3)) for r in range(3))

    def rot90(g):
        return [[g[2 - c][r] for c in range(3)] for r in range(3)]

    def flip(g):
        return [[g[r][2 - c] for c in range(3)] for r in range(3)]

    g = as_grid(sig)
    out = set()
    for _ in range(4):
        out.add(to_sig(g))
        out.add(to_sig(flip(g)))
        g = rot90(g)
    return out


def _match_signature(sig: tuple[str, str, str], min_approx_conf: float = 0.50) -> tuple[SectionType, float]:
    """
    Match a 3x3 signature to canonical open-section templates.

    Key fix: weak approximate matches return UNKNOWN instead of forcing the nearest
    section family. This prevents C/H/hat/Z mix-ups from coarse midspan slices.
    """
    orbit = _dihedral_orbit(sig)

    for stype, canon in _CANONICAL.items():
        if canon in orbit:
            return stype, 0.90

    best, best_d = SectionType.UNKNOWN, 99
    flat_obs = ["".join(s) for s in orbit]

    for stype, canon in _CANONICAL.items():
        canon_flat = "".join(canon)
        d = min(sum(a != b for a, b in zip(obs, canon_flat)) for obs in flat_obs)
        if d < best_d:
            best, best_d = stype, d

    conf = round(max(0.0, 1.0 - best_d / 9.0) * 0.60, 2)
    if conf < min_approx_conf:
        return SectionType.UNKNOWN, conf

    return best, conf


def _is_ring_signature(sig: tuple[str, str, str]) -> bool:
    """Rectangular hollow fallback: empty center with all border cells filled."""
    center = sig[1][1]
    border = sig[0] + sig[1][0] + sig[1][2] + sig[2]
    return center == "0" and border.count("1") == 8


def _clean_loop(loop: Loop) -> list[tuple[float, float]]:
    pts = list(loop)
    if not pts:
        return []

    out = [pts[0]]
    for p in pts[1:]:
        if math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > 1e-6:
            out.append(p)

    if len(out) > 1 and math.hypot(out[0][0] - out[-1][0], out[0][1] - out[-1][1]) < 1e-6:
        out.pop()

    return out


def _remove_nearly_collinear(loop: Loop, angle_tol_deg: float = 8.0) -> list[tuple[float, float]]:
    """Remove tiny intermediate points on nearly straight runs."""
    pts = _clean_loop(loop)
    if len(pts) < 4:
        return pts

    changed = True
    while changed and len(pts) >= 4:
        changed = False
        keep = []
        n = len(pts)
        for i in range(n):
            a = pts[(i - 1) % n]
            b = pts[i]
            c = pts[(i + 1) % n]
            v1 = (b[0] - a[0], b[1] - a[1])
            v2 = (c[0] - b[0], c[1] - b[1])
            n1 = math.hypot(*v1)
            n2 = math.hypot(*v2)
            if n1 < 1e-9 or n2 < 1e-9:
                changed = True
                continue
            cosang = max(-1.0, min(1.0, (v1[0] * v2[0] + v1[1] * v2[1]) / (n1 * n2)))
            angle = math.degrees(math.acos(cosang))
            if angle < angle_tol_deg:
                changed = True
                continue
            keep.append(b)
        if len(keep) >= 3:
            pts = keep
        else:
            break

    return pts


def _sharp_corner_count(loop: Loop, angle_threshold_deg: float = 45.0) -> int:
    pts = _remove_nearly_collinear(loop)
    n = len(pts)
    if n < 3:
        return 0

    sharp = 0
    for i in range(n):
        a = pts[(i - 1) % n]
        b = pts[i]
        c = pts[(i + 1) % n]
        v1 = (b[0] - a[0], b[1] - a[1])
        v2 = (c[0] - b[0], c[1] - b[1])
        n1 = math.hypot(*v1)
        n2 = math.hypot(*v2)
        if n1 < 1e-9 or n2 < 1e-9:
            continue
        cosang = max(-1.0, min(1.0, (v1[0] * v2[0] + v1[1] * v2[1]) / (n1 * n2)))
        angle = math.degrees(math.acos(cosang))
        if angle > angle_threshold_deg:
            sharp += 1
    return sharp


def _detect_lips(outer: Loop, stype: SectionType) -> bool:
    """
    Lip detector only. It does not decide whether the base family is C/H/L/Z;
    it only upgrades an already detected unlipped open section to lipped.
    """
    base = _LIP_BASE_CORNERS.get(stype)
    if base is None:
        return False

    pts = _remove_nearly_collinear(outer)
    if len(pts) < 6:
        return False

    sharp = _sharp_corner_count(outer)

    # Normal case: two return lips add about two additional sharp turns.
    if sharp >= base + 2:
        return True

    # C-channel-specific fallback: lipped channels often have 10+ simplified
    # vertices even when fillets weaken corner counting.
    if stype == SectionType.C_UNLIPPED and len(pts) >= 10 and sharp >= base:
        return True

    return False


def _perimeter_estimate(stype: SectionType, depth: float, width: float) -> float:
    if stype in {SectionType.ROUND_BAR, SectionType.ROUND_HSS}:
        return math.pi * depth
    if stype in {
        SectionType.C_UNLIPPED, SectionType.C_LIPPED,
        SectionType.H_UNLIPPED, SectionType.H_LIPPED,
        SectionType.Z_UNLIPPED, SectionType.Z_LIPPED,
    }:
        return depth + 2 * width
    if stype in {SectionType.L_UNLIPPED, SectionType.L_LIPPED}:
        return depth + width
    return 2 * (depth + width)


def _structural_hollow_holes(outer: Loop, holes: list[Loop], bb_w: float, bb_h: float) -> list[Loop]:
    """
    Keep only holes that look like a structural HSS/tube inner void.

    Small bolt holes/punched holes should not turn an open section into rect_hss.
    """
    if not holes or bb_w <= 1e-9 or bb_h <= 1e-9:
        return []

    outer_area = max(_polygon_area(outer), 1e-9)
    ox0, oy0, ox1, oy1 = _bbox(outer)
    ocx, ocy = 0.5 * (ox0 + ox1), 0.5 * (oy0 + oy1)

    good = []
    for h in holes:
        hx0, hy0, hx1, hy1 = _bbox(h)
        hw, hh = hx1 - hx0, hy1 - hy0
        hcx, hcy = 0.5 * (hx0 + hx1), 0.5 * (hy0 + hy1)
        h_area = _polygon_area(h)

        area_ratio = h_area / outer_area
        large_enough = area_ratio >= 0.12 and hw >= 0.25 * bb_w and hh >= 0.25 * bb_h
        centered = abs(hcx - ocx) <= 0.25 * bb_w and abs(hcy - ocy) <= 0.25 * bb_h

        if large_enough and centered:
            good.append(h)

    good.sort(key=_polygon_area, reverse=True)
    return good


def _rhs_wall_thickness(outer: Loop, holes: list[Loop], bb_w: float, bb_h: float) -> Optional[float]:
    """RHS wall from outer-vs-inner bbox offset, else area/perimeter fallback."""
    if holes:
        ix0, iy0, ix1, iy1 = _bbox(holes[0])
        iw, ih = ix1 - ix0, iy1 - iy0
        return round(((bb_w - iw) + (bb_h - ih)) / 4.0, 3)
    perim = 2 * (bb_w + bb_h)
    return round(_polygon_area(outer) / perim, 3) if perim > 0 else None


def classify_section(loops: list[Loop], detection_method: str = "extrude_profile") -> CrossSection:
    if not loops or len(loops[0]) < 3:
        return CrossSection(detection_method=detection_method, confidence=0.0)

    outer = loops[0]
    raw_holes = [lp for lp in loops[1:] if len(lp) >= 3]
    n_loops = 1 + len(raw_holes)

    minx, miny, maxx, maxy = _bbox(outer)
    bb_w, bb_h = maxx - minx, maxy - miny
    depth, width = max(bb_w, bb_h), min(bb_w, bb_h)

    # Hole filtering is separate from open-section family matching.
    holes = _structural_hollow_holes(outer, raw_holes, bb_w, bb_h)

    cs = CrossSection(
        detection_method=detection_method,
        n_loops=n_loops,
        depth=depth,
        width=width,
    )

    # 1) Round bar / round HSS.
    _, _, r_out, err_out = _fit_circle(outer)
    area_out = _polygon_area(outer)
    disc_ratio = area_out / (math.pi * r_out * r_out) if r_out > 1e-9 else 0.0

    if err_out < 0.04 and 0.85 <= disc_ratio <= 1.15:
        if holes:
            _, _, r_in, err_in = _fit_circle(holes[0])
            if err_in < 0.06:
                cs.section_type = SectionType.ROUND_HSS
                cs.family = SectionFamily.HOLLOW
                cs.confidence = 0.90
                cs.wall_thickness = round(r_out - r_in, 3)
                cs.is_lipped = False
                return cs

        cs.section_type = SectionType.ROUND_BAR
        cs.family = SectionFamily.OPEN
        cs.confidence = 0.88
        cs.is_lipped = False
        return cs

    # 2) Occupancy signature. Small/raw holes are ignored here unless they are a
    # structural inner void, preventing bolt holes from dominating the section type.
    sig, fill = _occupancy(outer, holes)
    cs.occupancy_signature = "/".join(sig)

    # 3) Rectangular HSS / box.
    if holes or _is_ring_signature(sig):
        cs.section_type = SectionType.RECT_HSS
        cs.family = SectionFamily.HOLLOW
        cs.confidence = 0.85 if holes else 0.70
        cs.is_lipped = False
        cs.wall_thickness = _rhs_wall_thickness(outer, holes, bb_w, bb_h)
        return cs

    # 4) Open sections via safe base matcher.
    stype, conf = _match_signature(sig)
    cs.section_type = stype
    cs.family = SectionFamily.OPEN
    cs.confidence = round(conf, 2)

    # 5) Lip refinement after base family is known.
    if stype in _UNLIPPED_OPEN_TYPES:
        cs.is_lipped = _detect_lips(outer, stype)
        if cs.is_lipped:
            cs.section_type = {
                SectionType.C_UNLIPPED: SectionType.C_LIPPED,
                SectionType.H_UNLIPPED: SectionType.H_LIPPED,
                SectionType.L_UNLIPPED: SectionType.L_LIPPED,
                SectionType.Z_UNLIPPED: SectionType.Z_LIPPED,
            }[stype]
            cs.confidence = max(cs.confidence, 0.78)
    elif stype in _LIPPED_OPEN_TYPES:
        cs.is_lipped = True
    else:
        cs.is_lipped = None

    # 6) Approximate wall thickness for open sections.
    if fill < 0.85 and depth > 0:
        area = _polygon_area(outer) - sum(_polygon_area(h) for h in holes)
        perim = _perimeter_estimate(cs.section_type, depth, width)
        if perim > 0:
            cs.wall_thickness = round(area / perim, 3)

    return cs


print("✓ Optimized section classifier loaded")


# ############################################################################
# >>> notebook cell [09]
# ############################################################################
# ============================================================================
# SECTION 7b: SECTION PROPERTIES (section_props.py)   [NEW for Block 2]
# Solver-ready A / Iy / Iz / J in canonical Block-2 units (in^2, in^4).
# Hollow (HSS/bar) -> closed form (authoritative). Open cold-formed -> thin-wall
# estimate, flagged needs_review (replace with company-manual catalog values).
# ============================================================================

def _segment_props(segments):
    """segments: (length, thickness, cx, cy, orient) with orient in {'v','h'}.
    Returns (A, Iy, Iz) where Iz is depth-axis bending, Iy is width-axis bending."""
    A = sum(L * t for L, t, _, _, _ in segments)
    if A <= 0:
        return 0.0, 0.0, 0.0
    xbar = sum(L * t * cx for L, t, cx, cy, o in segments) / A
    ybar = sum(L * t * cy for L, t, cx, cy, o in segments) / A
    Ixx = Iyy = Ixy = 0.0
    for L, t, cx, cy, o in segments:
        a = L * t
        if o == 'v':
            own_xx = t * L ** 3 / 12.0; own_yy = a * t * t / 12.0
        else:
            own_xx = a * t * t / 12.0;  own_yy = t * L ** 3 / 12.0
        Ixx += own_xx + a * (cy - ybar) ** 2
        Iyy += own_yy + a * (cx - xbar) ** 2
        Ixy += a * (cx - xbar) * (cy - ybar)   # own term is 0 for vertical/horizontal segments
    return A, Iyy, Ixx, Ixy


def _open_J(segments):
    return sum(L * t ** 3 for L, t, _, _, _ in segments) / 3.0   # open thin-wall St-Venant


def _segments_for(cs: CrossSection):
    st = cs.section_type
    d = cs.depth or 0.0; w = cs.width or 0.0; t = cs.wall_thickness or 0.0
    if t <= 0 or d <= 0:
        return None
    if st in (SectionType.C_UNLIPPED, SectionType.C_LIPPED):
        return [(d, t, 0.0, d / 2, 'v'), (w, t, w / 2, 0.0, 'h'), (w, t, w / 2, d, 'h')]
    if st in (SectionType.Z_UNLIPPED, SectionType.Z_LIPPED):
        return [(d, t, 0.0, d / 2, 'v'), (w, t, -w / 2, 0.0, 'h'), (w, t, w / 2, d, 'h')]
    if st in (SectionType.L_UNLIPPED, SectionType.L_LIPPED):
        return [(d, t, 0.0, d / 2, 'v'), (w, t, w / 2, 0.0, 'h')]
    if st in (SectionType.H_UNLIPPED, SectionType.H_LIPPED):
        return [(w, t, w / 2, d, 'h'), (d, t, 0.0, d / 2, 'v'), (d, t, w, d / 2, 'v'),
                (w / 2, t, -w / 4, 0.0, 'h'), (w / 2, t, w + w / 4, 0.0, 'h')]
    return None


def compute_section_properties(cs: CrossSection) -> CrossSection:
    """Populate cs.A/Iy/Iz/J in canonical Block-2 units: in^2 and in^4. Mutates and returns cs."""
    st = cs.section_type
    d, w, t = cs.depth, cs.width, cs.wall_thickness

    if st == SectionType.ROUND_HSS and d and t:
        D = d; di = D - 2 * t
        cs.A = math.pi / 4 * (D ** 2 - di ** 2)
        cs.Iy = cs.Iz = math.pi / 64 * (D ** 4 - di ** 4)
        cs.J = math.pi / 32 * (D ** 4 - di ** 4)
        cs.props_method = "closed_form_round_hss"
        return cs
    if st == SectionType.ROUND_BAR and d:
        D = d
        cs.A = math.pi / 4 * D ** 2
        cs.Iy = cs.Iz = math.pi / 64 * D ** 4
        cs.J = math.pi / 32 * D ** 4
        cs.props_method = "closed_form_round_bar"
        return cs
    if st == SectionType.RECT_HSS and d and w and t:
        h, b = d, w
        cs.A = b * h - (b - 2 * t) * (h - 2 * t)
        cs.Iz = (b * h ** 3 - (b - 2 * t) * (h - 2 * t) ** 3) / 12.0
        cs.Iy = (h * b ** 3 - (h - 2 * t) * (b - 2 * t) ** 3) / 12.0
        Am = (b - t) * (h - t); pm = 2 * ((b - t) + (h - t))
        cs.J = 4 * Am * Am * t / pm if pm > 0 else None
        cs.props_method = "closed_form_rect_hss"
        return cs

    segs = _segments_for(cs)
    if segs:
        A, Iy, Iz, Iyz = _segment_props(segs)
        cs.A, cs.Iy, cs.Iz, cs.J = A, Iy, Iz, _open_J(segs)
        cs.Iyz = Iyz
        cs.props_method = "thinwall_estimate"
        cs.needs_review = True
        cs.review_reason = "open_section_thinwall_estimate_verify_vs_manual_lips_radii_ignored"
        return cs

    cs.needs_review = True
    cs.review_reason = "missing_section_type_or_dimensions"
    return cs

print("OK: Section property calculator loaded (canonical in^2/in^4)")


# ############################################################################
# >>> notebook cell [11]
# ############################################################################
from dataclasses import dataclass, field

@dataclass
class _Contact:
    member_index: int
    at_end: bool
    direction: tuple[float, float, float]
    point: tuple[float, float, float] = (0.0, 0.0, 0.0)   # NEW: contact point (for convergence test)


@dataclass
class JointCandidate:
    location: tuple[float, float, float]
    contacts: list[_Contact] = field(default_factory=list)
    is_inferred: bool = False


def _sub(a, b): return (a[0] - b[0], a[1] - b[1], a[2] - b[2])
def _add(a, b): return (a[0] + b[0], a[1] + b[1], a[2] + b[2])
def _scale(a, s): return (a[0] * s, a[1] * s, a[2] * s)
def _dot(a, b): return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]
def _norm(a): return math.sqrt(_dot(a, a))


def _unit(a):
    n = _norm(a)
    return (a[0] / n, a[1] / n, a[2] / n) if n > 1e-12 else (0.0, 0.0, 0.0)


def _closest_params(p1, p2, q1, q2):
    d1 = _sub(p2, p1)
    d2 = _sub(q2, q1)
    r = _sub(p1, q1)
    a = _dot(d1, d1)
    e = _dot(d2, d2)
    f = _dot(d2, r)
    if a < 1e-12 and e < 1e-12:
        return _norm(r), 0.0, 0.0
    if a < 1e-12:
        s = 0.0
        t = min(max(f / e, 0.0), 1.0)
    else:
        c = _dot(d1, r)
        if e < 1e-12:
            t = 0.0
            s = min(max(-c / a, 0.0), 1.0)
        else:
            b = _dot(d1, d2)
            denom = a * e - b * b
            s = min(max((b * f - c * e) / denom, 0.0), 1.0) if denom > 1e-12 else 0.0
            t = (b * s + f) / e
            if t < 0.0:
                t = 0.0
                s = min(max(-c / a, 0.0), 1.0)
            elif t > 1.0:
                t = 1.0
                s = min(max((b - c) / a, 0.0), 1.0)
    cp = _add(p1, _scale(d1, s))
    cq = _add(q1, _scale(d2, t))
    return _norm(_sub(cp, cq)), s, t


def _away_direction(member: Member, param: float):
    s = member.start_point
    e = member.end_point
    contact = _add(s, _scale(_sub(e, s), param))
    far = e if param < 0.5 else s
    return _unit(_sub(far, contact))


def _member_size(m: Member) -> float:
    """Largest cross-section dimension; drives the per-pair snap tolerance."""
    cs = m.cross_section
    return max(cs.depth or 0.0, cs.width or 0.0)


def detect_joints(
    members: list[Member],
    tol_touch: float = 0.2,
    end_frac: float = 0.05,
    reach_factor: float = 0.7,
    cluster_factor: float = 1.2,
    min_reach: float = 2.0,
    min_cluster: float = 1.0,
    cluster_cap: Optional[float] = None,   # NEW: absolute cap on cluster tolerance
) -> list[JointCandidate]:
    """Find joint candidates among member centrelines.

    Tolerances scale with section size: the cross-section component models offset
    centrelines by ~half a section width, so a fixed 2 mm tolerance missed every
    joint. reach = touch + reach_factor*(size_i + size_j); cluster_tol = cluster_factor
    * median section size. inferred = closest approach exceeds physical touching.
    """
    raw = []
    for i in range(len(members)):
        for j in range(i + 1, len(members)):
            mi, mj = members[i], members[j]
            dist, s, t = _closest_params(
                mi.start_point, mi.end_point, mj.start_point, mj.end_point)
            reach = max(min_reach,
                        tol_touch + reach_factor * (_member_size(mi) + _member_size(mj)))
            if dist > reach:
                continue
            inferred = dist > tol_touch
            pi = _add(mi.start_point, _scale(_sub(mi.end_point, mi.start_point), s))
            pj = _add(mj.start_point, _scale(_sub(mj.end_point, mj.start_point), t))
            loc = _scale(_add(pi, pj), 0.5)
            ci = _Contact(i, (s <= end_frac or s >= 1 - end_frac), _away_direction(mi, s), pi)
            cj = _Contact(j, (t <= end_frac or t >= 1 - end_frac), _away_direction(mj, t), pj)
            raw.append((loc, ci, cj, inferred))

    sizes = [_member_size(m) for m in members if _member_size(m) > 0]
    char = sorted(sizes)[len(sizes) // 2] if sizes else 0.0
    cluster_tol = max(min_cluster, cluster_factor * char)
    if cluster_cap is not None:
        cluster_tol = min(cluster_tol, cluster_cap)

    joints: list[JointCandidate] = []
    for loc, ci, cj, inferred in raw:
        host = None
        for jc in joints:
            if _norm(_sub(jc.location, loc)) <= cluster_tol:
                host = jc
                break
        if host is None:
            host = JointCandidate(location=loc)
            joints.append(host)
        host.is_inferred = host.is_inferred or inferred
        for c in (ci, cj):
            if not any(existing.member_index == c.member_index for existing in host.contacts):
                host.contacts.append(c)
    return joints

print("✓ Joint detector loaded")


# ############################################################################
# >>> notebook cell [13]
# ############################################################################
def _angle_deg(a, b) -> float:
    d = max(-1.0, min(1.0, _dot(a, b)))
    return math.degrees(math.acos(d))


def _pairwise_angles(dirs) -> list[float]:
    out = []
    for i in range(len(dirs)):
        for j in range(i + 1, len(dirs)):
            out.append(_angle_deg(dirs[i], dirs[j]))
    return out


def _name_from_contact(members: list[Member], contact: _Contact) -> str:
    m = members[contact.member_index]
    return m.occurrence_path or m.occurrence_name   # MODIFIED: prefer UNIQUE path


_T_ANGLE_TOL = 20.0   # +/- deg around 90 still reads as a T; otherwise Y


def _hollow_2member_descriptor(angle: float) -> str:
    return "T" if abs(angle - 90.0) <= _T_ANGLE_TOL else "Y"


def _hollow_nbrace_descriptor(members, contacts):
    """K / X / KT for a hollow node with one through-chord + braces.
    Returns (descriptor, ok). ok=False => no clean chord, caller flags review."""
    through = [c for c in contacts if not c.at_end]
    ending = [c for c in contacts if c.at_end]
    if len(through) != 1 or len(ending) < 2:
        return "", False
    chord = _unit(through[0].direction)
    sides, axiality = [], []
    for c in ending:
        v = _unit(c.direction)
        para = _scale(chord, _dot(v, chord))
        sides.append(_sub(v, para))            # component perpendicular to chord
        axiality.append(abs(_dot(v, chord)))   # ~0 => brace perpendicular (KT signal)
    ref = _unit(sides[0]) if _norm(sides[0]) > 1e-9 else (0.0, 0.0, 0.0)
    signs = [_dot(_unit(s), ref) if _norm(s) > 1e-9 else 0.0 for s in sides]
    opp = any(x < 0 for x in signs)
    desc = "X" if opp else "K"
    if any(p < 0.25 for p in axiality):
        desc = "KT"
    return desc, True


def _contact_spread(contacts) -> float:
    pts = [c.point for c in contacts]
    return max((_norm(_sub(a, b)) for i, a in enumerate(pts) for b in pts[i + 1:]), default=0.0)


def classify_joint(members: list[Member], cand: JointCandidate, jid: str) -> Joint:
    """Classify a joint by connection configuration.

    joint_type  = connection config (corner/tee/splice/gusset/crossing)
    geom_descriptor = HSS geometry T/Y/K/X/KT  (hollow members only; AISC Ch. K)
    taxonomy    = 'hollow' -> AISC 360-22 ;  'open' -> AISI S100
    Conservative: anything ambiguous is flagged needs_review.
    """
    contacts = cand.contacts
    names = [_name_from_contact(members, c) for c in contacts]
    dirs = [c.direction for c in contacts]
    angles = [round(a, 1) for a in _pairwise_angles(dirs)]

    all_hollow = all(
        members[c.member_index].cross_section.family == SectionFamily.HOLLOW
        for c in contacts
    ) and len(contacts) > 0

    joint = Joint(
        joint_id=jid,
        location=tuple(round(x, 4) for x in cand.location),
        member_names=names,
        angles_deg=angles,
        is_inferred=cand.is_inferred,
        taxonomy="hollow" if all_hollow else "open",
    )

    through = [c for c in contacts if not c.at_end]
    ending = [c for c in contacts if c.at_end]
    n = len(contacts)

    if n >= 3:
        joint.joint_type = JointType.GUSSET
        joint.confidence = 0.6
        for c in contacts:
            joint.member_roles[_name_from_contact(members, c)] = "leg"
        # convergence test: wide contact spread => likely an over-merged cluster
        char = max((_member_size(members[c.member_index]) for c in contacts), default=0.0)
        spread = _contact_spread(contacts)
        if char > 0 and spread > 1.5 * char:
            joint.needs_review = True
            joint.review_reason = "wide_contact_spread_possible_overmerge"
        if all_hollow:
            desc, ok = _hollow_nbrace_descriptor(members, contacts)
            joint.geom_descriptor = desc
            if not ok:
                joint.needs_review = True
                joint.review_reason = (joint.review_reason + ";no_clear_chord").strip(";")
        else:
            joint.needs_review = True   # open cold-formed multi-member -> bolted/plate review
            joint.review_reason = (joint.review_reason + ";open_multimember_bolted_or_gusset").strip(";")
        return joint

    if n == 2:
        if len(through) == 2:
            joint.joint_type = JointType.CROSSING
            joint.confidence = 0.7
            member_a = members[contacts[0].member_index]
            member_b = members[contacts[1].member_index]
            if member_a.length >= member_b.length:
                joint.member_roles[_name_from_contact(members, contacts[0])] = "primary"
                joint.member_roles[_name_from_contact(members, contacts[1])] = "crossing"
            else:
                joint.member_roles[_name_from_contact(members, contacts[0])] = "crossing"
                joint.member_roles[_name_from_contact(members, contacts[1])] = "primary"
            return joint

        if len(through) == 1:
            joint.joint_type = JointType.TEE_CONNECTION
            joint.confidence = 0.75
            primary_role = "chord" if all_hollow else "primary"
            secondary_role = "brace" if all_hollow else "secondary"
            joint.member_roles[_name_from_contact(members, through[0])] = primary_role
            joint.member_roles[_name_from_contact(members, ending[0])] = secondary_role
            if all_hollow:                                  # NEW: T vs Y by angle
                joint.geom_descriptor = _hollow_2member_descriptor(angles[0] if angles else 90.0)
            return joint

        ang = angles[0] if angles else 0.0
        primary_role = "chord" if all_hollow else "primary"
        if ang >= 165:
            joint.joint_type = JointType.SPLICE
        else:
            joint.joint_type = JointType.CORNER
        joint.confidence = 0.75
        joint.member_roles[_name_from_contact(members, contacts[0])] = primary_role
        joint.member_roles[_name_from_contact(members, contacts[1])] = primary_role
        return joint

    joint.joint_type = JointType.UNKNOWN
    joint.confidence = 0.3
    joint.needs_review = True
    joint.review_reason = "single_or_zero_contact"
    return joint


def classify_all(members: list[Member], candidates: list[JointCandidate]) -> list[Joint]:
    return [classify_joint(members, c, f"J{idx+1:03d}")
            for idx, c in enumerate(candidates)]

print("OK: Joint classifier loaded (T/Y/K/X + unique names + review flags)")


# ############################################################################
# >>> notebook cell [15]
# ############################################################################
# ============================================================================
# SECTION 5: GLV FILTER (glv_filter.py)
# ============================================================================

_DESIGN_TRACKING = "Design Tracking Properties"


def _prop(definition_doc, name: str, default: str = "") -> str:
    """Read a Design Tracking Property value, tolerating absence."""
    try:
        ps = definition_doc.PropertySets.Item(_DESIGN_TRACKING)
        return str(ps.Item(name).Value or default)
    except Exception:
        return default


def _safe_full_file_name(occ) -> str:
    try:
        return str(occ.Definition.Document.FullFileName)
    except Exception:
        return ""


def _occurrence_path(parent_path: str, occ) -> str:
    """
    Create a stable-ish display path for nested Inventor occurrences.

    This prevents two parts named Part:1 in different subassemblies from looking
    like the same member in the JSON output.
    """
    try:
        name = str(occ.Name)
    except Exception:
        name = "UNKNOWN_OCCURRENCE"

    return f"{parent_path}/{name}" if parent_path else name


def _iter_occurrences(occurrences, parent_path: str = ""):
    """
    Yield every leaf part occurrence, descending into subassemblies.

    Returns:
        (occurrence, occurrence_path)
    """
    try:
        count = occurrences.Count
    except Exception:
        return

    for i in range(1, count + 1):
        try:
            occ = occurrences.Item(i)
        except Exception:
            continue

        path = _occurrence_path(parent_path, occ)

        try:
            subs = occ.SubOccurrences
            has_subs = subs is not None and subs.Count > 0
        except Exception:
            subs, has_subs = None, False

        if has_subs:
            yield from _iter_occurrences(subs, path)
        else:
            yield occ, path


def filter_glv_members(
    assembly_doc,
    tags: tuple[str, ...] = ("GLV", "HDG"),
    dedupe_exact_occurrence_paths: bool = True
) -> list[Member]:
    """
    Return one Member per physical GLV/HDG occurrence.

    Important:
    - This does NOT dedupe by part number or file name, because repeated parts
      can be real separate members in the assembly.
    - It only prevents the exact same occurrence path from being added twice.
    """
    tags = tuple(t.upper() for t in tags)
    members: list[Member] = []
    seen_paths = set()

    asm_def = assembly_doc.ComponentDefinition

    for occ, occ_path in _iter_occurrences(asm_def.Occurrences):
        if dedupe_exact_occurrence_paths and occ_path in seen_paths:
            continue

        try:
            def_doc = occ.Definition.Document
            desc = _prop(def_doc, "Description")
        except Exception:
            continue

        if not any(t in desc.upper() for t in tags):
            continue

        try:
            m = Member(
                occurrence_name=str(occ.Name),
                bom_description=desc,
                part_number=_prop(def_doc, "Part Number"),
                material=_prop(def_doc, "Material"),
                is_glv_or_hdg=True,
                cross_section=CrossSection(),
            )

            # Add useful unique identifiers.
            # These are dynamic attributes, so you do not need to change
            # the Member dataclass unless you want them exported to JSON.
            m.occurrence_path = occ_path
            m.source_file = _safe_full_file_name(occ)

            m._occ = occ
            members.append(m)
            seen_paths.add(occ_path)

        except Exception:
            continue

    return members


print("✓ GLV/HDG filter loaded with unique occurrence paths")


# ############################################################################
# >>> notebook cell [19]
# ############################################################################
# ============================================================================
# SECTION 6: MEMBER GEOMETRY (member_geometry.py)
# ============================================================================

def _transform_point(matrix, p):
    """Apply an Inventor Matrix to a (x,y,z) tuple."""
    x, y, z = p
    out = []
    for r in range(1, 4):
        out.append(
            matrix.Cell(r, 1) * x + matrix.Cell(r, 2) * y
            + matrix.Cell(r, 3) * z + matrix.Cell(r, 4)
        )
    return tuple(out)


def _centerline(occ, part_def):
    """
    Centreline endpoints in assembly space, exported in canonical inches.

    Inventor API/database length values are treated as cm. We transform the
    raw cm points into assembly space first, then convert the final points to
    inches exactly once at the boundary.
    """
    body = part_def.SurfaceBodies.Item(1)
    rb = body.RangeBox

    mn_cm = (rb.MinPoint.X, rb.MinPoint.Y, rb.MinPoint.Z)
    mx_cm = (rb.MaxPoint.X, rb.MaxPoint.Y, rb.MaxPoint.Z)

    dims_cm = [mx_cm[i] - mn_cm[i] for i in range(3)]
    axis = max(range(3), key=lambda i: dims_cm[i])
    center_cm = [(mn_cm[i] + mx_cm[i]) / 2.0 for i in range(3)]
    half_cm = dims_cm[axis] / 2.0

    s_cm = list(center_cm); s_cm[axis] -= half_cm
    e_cm = list(center_cm); e_cm[axis] += half_cm

    m = occ.Transformation
    s_asm_cm = _transform_point(m, tuple(s_cm))
    e_asm_cm = _transform_point(m, tuple(e_cm))

    s_in = pt_cm_to_in(s_asm_cm)
    e_in = pt_cm_to_in(e_asm_cm)
    length_in = math.dist(s_in, e_in)
    require_reasonable_member_length_in(length_in)

    return s_in, e_in, length_in


def _tessellate_arc(arc, n: int = 10):
    c = arc.CenterSketchPoint.Geometry
    r = arc.Radius
    a0 = arc.StartAngle
    sweep = arc.SweepAngle
    pts = []
    for i in range(n + 1):
        a = a0 + sweep * i / n
        pts.append((c.X + r * math.cos(a), c.Y + r * math.sin(a)))
    return pts


def _path_points(path):
    pts = []
    for i in range(1, path.Count + 1):
        entity = path.Item(i)
        se = entity.SketchEntity
        kind = se.Type
        if "Line" in str(kind) or hasattr(se, "StartSketchPoint") and not hasattr(se, "Radius"):
            g0 = se.StartSketchPoint.Geometry
            pts.append((g0.X, g0.Y))
        elif hasattr(se, "Radius") and hasattr(se, "SweepAngle"):
            pts.extend(_tessellate_arc(se))
        elif hasattr(se, "Radius"):
            c = se.CenterSketchPoint.Geometry
            r = se.Radius
            pts.extend((c.X + r * math.cos(2 * math.pi * k / 32),
                        c.Y + r * math.sin(2 * math.pi * k / 32)) for k in range(32))
    return pts


def _loop_area(pts):
    a = 0.0
    n = len(pts)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def _extrude_loops(part_def):
    """Profile loops from the dominant extrude (works for HSS / solid extrudes)."""
    extrudes = part_def.Features.ExtrudeFeatures
    if extrudes.Count == 0:
        return []
    best, best_area = None, -1.0
    for i in range(1, extrudes.Count + 1):
        ext = extrudes.Item(i)
        try:
            prof = ext.Profile
            total = sum(_loop_area(_path_points(prof.Item(k)))
                        for k in range(1, prof.Count + 1))
        except Exception:
            total = -1.0
        if total > best_area:
            best, best_area = ext, total
    if best is None:
        return []
    prof = best.Profile
    loops = [_path_points(prof.Item(k)) for k in range(1, prof.Count + 1)]
    loops = [lp for lp in loops if len(lp) >= 3]
    loops.sort(key=_loop_area, reverse=True)
    return loops


def _slice_to_loops(verts, tris, axis, cut, weld=1e-4):
    """Intersect a triangle mesh with a plane (axis == cut) and chain the
    crossing segments into closed 2D loops. Feature-agnostic."""
    others = [i for i in range(3) if i != axis]
    segs = []
    for (ia, ib, ic) in tris:
        tri = [verts[ia], verts[ib], verts[ic]]
        pts = []
        for k in range(3):
            a = tri[k]; b = tri[(k + 1) % 3]
            da = a[axis] - cut; db = b[axis] - cut
            if abs(da) < 1e-12: da = 0.0
            if abs(db) < 1e-12: db = 0.0
            if (da < 0 and db > 0) or (da > 0 and db < 0):
                t = da / (da - db)
                p = tuple(a[i] + t * (b[i] - a[i]) for i in range(3))
                pts.append((p[others[0]], p[others[1]]))
            elif da == 0.0:
                pts.append((a[others[0]], a[others[1]]))
        uniq = []
        for p in pts:
            if not any(math.hypot(p[0] - q[0], p[1] - q[1]) < weld for q in uniq):
                uniq.append(p)
        if len(uniq) == 2:
            segs.append((uniq[0], uniq[1]))

    def key(p):
        return (round(p[0] / weld), round(p[1] / weld))

    coord, adj = {}, {}
    for (p, q) in segs:
        kp, kq = key(p), key(q)
        coord.setdefault(kp, p); coord.setdefault(kq, q)
        if kp == kq:
            continue
        adj.setdefault(kp, set()).add(kq)
        adj.setdefault(kq, set()).add(kp)

    loops, used = [], set()
    for start in list(adj.keys()):
        if start in used or not adj.get(start):
            continue
        loop = [start]; used.add(start); prev, cur = None, start
        while True:
            nxts = [n for n in adj[cur] if n != prev and n not in used]
            if not nxts:
                break
            nxt = nxts[0]
            if nxt == start:
                break
            loop.append(nxt); used.add(nxt); prev, cur = cur, nxt
        if len(loop) >= 3:
            loops.append([coord[k] for k in loop])
    return loops


def _slice_loops(part_def):
    """Profile loops by slicing the solid body at mid-span. Works regardless of
    how the part was modelled (sheet-metal/cold-formed, sweep, import, etc.)."""
    body = part_def.SurfaceBodies.Item(1)
    rb = body.RangeBox
    mn = (rb.MinPoint.X, rb.MinPoint.Y, rb.MinPoint.Z)
    mx = (rb.MaxPoint.X, rb.MaxPoint.Y, rb.MaxPoint.Z)
    dims = [mx[i] - mn[i] for i in range(3)]
    axis = max(range(3), key=lambda i: dims[i])
    cut = (mn[axis] + mx[axis]) / 2.0
    small = min(d for d in dims if d > 0)
    tol = max(1e-4, 0.02 * small)

    # Inventor tessellation. win32com returns the [out] params as a tuple.
    # If this binding ever fails, it is caught upstream and the member is left
    # unclassified (no regression vs. the old extrude-only behaviour).
    vc, fc, coords, normals, idx = body.CalculateFacets(tol)
    verts = [(coords[3 * i], coords[3 * i + 1], coords[3 * i + 2]) for i in range(vc)]
    tris = [(idx[3 * i] - 1, idx[3 * i + 1] - 1, idx[3 * i + 2] - 1) for i in range(fc)]

    loops = _slice_to_loops(verts, tris, axis, cut)
    loops = [lp for lp in loops if len(lp) >= 3]
    loops.sort(key=_loop_area, reverse=True)
    return loops


def _section_loops(part_def):
    """
    Return (loops_cm, detection_method).

    The returned loop coordinates are still raw Inventor API/database cm. They
    must be converted to inches before classification/properties/export.
    """
    loops = _extrude_loops(part_def)
    if loops:
        return loops, "extrude_profile"
    try:
        loops = _slice_loops(part_def)
        if loops:
            return loops, "midspan_slice"
    except Exception:
        pass
    return [], "geometry"




def _section_hint_from_member(m: Member) -> Optional[SectionType]:
    """
    Read a final section-label hint from Inventor occurrence/part/BOM text.

    Geometry still provides depth/width/thickness. This hint only corrects the
    final label when the coarse slice signature is ambiguous, especially for
    hat/omega and lipped sections.
    """
    text = " ".join([
        str(getattr(m, "occurrence_name", "")),
        str(getattr(m, "part_number", "")),
        str(getattr(m, "bom_description", "")),
    ]).lower()

    text = (
        text.replace("-", "_")
            .replace(" ", "_")
            .replace("(", "_")
            .replace(")", "_")
            .replace("/", "_")
    )

    without_lips = any(k in text for k in [
        "without_lips", "without_lip", "no_lips", "no_lip", "unlipped"
    ])

    with_lips = any(k in text for k in [
        "with_lips", "with_lip", "lipped", "lip"
    ]) and not without_lips

    # Hollow sections first so "circular_section" is not confused with C-section.
    if "hollow" in text and any(k in text for k in ["circular", "round", "pipe", "tube"]):
        return SectionType.ROUND_HSS

    if "hollow" in text and any(k in text for k in ["box", "rect", "square", "rhs", "hss"]):
        return SectionType.RECT_HSS

    # IMPORTANT: check Hat/H/Omega before C/channel logic.
    # Hat profiles can produce a C-like 3x3 occupancy signature in midspan slices.
    if any(k in text for k in ["h_section", "hat", "omega"]):
        return SectionType.H_LIPPED if with_lips else SectionType.H_UNLIPPED

    if any(k in text for k in ["c_section", "channel", "c_channel"]):
        return SectionType.C_LIPPED if with_lips else SectionType.C_UNLIPPED

    if any(k in text for k in ["l_section", "angle"]):
        return SectionType.L_LIPPED if with_lips else SectionType.L_UNLIPPED

    if any(k in text for k in ["z_section", "zed"]):
        return SectionType.Z_LIPPED if with_lips else SectionType.Z_UNLIPPED

    return None


def _apply_section_hint(cs: CrossSection, hint: Optional[SectionType]) -> CrossSection:
    """
    Apply metadata as a final label correction only.

    This preserves measured geometry values from extrude_profile/midspan_slice:
    depth, width, wall_thickness, n_loops, and occupancy_signature.
    """
    if hint is None:
        return cs

    changed = cs.section_type != hint

    if changed:
        cs.section_type = hint
        cs.family = SectionFamily.HOLLOW if hint in HOLLOW_TYPES else SectionFamily.OPEN
        cs.confidence = max(cs.confidence, 0.80)

    if hint in _LIPPED_OPEN_TYPES:
        cs.is_lipped = True
    elif hint in _UNLIPPED_OPEN_TYPES or hint in HOLLOW_TYPES:
        cs.is_lipped = False

    if "+metadata_hint" not in cs.detection_method:
        cs.detection_method = f"{cs.detection_method}+metadata_hint"

    return cs


_Z_OPEN_TYPES = {
    SectionType.Z_UNLIPPED,
    SectionType.Z_LIPPED,
}

_EXPLICIT_SECTION_WORDS = [
    "c_section", "c_channel", "channel",
    "h_section", "hat", "omega",
    "l_section", "angle",
    "z_section", "zed",
    "hollow", "hss", "rhs", "tube", "pipe", "round", "circular", "box", "rect", "square",
]

_NON_PROFILE_WORDS = [
    "plate", "sheet", "panel", "pan", "tray", "bracket", "gusset", "cover", "cap", "clip", "guard",
]


def _member_text_normalized(m: Member) -> str:
    """Normalize Inventor occurrence/part/BOM text for metadata checks."""
    text = " ".join([
        str(getattr(m, "occurrence_name", "")),
        str(getattr(m, "part_number", "")),
        str(getattr(m, "bom_description", "")),
    ]).lower()
    return (
        text.replace("-", "_")
            .replace(" ", "_")
            .replace("(", "_")
            .replace(")", "_")
            .replace("/", "_")
    )


def _mark_unknown(cs: CrossSection, reason: str, max_conf: float = 0.35) -> CrossSection:
    """Keep measured geometry, but suppress an unsafe section label."""
    cs.section_type = SectionType.UNKNOWN
    cs.family = SectionFamily.OPEN
    cs.is_lipped = None
    cs.confidence = min(cs.confidence, max_conf)
    if reason not in cs.detection_method:
        cs.detection_method = f"{cs.detection_method}+{reason}"
    return cs


def _guard_false_z_from_midspan(m: Member, cs: CrossSection, hint: Optional[SectionType]) -> CrossSection:
    """
    Prevent bent plates/trays/brackets from being classified as Z-sections.

    A real Z section may still be accepted when the Inventor metadata explicitly
    says Z/zed. Otherwise, Z labels coming only from a midspan slice are treated
    as ambiguous and downgraded to UNKNOWN.
    """
    if cs.section_type not in _Z_OPEN_TYPES:
        return cs

    if hint in _Z_OPEN_TYPES:
        return cs

    text = _member_text_normalized(m)
    has_explicit_section_word = any(k in text for k in _EXPLICIT_SECTION_WORDS)
    looks_like_non_profile_part = any(k in text for k in _NON_PROFILE_WORDS)

    # The safest rule for your current dataset: a Z detected only from a
    # midspan mesh slice needs explicit Z metadata confirmation. This stops
    # formed sheet-metal plates like the screenshot from becoming z_unlipped.
    if "midspan_slice" in cs.detection_method:
        return _mark_unknown(cs, "z_guard")

    # Extra safety for extrude/sketch profiles named like plate/bracket/tray.
    if looks_like_non_profile_part and not has_explicit_section_word:
        return _mark_unknown(cs, "non_profile_guard")

    return cs


def fill_geometry(members: list[Member]) -> list[Member]:
    """
    Populate centreline and cross_section for each member.

    Optimized order:
    1. Get centerline from geometry.
    2. Extract cross-section loops from extrude profile; otherwise midspan slice.
    3. Classify section from geometry.
    4. Apply metadata hint as the final label correction.
    """
    for m in members:
        occ = getattr(m, "_occ", None)
        if occ is None:
            continue

        part_def = occ.Definition

        try:
            s, e, length = _centerline(occ, part_def)
            m.start_point, m.end_point, m.length = s, e, length
        except Exception:
            pass

        try:
            loops_cm, method = _section_loops(part_def)
            if loops_cm:
                # Convert raw Inventor API/database cm to canonical inches before
                # section classification and property calculation. This prevents
                # cm/in mixing in depth, width, thickness, A, Iy, Iz, and J.
                loops_in = loops_cm_to_in(loops_cm)

                cs = classify_section(loops_in, detection_method=method)
                hint = _section_hint_from_member(m)
                cs = _guard_false_z_from_midspan(m, cs, hint)
                m.cross_section = _apply_section_hint(cs, hint)
        except Exception:
            pass

    return members

print("✓ Optimized member geometry loader loaded with false-Z guard and cm→in boundary conversion")


# ############################################################################
# >>> notebook cell [23]
# ############################################################################
# ============================================================================
# SECTION 7: INVENTOR SESSION (inventor_session.py)
# ============================================================================

class InventorError(RuntimeError):
    pass


class InventorSession:
    """Attaches to a running Inventor instance."""

    def __init__(self, launch_if_absent: bool = False, visible: bool = True):
        if not _HAVE_WIN32:
            raise InventorError(
                "pywin32 not available. Install: pip install pywin32"
            )
        self.launch_if_absent = launch_if_absent
        self.visible = visible
        self.app = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, *exc):
        self.app = None

    def connect(self):
        pythoncom.CoInitialize()
        try:
            self.app = _w32.GetActiveObject("Inventor.Application")
        except Exception:
            if not self.launch_if_absent:
                raise InventorError(
                    "No running Inventor instance found. Open Inventor with the model."
                )
            self.app = _w32.Dispatch("Inventor.Application")
            self.app.Visible = self.visible
        return self.app

    def active_document(self):
        if self.app is None:
            self.connect()
        doc = self.app.ActiveDocument
        if doc is None:
            raise InventorError("No active document is open in Inventor.")
        return doc

    def active_assembly(self):
        """Return the active document, asserting it is an assembly."""
        doc = self.active_document()
        if int(doc.DocumentType) != 12291:
            raise InventorError(
                "Active document is not an assembly (.iam). "
                "Open the assembly file as the active document."
            )
        return doc

print("✓ Inventor session loaded")


# ############################################################################
# >>> notebook cell [25]
# ############################################################################
# ============================================================================
# SECTION 8: ORCHESTRATOR (orchestrator.py)   [UNIT-SAFE for Block 2 hand-off]
# ============================================================================


def _inventor_display_units_audit(asm) -> dict:
    """
    Non-authoritative audit only. Block 1 does NOT use display units for math;
    all raw API values are converted to canonical units in the unit guard.
    """
    audit = {}
    try:
        uom = asm.ComponentDefinition.Document.UnitsOfMeasure
        audit["inventor_display_length"] = str(uom.LengthUnits)
        audit["inventor_display_mass"] = str(uom.MassUnits)
    except Exception:
        audit["inventor_display_units"] = "unavailable"
    return audit


def run(
    tags: tuple[str, ...] = ("GLV", "HDG"),
    tol_touch_in: float = DEFAULT_TOL_TOUCH_IN,
    cluster_cap_in: float | None = DEFAULT_CLUSTER_CAP_IN,
    support_z_tol_in: float = DEFAULT_SUPPORT_Z_TOL_IN,
    # Backward-compatible aliases. Prefer the *_in names in new code.
    tol_touch: float | None = None,
    cluster_cap: float | None = None,
    support_z_tol: float | None = None,
) -> AnalysisResult:
    """
    Run the full Block 1 pipeline and emit a Block-2-ready result.

    Unit contract:
      - Inventor API/database length is converted from cm to in.
      - Inventor MassProperties.Mass is converted from kg to lbm and lbf.
      - All exported points, lengths, tolerances, and section properties are in
        canonical Block-2 units.
    """
    # Legacy argument support so older cells/scripts do not silently change behavior.
    if tol_touch is not None:
        tol_touch_in = tol_touch
    if cluster_cap is not None:
        cluster_cap_in = cluster_cap
    if support_z_tol is not None:
        support_z_tol_in = support_z_tol

    tol_touch_in = require_positive_length_in(tol_touch_in, "tol_touch_in")
    support_z_tol_in = require_positive_length_in(support_z_tol_in, "support_z_tol_in")
    if cluster_cap_in is not None:
        cluster_cap_in = require_positive_length_in(cluster_cap_in, "cluster_cap_in")

    with InventorSession() as inv:
        asm = inv.active_assembly()
        result = AnalysisResult(source_document=str(asm.FullFileName))

        # Step 0: canonical unit declaration + non-authoritative Inventor display audit.
        result.units = {
            **CANONICAL_UNITS,
            **_inventor_display_units_audit(asm),
            "tol_touch": f"{tol_touch_in} in",
            "cluster_cap": None if cluster_cap_in is None else f"{cluster_cap_in} in",
            "support_z_tol": f"{support_z_tol_in} in",
            "_note": (
                "All exported geometry is canonical: in, in^2, in^4, lbm, lbf, deg. "
                "Inventor API/database length is treated as cm and mass as kg. "
                "Inventor display units are recorded only for audit, not used for math."
            ),
        }
        assert_block2_canonical_units(result)

        # Step 1: filter GLV/HDG members (with unique occurrence_path).
        members = filter_glv_members(asm, tags=tags)

        # Step 2: fill geometry. _centerline and profile loops are converted to inches here.
        members = fill_geometry(members)

        # Step 2b: solver-ready section properties in in^2/in^4.
        for m in members:
            compute_section_properties(m.cross_section)

        # Step 2c: per-member dry mass + self-weight force for Block 2 dead load.
        for m in members:
            try:
                mass_kg = float(m._occ.MassProperties.Mass)  # Inventor API/database mass: kg
                m.dry_mass = kg_to_lbm(mass_kg)              # canonical mass: lbm
                m.mass_unit = "lbm"
                m.self_weight_lbf = kg_to_lbf_weight(mass_kg) # canonical force: lbf
                m.force_unit = "lbf"
            except Exception:
                pass

        result.members = members

        # Step 3: detect joint candidates using inch-based tolerances.
        candidates = detect_joints(
            members,
            tol_touch=tol_touch_in,
            cluster_cap=cluster_cap_in,
        )

        # Step 4: classify joints (T/Y/K/X + review flags).
        result.joints = classify_all(members, candidates)
        # near the other module constants
        VERTICAL_AXIS_INDEX = 1   # 0=X, 1=Y, 2=Z. Model is Y-up (gravity acts -Y),
                                # matching block3_solver DEFAULT_CONFIG["vertical_axis"].

        # Step 5: flag candidate support nodes (lowest-Z joints) using inch tolerance.
        # Step 5 — flag candidate supports: lowest joints along the vertical axis,
        # within an inch tolerance band. Y-up model, so the base sits at minimum Y.
        if result.joints:
            i = VERTICAL_AXIS_INDEX
            vmin = min(j.location[i] for j in result.joints)
            for j in result.joints:
                if j.location[i] <= vmin + support_z_tol_in:
                    j.is_support_candidate = True

    # Drop live COM handles before serialising.
    for m in result.members:
        if hasattr(m, "_occ"):
            del m._occ

    assert_block2_canonical_units(result)
    return result


def summarise(result: AnalysisResult) -> None:
    assert_block2_canonical_units(result)

    n_inf = sum(1 for j in result.joints if j.is_inferred)
    n_rev = sum(1 for j in result.joints if j.needs_review)
    n_props = sum(1 for m in result.members if m.cross_section.A is not None)
    n_props_clean = sum(
        1 for m in result.members
        if m.cross_section.A is not None and not m.cross_section.needs_review
    )
    n_sup = sum(1 for j in result.joints if j.is_support_candidate)
    n_weight = sum(1 for m in result.members if m.self_weight_lbf is not None)

    print(f"\n{'='*60}")
    print(f"Units (canonical):     length={result.units.get('length')}, "
          f"mass={result.units.get('mass')}, force={result.units.get('force')}")
    print(f"Inventor display audit:{ {k: v for k, v in result.units.items() if k.startswith('inventor_')} }")
    print(f"Members (GLV/HDG):     {len(result.members)}")
    print(f"  with self weight:    {n_weight}")
    print(f"  with section props:  {n_props}  ({n_props_clean} clean / authoritative)")
    print(f"Joints:                {len(result.joints)}  ({n_inf} inferred, {n_rev} need review)")
    print(f"  support candidates:  {n_sup}")

    by_type = {}
    by_desc = {}
    for j in result.joints:
        by_type[j.joint_type.value] = by_type.get(j.joint_type.value, 0) + 1
        if j.geom_descriptor:
            by_desc[j.geom_descriptor] = by_desc.get(j.geom_descriptor, 0) + 1

    print("\nBy configuration:")
    for k, v in sorted(by_type.items()):
        print(f"  {k:18} {v}")
    if by_desc:
        print("\nHSS geometry (T/Y/K/X):")
        for k, v in sorted(by_desc.items()):
            print(f"  {k:18} {v}")
    print(f"{'='*60}\n")

print("OK: Orchestrator loaded (canonical units, inch tolerances, lbm/lbf mass/weight)")


# ############################################################################
# >>> app integration: capability probe + one-call live extraction
# (Added for the Streamlit app; not part of the notebook. Keeps the live-vs-
#  upload routing decision next to the pipeline it gates.)
# ############################################################################
def live_inventor_status() -> dict:
    """Cheap, exception-safe probe for a reachable live Inventor session.

    Returns {"available": bool, "reason": str}. Safe to call on any host and on
    every Streamlit rerun: it never raises and never launches Inventor.
    """
    if not _HAVE_WIN32:
        return {"available": False,
                "reason": "pywin32 not installed (expected off Windows)."}
    try:
        pythoncom.CoInitialize()
        _w32.GetActiveObject("Inventor.Application")
        return {"available": True, "reason": "Live Inventor session found."}
    except Exception:
        return {"available": False,
                "reason": "No running Inventor instance. Open Inventor with the model."}


def extract_live_to_dict() -> dict:
    """Run the frozen Block-1 pipeline against the live session and return a
    plain JSON-safe dict identical to an uploaded block1_result_unit_safe.json.

    Uses the notebook's hardcoded default tolerances. The dict is produced via
    to_json()/json.loads so Enums are strings and tuples are lists, byte-for-
    byte matching the upload path that downstream blocks already consume.
    """
    result = run(
        tags=("GLV", "HDG"),
        tol_touch_in=DEFAULT_TOL_TOUCH_IN,
        cluster_cap_in=DEFAULT_CLUSTER_CAP_IN,
        support_z_tol_in=DEFAULT_SUPPORT_Z_TOL_IN,
    )
    return json.loads(result.to_json())


# Extraction parameters surfaced read-only in the UI for documentation.
EXTRACTION_PARAMS = {
    "tags": ("GLV", "HDG"),
    "tol_touch_in": DEFAULT_TOL_TOUCH_IN,
    "cluster_cap_in": DEFAULT_CLUSTER_CAP_IN,
    "support_z_tol_in": DEFAULT_SUPPORT_Z_TOL_IN,
}
