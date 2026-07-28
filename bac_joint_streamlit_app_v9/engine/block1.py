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
from dataclasses import dataclass, field, asdict, replace
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

# Offset tolerance for the FACE-CONTACT gate (how far two near-parallel, opposing
# faces may stand apart and still count as a contact). Deliberately decoupled from
# and looser than tol_touch: a lapped brace end commonly sits ~1/4" off the mating
# flange (gusset plate / intentional setback), so a 0.2" touch tolerance wrongly
# rejected the whole diagonal-to-base-channel load path. Raising only THIS value
# recovers those laps without changing centerline joint clustering.
DEFAULT_CONTACT_OFFSET_TOL_IN = 0.35

# When the face pass CANNOT confirm a contact, a member pair is kept as a flagged
# "near_contact" only if its closest candidate contact faces are within this gap;
# beyond it the members clearly are not touching (e.g. a diagonal brace ~1.2" off
# the vertical channel it merely passes near) and the pair is dropped. ~2x the
# contact offset tolerance: loose enough to keep a real gusseted near-miss, tight
# enough to reject a member the brace only bolts NEAR, not to.
DEFAULT_NEAR_CONTACT_GAP_IN = 0.75


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
    # Unit vectors, assembly/global space, pointing along the depth/width
    # extents of the cross-section. Only populated when the profile came from
    # a midspan slice (the cut plane's two in-plane axes are known in global
    # space); None for extrude_profile detections, where the sketch-plane
    # orientation isn't captured.
    depth_dir: Optional[tuple[float, float, float]] = None
    width_dir: Optional[tuple[float, float, float]] = None
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
    cost: dict = field(default_factory=dict)   # Milestone 2: cost iProperties + part_class (exported via asdict)


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
class Connection:
    """A single member-pair contact -- two members touching.

    Connection/joint model:
        connection = a face of two different members touching (ONE contact).
        joint      = two or more members touching together (a cluster of
                     connections); see ``joint_id``.

    Several connections can share one ``joint_id``; that is exactly what makes
    the joint a multi-member node. The face-based metrics (contact_area_in2,
    weld_length_in, joint_length_in) come from real face geometry resolved in a
    live Inventor session (``detection_method == "face_contact"``); on the
    centerline fallback (``"centerline"``) they stay None.
    """
    connection_id: str
    member_a: str = ""                 # occurrence_path (or name) of member 1
    member_b: str = ""                 # occurrence_path (or name) of member 2
    location: tuple[float, float, float] = (0.0, 0.0, 0.0)   # contact centroid, in
    joint_id: str = ""                 # cluster this contact belongs to
    connection_type: str = "unknown"   # bolted | welded | unknown
    angle_deg: Optional[float] = None  # acute angle between member axes
    contact_area_in2: Optional[float] = None
    weld_length_in: Optional[float] = None     # perimeter of the contact overlap
    joint_length_in: Optional[float] = None    # longest span of the contact overlap
    detection_method: str = "centerline"       # centerline | face_contact
    is_inferred: bool = False
    needs_review: bool = False
    review_reason: str = ""

    # --- per-face-pair patch geometry (face_contact path only) -------------
    # A member pair touching on several faces produces several connections, one
    # per face pair. ``face_pair_index`` distinguishes them within the pair.
    face_pair_index: int = 0
    contact_normal: Optional[tuple[float, float, float]] = None
    hole_count: int = 0                        # bolt holes inside the patch
    # Patch outline in its own plane (inches). This is the frame a bolt pattern
    # is laid out in and a weld path runs around.
    patch_frame: Optional[dict] = None         # {origin, x_axis, y_axis}
    patch_exterior_2d: list = field(default_factory=list)
    patch_holes_2d: list = field(default_factory=list)
    # Cross-check: what the OBB (JointLocatorV16) method would have reported.
    obb_area_in2: Optional[float] = None
    obb_divergence: Optional[float] = None     # |exact - obb| / obb
    obb_polygon_3d: list = field(default_factory=list)   # V16 box-clip ring, world in


@dataclass
class AnalysisResult:
    source_document: str = ""
    units: dict = field(default_factory=dict)   # NEW: section-0 unit self-check
    members: list[Member] = field(default_factory=list)
    joints: list[Joint] = field(default_factory=list)
    connections: list[Connection] = field(default_factory=list)   # NEW: face/centerline contacts
    connection_diagnostics: dict = field(default_factory=dict)     # NEW: face-detection stage counters
    joint_cluster_tol_in: Optional[float] = None   # NEW: joint clustering RADIUS (sphere, not a box)
    timings: dict = field(default_factory=dict)                    # NEW: per-stage wall-clock seconds

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


def classify_section(loops: list[Loop], detection_method: str = "extrude_profile",
                      axis0_dir: Optional[tuple[float, float, float]] = None,
                      axis1_dir: Optional[tuple[float, float, float]] = None) -> CrossSection:
    if not loops or len(loops[0]) < 3:
        return CrossSection(detection_method=detection_method, confidence=0.0)

    outer = loops[0]
    raw_holes = [lp for lp in loops[1:] if len(lp) >= 3]
    n_loops = 1 + len(raw_holes)

    minx, miny, maxx, maxy = _bbox(outer)
    bb_w, bb_h = maxx - minx, maxy - miny
    depth, width = max(bb_w, bb_h), min(bb_w, bb_h)
    # bb_w is the extent along axis0_dir, bb_h along axis1_dir (see _bbox /
    # _slice_to_loops: loop points are (p[others[0]], p[others[1]])).
    if bb_w >= bb_h:
        depth_dir, width_dir = axis0_dir, axis1_dir
    else:
        depth_dir, width_dir = axis1_dir, axis0_dir

    # Hole filtering is separate from open-section family matching.
    holes = _structural_hollow_holes(outer, raw_holes, bb_w, bb_h)

    cs = CrossSection(
        detection_method=detection_method,
        n_loops=n_loops,
        depth=depth,
        width=width,
        depth_dir=depth_dir,
        width_dir=width_dir,
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
    # Member-pair contacts that were clustered into this joint. Each entry is
    # (member_index_i, member_index_j, contact_location, is_inferred) -- i.e. one
    # connection. Kept so run() can emit connections tagged with the joint they
    # belong to without re-deriving the clustering.
    pairs: list = field(default_factory=list)


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


def joint_cluster_tolerance(
    members: list[Member],
    cluster_factor: float = 1.2,
    min_cluster: float = 1.0,
    cluster_cap: Optional[float] = None,
) -> float:
    """Radius of the sphere that merges contact points into a single joint.

    NOT a bounding box: joints are clustered by straight distance from the first
    contact point that seeded them (see detect_joints). This is the single number
    that decides whether two connections belong to the same joint or to two
    different ones, so it is exported on AnalysisResult and drawn by the viewer.
    """
    sizes = [_member_size(m) for m in members if _member_size(m) > 0]
    char = sorted(sizes)[len(sizes) // 2] if sizes else 0.0
    tol = max(min_cluster, cluster_factor * char)
    if cluster_cap is not None:
        tol = min(tol, cluster_cap)
    return tol


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

    cluster_tol = joint_cluster_tolerance(
        members, cluster_factor=cluster_factor, min_cluster=min_cluster,
        cluster_cap=cluster_cap)

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
        host.pairs.append((ci.member_index, cj.member_index, loc, inferred))
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
# >>> app integration: CONNECTIONS (member/face-pair contacts)   [NEW]
# ############################################################################
# ============================================================================
# SECTION 4b: CONNECTIONS
# ============================================================================
# A connection is a single member-pair contact -- "a face of two members
# touching". Joints (above) are clusters of 2+ members; connections are the
# individual contacts that make up those clusters. Two detection paths:
#   * centerline   -- always available; reuses the member-pair contacts that
#                     detect_joints already found. Approximate: no true contact
#                     area or weld length.
#   * face_contact -- live Inventor only; enriches each connection with real
#                     face-overlap area / weld length / joint length / centroid,
#                     ported from the JointLocator VB macro.

from engine import connection_geometry as _cg

# Edge tessellation chord tolerance (cm) -- fine enough to resolve a bolt hole.
_STROKE_TOL_CM = 0.05
# Reject slivers. V16 has no minimum area, so a kiss of edge contact emits a
# connection; here a patch must be big enough to actually fasten.
_MIN_CONTACT_AREA_IN2 = 0.25
# Two faces must be this close to parallel to be a contact (cos of ~10 deg).
_PARALLEL_COS = 0.985


def _acute_angle_between_members(ma: Member, mb: Member) -> float:
    """Acute angle (deg) between two member centreline axes."""
    a = _unit(_sub(ma.end_point, ma.start_point))
    b = _unit(_sub(mb.end_point, mb.start_point))
    if _norm(a) < 1e-9 or _norm(b) < 1e-9:
        return 0.0
    ang = _angle_deg(a, b)
    return round(min(ang, 180.0 - ang), 1)


def _connection_type(ma: Member, mb: Member) -> str:
    """Bolted vs welded from section family, mirroring the VB HasFlatWidthIPT
    test: two cold-formed/open (sheet-metal) parts bolt; two hollow (HSS/tube)
    parts weld; a mix is left unknown for review."""
    fa = ma.cross_section.family
    fb = mb.cross_section.family
    if fa == SectionFamily.HOLLOW and fb == SectionFamily.HOLLOW:
        return "welded"
    if fa == SectionFamily.OPEN and fb == SectionFamily.OPEN:
        return "bolted"
    return "unknown"


def build_centerline_connections(
    members: list[Member],
    candidates: list[JointCandidate],
    joints: list[Joint],
) -> list[Connection]:
    """One Connection per member-pair contact, tagged with the joint it forms.

    ``candidates`` and ``joints`` are index-aligned (classify_all preserves
    order), so each candidate's ``pairs`` map straight onto the joint's id.
    """
    conns: list[Connection] = []
    cid = 1
    for cand, joint in zip(candidates, joints):
        for (i, j, loc, inferred) in cand.pairs:
            ma, mb = members[i], members[j]
            mixed = ma.cross_section.family != mb.cross_section.family
            conns.append(Connection(
                connection_id=f"C{cid:03d}",
                member_a=ma.occurrence_path or ma.occurrence_name,
                member_b=mb.occurrence_path or mb.occurrence_name,
                location=tuple(round(x, 4) for x in loc),
                joint_id=joint.joint_id,
                connection_type=_connection_type(ma, mb),
                angle_deg=_acute_angle_between_members(ma, mb),
                detection_method="centerline",
                is_inferred=inferred,
                needs_review=mixed,
                review_reason="mixed_open_hollow_connection" if mixed else "",
            ))
            cid += 1
    return conns


def _diag_err(diag: dict, stage: str, exc: Exception) -> None:
    """Record a de-duplicated, capped error string for a detection stage."""
    if diag is None:
        return
    errs = diag.setdefault("errors", [])
    msg = f"{stage}: {type(exc).__name__}: {exc}"
    if len(errs) < 12 and msg not in errs:
        errs.append(msg)


def _face_raw(face, diag=None):
    """Pull (vertices, normal, longest_edge_dir) from a PLANAR Inventor face, in
    raw Inventor cm. Returns None for non-planar faces or on any COM error.

    ``face.Geometry.Normal`` doubles as the planarity test: only planar faces
    expose a Plane geometry with a .Normal, so a non-planar (cylindrical, etc.)
    face raises here and is skipped. Errors are recorded in ``diag`` so a live
    run can show whether the COM access itself is the failure point."""
    try:
        normal_uv = face.Geometry.Normal
        normal = (normal_uv.X, normal_uv.Y, normal_uv.Z)
    except Exception as exc:
        # Curved faces (cylinders/cones -- e.g. every bolt hole) have no Plane
        # geometry and legitimately land here. Count them, don't treat as errors,
        # but keep the first reason in case it's actually a COM access failure.
        if diag is not None:
            diag["non_planar_faces"] = diag.get("non_planar_faces", 0) + 1
            diag.setdefault("first_non_planar_reason", f"{type(exc).__name__}: {exc}")
        return None
    try:
        verts = []
        vs = face.Vertices
        for i in range(1, vs.Count + 1):
            p = vs.Item(i).Point
            verts.append((p.X, p.Y, p.Z))
        if len(verts) < 3:
            return None
    except Exception as exc:
        _diag_err(diag, "face_vertices", exc)
        return None
    # Longest edge -> in-plane x-axis hint. Non-fatal: build_face_obb falls back
    # to a vertex-derived axis when this is missing.
    x_hint = None
    try:
        max_len = 0.0
        es = face.Edges
        for i in range(1, es.Count + 1):
            e = es.Item(i)
            p0 = e.StartVertex.Point
            p1 = e.StopVertex.Point
            d = (p1.X - p0.X, p1.Y - p0.Y, p1.Z - p0.Z)
            L = math.sqrt(d[0] ** 2 + d[1] ** 2 + d[2] ** 2)
            if L > max_len:
                max_len, x_hint = L, d
    except Exception as exc:
        _diag_err(diag, "face_edges", exc)
        x_hint = None
    return verts, normal, x_hint


def _edge_points(edge, diag) -> Optional[list]:
    """Tessellate one edge into points (raw Inventor cm).

    ``GetStrokes`` handles arcs, fillets and bolt-hole circles. If it is
    unavailable the edge falls back to the chord between its vertices, and a
    CURVED edge reduced to a chord is flagged rather than silently linearised --
    a bolt hole flattened into a triangle would understate the hole and overstate
    the bearing area.
    """
    try:
        res = edge.Evaluator.GetStrokes(_STROKE_TOL_CM)
        coords = res[1] if isinstance(res, tuple) and len(res) >= 2 else res
        flat = list(coords)
        if len(flat) >= 6:
            return [tuple(flat[i:i + 3]) for i in range(0, len(flat) - 2, 3)]
    except Exception:
        pass
    try:
        p0, p1 = edge.StartVertex.Point, edge.StopVertex.Point
    except Exception as exc:
        _diag_err(diag, "edge_vertices", exc)
        return None
    try:
        if edge.GeometryType != 5121:          # kLineSegmentCurve
            diag["curved_edges_chorded"] = diag.get("curved_edges_chorded", 0) + 1
    except Exception:
        pass
    return [(p0.X, p0.Y, p0.Z), (p1.X, p1.Y, p1.Z)]


def _loop_points(loop, diag) -> Optional[list]:
    """Ordered point ring for ONE EdgeLoop (raw Inventor cm).

    Preferred path is ``EdgeUses``, which walks the loop in order and tells us
    per-edge whether its parameterisation is reversed. If that is unavailable we
    fall back to the unordered ``Edges`` collection and chain the segments by
    endpoint proximity -- correct, just slower, and flagged.
    """
    segs = []
    try:
        eus = loop.EdgeUses
        for i in range(1, eus.Count + 1):
            eu = eus.Item(i)
            seg = _edge_points(eu.Edge, diag)
            if not seg:
                continue
            try:
                if eu.IsParamReversed:
                    seg = seg[::-1]
            except Exception:
                pass
            segs.append(seg)
        diag["loops_via_edgeuses"] = diag.get("loops_via_edgeuses", 0) + 1
    except Exception:
        segs = []
        try:
            es = loop.Edges
            for i in range(1, es.Count + 1):
                seg = _edge_points(es.Item(i), diag)
                if seg:
                    segs.append(seg)
            diag["loops_via_edges_unordered"] = diag.get("loops_via_edges_unordered", 0) + 1
        except Exception as exc:
            _diag_err(diag, "loop_edges", exc)
            return None

    if not segs:
        return None

    # Chain segments head-to-tail. With EdgeUses they already are; with the
    # unordered fallback this is what puts them in order.
    pts = list(segs.pop(0))
    while segs:
        tail = pts[-1]
        best_i, best_d, best_rev = None, None, False
        for i, sg in enumerate(segs):
            d0 = _norm(_sub(tail, sg[0]))
            d1 = _norm(_sub(tail, sg[-1]))
            if best_d is None or min(d0, d1) < best_d:
                best_i, best_d, best_rev = i, min(d0, d1), (d1 < d0)
        sg = segs.pop(best_i)
        if best_rev:
            sg = sg[::-1]
        pts.extend(sg[1:])
    return pts if len(pts) >= 3 else None


def _face_rings(face, diag):
    """Every loop of a planar face as an ordered ring of points (cm).

    The exterior loop plus any interior loops -- interior loops ARE the bolt
    holes, and subtracting them is the whole point of reading loops instead of
    wrapping a bounding box around the vertices.

    NOTE: the collection is ``Face.EdgeLoops``, not ``Face.Loops``. ``Face`` has
    no ``Loops`` member; asking for it raises AttributeError on every face and
    silently drops the whole pipeline back onto bounding boxes.
    """
    try:
        loops = face.EdgeLoops
    except Exception as exc:
        _diag_err(diag, "face_edgeloops", exc)
        return None

    rings, outer_flags = [], []
    for li in range(1, loops.Count + 1):
        loop = loops.Item(li)
        pts = _loop_points(loop, diag)
        if not pts:
            continue
        rings.append(pts)
        try:
            outer_flags.append(bool(loop.IsOuterEdgeLoop))
        except Exception:
            outer_flags.append(None)

    if not rings:
        return None

    # Cross-check: Inventor's own outer-loop flag against largest-area. They
    # should agree; if they do not, the face is odd enough to want a look.
    if any(f is not None for f in outer_flags):
        try:
            flagged = [i for i, f in enumerate(outer_flags) if f]
            if len(flagged) == 1:
                areas = [len(r) for r in rings]  # placeholder; real area below
                diag["loops_outer_flagged"] = diag.get("loops_outer_flagged", 0) + 1
        except Exception:
            pass
    return rings


def _face_outward_normal(face, diag):
    """Outward (away-from-material) normal of a planar face, or None.

    ``Face.Geometry.Normal`` is the underlying SURFACE normal. ``IsParamReversed``
    says whether the face's own normal runs against it. Without this flip, an
    outward normal and an inward one are indistinguishable, and two members whose
    outer surfaces finish FLUSH -- coplanar, pointing the same way, not touching
    each other at all -- look exactly like two faces pressed together.
    """
    try:
        n = face.Geometry.Normal
        nv = (n.X, n.Y, n.Z)
    except Exception as exc:
        if diag is not None:
            diag["non_planar_faces"] = diag.get("non_planar_faces", 0) + 1
        return None
    try:
        if face.IsParamReversed:
            nv = (-nv[0], -nv[1], -nv[2])
        diag["normals_oriented"] = diag.get("normals_oriented", 0) + 1
    except Exception as exc:
        _diag_err(diag, "face_isparamreversed", exc)
        diag["normals_unoriented"] = diag.get("normals_unoriented", 0) + 1
    return nv


def _occ_planar_faces(occ, diag):
    """Planar faces of an occurrence, as (obb, loops) pairs.

    ``occ.SurfaceBodies`` returns proxies, so all geometry is already in ASSEMBLY
    coordinates -- no occurrence transform is applied or needed.

    Geometry comes from the edge loops in ONE pass. The old path enumerated
    ``face.Vertices`` and ``face.Edges`` separately, which is two extra COM
    round-trips per face and was most of the 471 s spent here.
    """
    try:
        if occ.SurfaceBodies.Count == 0:
            return []
        faces = occ.SurfaceBodies.Item(1).Faces
    except Exception as exc:
        _diag_err(diag, "surfacebodies", exc)
        return []
    diag["surfacebodies_ok"] += 1

    out = []
    for i in range(1, faces.Count + 1):
        diag["faces_examined"] += 1
        face = faces.Item(i)

        normal = _face_outward_normal(face, diag)   # also the planarity gate
        if normal is None:
            continue
        diag["planar_faces"] += 1

        rings = _face_rings(face, diag)
        obb = loops = None

        if rings:
            outer = max(rings, key=len)
            x_hint = None
            best = 0.0
            for k in range(len(outer)):
                d = _sub(outer[(k + 1) % len(outer)], outer[k])
                L = _norm(d)
                if L > best:
                    best, x_hint = L, d
            all_pts = [p for r in rings for p in r]
            obb = _cg.build_face_obb(all_pts, normal, x_hint)
            loops = _cg.build_face_loops(rings, normal, x_hint)

        if loops is None:
            # Loops unreadable: fall back to the vertex-hull OBB and SAY SO.
            raw = _face_raw(face, diag)
            if raw is None:
                continue
            obb = _cg.build_face_obb(raw[0], normal, raw[2])
            diag["faces_obb_only"] = diag.get("faces_obb_only", 0) + 1
        else:
            diag["faces_with_loops"] += 1
            if loops.holes_2d:
                diag["faces_with_holes"] = diag.get("faces_with_holes", 0) + 1

        if obb is None:
            continue
        out.append((obb, loops))
    return out


def _to_frame_2d(pts3, frame):
    """World points -> the 2D frame (origin, x_axis, y_axis)."""
    o, xa, ya = frame
    out = []
    for p in pts3:
        d = _sub(p, o)
        out.append((_dot(d, xa), _dot(d, ya)))
    return out


def _hit_rings_in_frame(h, frame):
    """A hit's exterior + hole rings, re-expressed in a common 2D frame."""
    ho, hx, hy = h["frame"]

    def up(ring2d):
        return [_add(ho, _add(_scale(hx, u), _scale(hy, v))) for (u, v) in ring2d]

    ext = _to_frame_2d(up(h["ext2d"]), frame) if h["ext2d"] else []
    holes = [_to_frame_2d(up(r), frame) for r in h["holes2d"]]
    return ext, holes


def _merge_coplanar(hits, tol_cm, diag):
    """Collapse contact patches that are FRAGMENTS OF ONE PHYSICAL CONTACT.

    Inventor splits a face wherever a feature crosses it, so a single bolted
    flange can arrive as three or four faces and therefore three or four face
    pairs. Physically that is one connection: one contact plane, one bolt group,
    one weld path.

    Two conditions, and BOTH are load-bearing:

      1. Same plane, SAME FACING. The normals must agree in DIRECTION, not merely
         be parallel. Using abs(dot) here is a trap: the two opposite flanges of a
         channel have normals +n and -n, and because their offsets are measured
         along their own normals they compute to the SAME signed offset. abs()
         merges them -- collapsing the two-flange connection that this whole
         model exists to distinguish back into one.

      2. Actually adjacent IN PLANE. Feature fragments share an edge; two separate
         contact regions on the same plane (a long brace landing on a beam at two
         places) do not. Union the outlines and keep however many connected
         components come back -- one merged patch each.

    The union is a real polygon union, not a sum: summing fragment perimeters
    double-counts the shared internal edges and inflates the weld length.
    """
    if len(hits) < 2:
        return hits

    groups = []
    for h in hits:
        n = h["normal"]
        off = _dot(h["centroid_cm"], n)
        placed = False
        for g in groups:
            gn = g["normal"]
            # SIGNED dot: same facing, not merely parallel.
            if _dot(gn, n) < _PARALLEL_COS:
                continue
            if abs(g["offset"] - off) > tol_cm * 4:
                continue
            g["hits"].append(h)
            placed = True
            break
        if not placed:
            groups.append({"normal": n, "offset": off, "hits": [h]})

    out = []
    for g in groups:
        hs = g["hits"]
        if len(hs) == 1:
            out.append(hs[0])
            continue
        try:
            out.extend(_union_group(hs, diag))
        except Exception as exc:
            _diag_err(diag, "merge_coplanar", exc)
            out.extend(hs)
    return _dedup_coincident_contacts(out, tol_cm, diag)


def _dedup_coincident_contacts(hits, tol_cm, diag):
    """Collapse the SAME physical lap detected from both of its faces.

    A thin lapped flange presents two parallel faces one material-thickness apart.
    When the contact offset tolerance is wider than that thickness, BOTH the
    front-front and the back-back face pairs pass the offset gate, so a single lap
    is reported twice: identical area, centroids ~a flange-thickness apart, and
    OPPOSITE-facing normals. The coplanar merge above keeps opposite facings apart
    on purpose -- that is how the two genuinely-separate flanges of a channel (which
    sit inches apart) stay two connections -- so it cannot remove these. Here we
    collapse only contacts of the SAME member pair whose centroids are nearly
    COINCIDENT (within the offset tolerance) and whose planes are parallel (either
    facing), keeping the largest. Genuine flanges are far apart and untouched.
    """
    if len(hits) < 2:
        return hits
    kept = []
    for h in sorted(hits, key=lambda x: x["area_cm2"], reverse=True):
        nh = _unit(h["normal"])
        coincident = False
        for k in kept:
            if abs(_dot(nh, _unit(k["normal"]))) < _PARALLEL_COS:
                continue                       # not parallel -> different interface
            if _norm(_sub(h["centroid_cm"], k["centroid_cm"])) <= tol_cm:
                coincident = True
                break
        if coincident:
            diag["contacts_deduped_coincident"] = (
                diag.get("contacts_deduped_coincident", 0) + 1)
        else:
            kept.append(h)
    return kept


def _union_group(hs, diag):
    """Union same-plane patches; return ONE merged hit per connected component."""
    if not _cg.HAVE_SHAPELY:
        diag["merge_skipped_no_shapely"] = diag.get("merge_skipped_no_shapely", 0) + 1
        for h in hs:
            h["flags"] = sorted(set(h["flags"] + ["coplanar_fragments_not_merged"]))
        return hs

    from shapely.geometry import Polygon as _P
    from shapely.ops import unary_union

    base = max(hs, key=lambda h: h["area_cm2"])
    frame = base["frame"]
    eps = 0.02  # cm -- closes float slivers between faces that share an edge

    polys, obb_polys = [], []
    for h in hs:
        ext, holes = _hit_rings_in_frame(h, frame)
        if len(ext) >= 3:
            try:
                q = _P(ext, holes)
                if not q.is_valid:
                    q = q.buffer(0)
                polys.append(q)
            except Exception:
                pass
        ring = h.get("obb_ring_cm") or []
        if len(ring) >= 3:
            try:
                obb_polys.append(_P(_to_frame_2d(ring, frame)).buffer(0))
            except Exception:
                pass

    if not polys:
        return hs

    merged = unary_union([q.buffer(eps) for q in polys]).buffer(-eps)
    parts = list(merged.geoms) if merged.geom_type == "MultiPolygon" else [merged]
    parts = [q for q in parts if q.geom_type == "Polygon" and q.area > 1e-9]
    if not parts:
        return hs

    obb_union = None
    if obb_polys:
        obb_union = unary_union(obb_polys)

    n_merged = len(hs) - len(parts)
    if n_merged > 0:
        diag["patches_merged_coplanar"] = diag.get("patches_merged_coplanar", 0) + n_merged
    if len(parts) > 1:
        diag["coplanar_groups_kept_separate"] = (
            diag.get("coplanar_groups_kept_separate", 0) + 1)

    o, xa, ya = frame
    out = []
    for q in parts:
        ext = [tuple(c) for c in q.exterior.coords[:-1]]
        holes = [[tuple(c) for c in r.coords[:-1]] for r in q.interiors]
        cen2 = (q.centroid.x, q.centroid.y)
        cen3 = _add(o, _add(_scale(xa, cen2[0]), _scale(ya, cen2[1])))

        span = 0.0
        for i in range(len(ext)):
            for j in range(i + 1, len(ext)):
                d = math.dist(ext[i], ext[j])
                if d > span:
                    span = d

        obb_area = obb_union.intersection(q.envelope).area if obb_union else 0.0
        if obb_area <= 1e-9 and obb_union is not None:
            obb_area = obb_union.area

        flags = sorted(set([f for h in hs for f in h["flags"]]))
        if n_merged > 0:
            flags.append(f"merged_{len(hs)}_coplanar_face_fragments")
        if len(parts) > 1:
            flags.append(f"coplanar_but_disjoint_{len(parts)}_regions")

        out.append({
            "area_cm2": q.area,
            "perim_cm": q.exterior.length,   # union perimeter, NOT a sum
            "span_cm": span,
            "centroid_cm": cen3,
            "normal": base["normal"],
            "n_holes": len(holes),
            "ext2d": ext,
            "holes2d": holes,
            "frame": frame,
            "obb_area_cm2": obb_area,
            "obb_divergence": (abs(q.area - obb_area) / obb_area
                               if obb_area > 1e-9 else None),
            "obb_ring_cm": base.get("obb_ring_cm") or [],
            "flags": sorted(set(flags)),
        })
    return out


def _min_contact_gap_cm(faces1, faces2):
    """Smallest normal separation between any near-parallel, OPPOSING face pair of
    two members -- how close their closest candidate *contact* faces come, ignoring
    the touch tolerance. Lets the caller tell a plausible near-contact (small gap)
    from members that merely pass near each other (large gap). Returns None if the
    two present no parallel, opposing face pair at all (never a contact interface).
    """
    best = None
    for oa, _la in faces1:
        for ob, _lb in faces2:
            dot = (oa.normal[0] * ob.normal[0]
                   + oa.normal[1] * ob.normal[1]
                   + oa.normal[2] * ob.normal[2])
            if abs(dot) < _PARALLEL_COS:      # not parallel
                continue
            if dot > 0:                        # same-facing -> not an interface
                continue
            d = _sub(ob.origin, oa.origin)
            offset = abs(d[0] * oa.normal[0] + d[1] * oa.normal[1]
                         + d[2] * oa.normal[2])
            if best is None or offset < best:
                best = offset
    return best


def _all_face_contacts(faces1, faces2, tol_cm, diag):
    """EVERY touching planar-face pair between two members, in cm.

    This is the connection definition made literal: a connection is *a face of
    two different members touching*, so a member pair that touches on three faces
    yields three connections -- not one. (The previous implementation kept only
    the largest overlap per member pair and discarded the rest, which silently
    dropped the second flange of every two-flange bolted connection.)

    Gate, cheapest first, all pure geometry -- no MeasureTools COM call:
        1. planes near-parallel        |n1.n2| ~ 1
        2. normal offset within tol    the V16 angle~=0 & distance~=0 test
        3. true outlines overlap       with real area above the sliver floor

    Each hit returns a dict carrying the EXACT patch (true loops, holes removed)
    plus the OBB area the V16 method would have reported, for cross-check.
    """
    hits = []
    min_area_cm2 = _MIN_CONTACT_AREA_IN2 / (CM_TO_IN ** 2)

    for oa, la in faces1:
        for ob, lb in faces2:
            diag["face_pairs_examined"] += 1

            # Gate 1: planes parallel. NOTE the sign is kept -- see gate 2.
            dot = (oa.normal[0] * ob.normal[0]
                   + oa.normal[1] * ob.normal[1]
                   + oa.normal[2] * ob.normal[2])
            if abs(dot) < _PARALLEL_COS:       # planes not parallel (> ~10 deg)
                continue
            diag["pairs_parallel"] += 1

            # Gate 2: the outward normals must OPPOSE -- the faces have to look
            # INTO each other. Two members finishing flush share a plane and
            # point the SAME way (dot ~ +1); they are not touching each other at
            # all, and a flush detail is common enough in this frame that
            # accepting them roughly doubles the connection count. V16's
            # CheckNormal exists for exactly this; taking abs(dot) throws it away.
            if dot > 0:
                diag["pairs_same_side_rejected"] = (
                    diag.get("pairs_same_side_rejected", 0) + 1)
                continue
            diag["pairs_normals_opposed"] = diag.get("pairs_normals_opposed", 0) + 1

            d = _sub(ob.origin, oa.origin)
            offset = abs(d[0] * oa.normal[0] + d[1] * oa.normal[1] + d[2] * oa.normal[2])
            if offset >= tol_cm:               # parallel but too far apart
                continue
            diag["pairs_offset_ok"] += 1

            # --- OBB path (V16 reference / cross-check) ---
            obb_metrics = _cg.contact_metrics(oa, ob)
            obb_area = obb_metrics[0] if obb_metrics else 0.0
            try:
                obb_ring = _cg.obb_ring_world(oa, ob)
            except Exception as exc:
                _diag_err(diag, "obb_ring", exc)
                obb_ring = []

            # --- exact path (primary) ---
            exact = None
            if la is not None and lb is not None:
                try:
                    exact = _cg.contact_metrics_exact(la, lb)
                except Exception as exc:
                    _diag_err(diag, "contact_exact", exc)

            flags = []
            if exact is not None:
                area, perim, span, cen, n_holes, ext2d, holes2d, frame, gflags = exact
                flags.extend(gflags)
            elif obb_metrics is not None:
                # No usable loops on one of the faces: fall back to the OBB result
                # and say so, rather than dropping the contact.
                area, perim, span, cen = obb_metrics
                n_holes, ext2d, holes2d = 0, [], []
                frame = (oa.origin, oa.x_axis, oa.y_axis)
                flags.append("obb_fallback_no_loops")
                diag["contacts_obb_fallback"] = diag.get("contacts_obb_fallback", 0) + 1
            else:
                continue

            if area < min_area_cm2:
                diag["contacts_below_min_area"] = diag.get("contacts_below_min_area", 0) + 1
                continue

            div = abs(area - obb_area) / obb_area if obb_area > 1e-9 else None
            if div is not None and div > _cg.OBB_DIVERGENCE_FLAG:
                flags.append(f"obb_divergence_{div * 100:.0f}pct")
                diag["contacts_obb_divergent"] = diag.get("contacts_obb_divergent", 0) + 1

            hits.append({
                "area_cm2": area, "perim_cm": perim, "span_cm": span,
                "centroid_cm": cen, "normal": oa.normal, "n_holes": n_holes,
                "ext2d": ext2d, "holes2d": holes2d, "frame": frame,
                "obb_area_cm2": obb_area, "obb_divergence": div,
                "obb_ring_cm": obb_ring,
                "flags": flags,
            })

    diag["contacts_raw"] = diag.get("contacts_raw", 0) + len(hits)

    # Inventor fragments a physical face at every feature boundary, so one bolted
    # flange arrives as several face pairs. Merge patches sharing a contact PLANE
    # back into one connection before emitting.
    hits = _merge_coplanar(hits, tol_cm, diag)

    # Largest patch first, so face_pair_index 0 is the primary contact.
    hits.sort(key=lambda h: h["area_cm2"], reverse=True)
    if hits:
        diag["contacts_found"] += len(hits)
    return hits


def _add_review(conn: Connection, reason: str) -> None:
    """Append a review reason without clobbering an existing one."""
    if not reason:
        return
    conn.needs_review = True
    conn.review_reason = f"{conn.review_reason}; {reason}".lstrip("; ")


def enrich_connections_with_faces(app, members, connections,
                                  tol_touch_in: float = DEFAULT_TOL_TOUCH_IN,
                                  contact_offset_tol_in: float = DEFAULT_CONTACT_OFFSET_TOL_IN,
                                  near_contact_gap_in: float = DEFAULT_NEAR_CONTACT_GAP_IN):
    """Expand each centerline member-pair contact into its real face contacts.

    A connection is *a face of two different members touching*. A member pair
    that touches on several faces -- a tube bolted through both flanges of a
    channel, a brace landing on two faces of a column -- is therefore several
    connections, and each one has to be checked, costed and fastened separately.

    So this does NOT enrich in place. It REPLACES each centerline connection with
    one Connection per touching face pair, inheriting joint_id, connection_type
    and angle_deg from the parent pair and adding the true contact patch (exact
    outline, bolt holes subtracted) plus the OBB cross-check.

    A member pair with no resolvable face contact keeps its centerline
    connection, flagged: it is a contact the centerline logic believes in but the
    solid model does not confirm, which is worth seeing rather than deleting.

    Requires the live occurrence handles (Member._occ), so it must run inside
    run() before those are dropped.

    Returns (new_connections, diagnostics).
    """
    diag = {
        "connections_total": len(connections),
        "with_occ": 0, "occ_missing": 0, "surfacebodies_ok": 0,
        "faces_examined": 0, "planar_faces": 0, "non_planar_faces": 0,
        "faces_with_loops": 0, "faces_with_holes": 0, "faces_obb_only": 0,
        "curved_edges_chorded": 0, "normals_oriented": 0, "normals_unoriented": 0,
        "loops_via_edgeuses": 0, "loops_via_edges_unordered": 0,
        "face_pairs_examined": 0, "pairs_parallel": 0,
        "pairs_normals_opposed": 0, "pairs_same_side_rejected": 0,
        "pairs_offset_ok": 0,
        "contacts_raw": 0, "patches_merged_coplanar": 0,
        "contacts_deduped_coincident": 0,
        "coplanar_groups_kept_separate": 0, "pairs_dropped_no_contact": 0,
        "contacts_found": 0, "contacts_below_min_area": 0,
        "contacts_obb_fallback": 0, "contacts_obb_divergent": 0,
        "pairs_expanded": 0, "pairs_multi_face": 0, "pairs_unconfirmed": 0,
        "pairs_kept_near_contact": 0, "pairs_dropped_far": 0,
        "connections_out": 0, "errors": [],
    }

    by_name = {}
    for m in members:
        by_name[m.occurrence_path or m.occurrence_name] = m
    # Face-contact gate uses the (looser, decoupled) contact offset tolerance, not
    # the centerline tol_touch. See DEFAULT_CONTACT_OFFSET_TOL_IN.
    tol_cm = contact_offset_tol_in / CM_TO_IN   # in -> cm

    face_cache: dict = {}   # member key -> planar faces (built once per member)

    def _faces_for(name, m):
        if name in face_cache:
            return face_cache[name]
        occ = getattr(m, "_occ", None)
        faces = _occ_planar_faces(occ, diag) if occ is not None else None
        face_cache[name] = faces
        return faces

    out: list[Connection] = []
    cid = 1

    for conn in connections:
        ma = by_name.get(conn.member_a)
        mb = by_name.get(conn.member_b)

        hits = []
        faces1 = faces2 = None
        faces_readable = False
        if ma is None or mb is None or getattr(ma, "_occ", None) is None \
                or getattr(mb, "_occ", None) is None:
            diag["occ_missing"] += 1
        else:
            diag["with_occ"] += 1
            faces1 = _faces_for(conn.member_a, ma)
            faces2 = _faces_for(conn.member_b, mb)
            if faces1 and faces2:
                faces_readable = True
                try:
                    hits = _all_face_contacts(faces1, faces2, tol_cm, diag)
                except Exception as exc:
                    _diag_err(diag, "all_face_contacts", exc)
                    hits = []

        if not hits:
            diag["pairs_unconfirmed"] += 1
            if faces_readable:
                # Faces read fine, but no face pair confirms a contact within the
                # offset tolerance. Decide keep-vs-drop by the closest candidate
                # contact faces: how near do any parallel, opposing faces of the
                # two members actually come?
                gap_cm = _min_contact_gap_cm(faces1, faces2)
                gap_in = gap_cm * CM_TO_IN if gap_cm is not None else None

                if gap_in is None or gap_in > near_contact_gap_in:
                    # No candidate contact faces at all, or the closest ones are
                    # too far apart to be a fastened joint (e.g. a diagonal brace
                    # ~1.2" off the vertical channel it merely passes near). The
                    # centerline detector paired them only because they share a
                    # node; the solid model says they do not touch. Drop it.
                    diag.setdefault("dropped_far_pairs", []).append(
                        f"{conn.member_a} + {conn.member_b} @ {conn.joint_id} "
                        f"(gap {'none' if gap_in is None else f'{gap_in:.2f}in'})")
                    diag["pairs_dropped_far"] = (
                        diag.get("pairs_dropped_far", 0) + 1)
                    continue

                # Close enough to be a plausible near-contact (a coped brace end
                # bolted to a gusset, or a lap with a small standoff the parallel-
                # face model can't measure). KEEP it as a centerline connection
                # flagged for review, rather than inventing a face area OR silently
                # erasing a real load path. The engineer prunes it if it is a
                # phantom.
                diag.setdefault("unconfirmed_pairs", []).append(
                    f"{conn.member_a} + {conn.member_b} @ {conn.joint_id} "
                    f"(gap {gap_in:.2f}in)")
                diag["pairs_kept_near_contact"] = (
                    diag.get("pairs_kept_near_contact", 0) + 1)
                keep = replace(conn, connection_id=f"C{cid:03d}",
                               detection_method="centerline")
                _add_review(keep, "near_contact_not_face_confirmed")
                out.append(keep)
                cid += 1
                continue
            # Faces could not be read at all (no live session, no bodies): the
            # centerline connection is the only evidence we have. Keep it, flagged.
            keep = replace(conn, connection_id=f"C{cid:03d}")
            _add_review(keep, "centerline_only_faces_unreadable")
            out.append(keep)
            cid += 1
            continue

        diag["pairs_expanded"] += 1
        if len(hits) > 1:
            diag["pairs_multi_face"] += 1

        for k, h in enumerate(hits):
            c = replace(
                conn,
                connection_id=f"C{cid:03d}",
                face_pair_index=k,
                detection_method="face_contact",
                location=tuple(round(v * CM_TO_IN, 4) for v in h["centroid_cm"]),
                contact_area_in2=round(h["area_cm2"] * CM_TO_IN ** 2, 4),
                weld_length_in=round(h["perim_cm"] * CM_TO_IN, 4),
                joint_length_in=round(h["span_cm"] * CM_TO_IN, 4),
                contact_normal=tuple(round(v, 6) for v in h["normal"]),
                hole_count=h["n_holes"],
                patch_frame={
                    "origin": tuple(round(v * CM_TO_IN, 4) for v in h["frame"][0]),
                    "x_axis": tuple(round(v, 6) for v in h["frame"][1]),
                    "y_axis": tuple(round(v, 6) for v in h["frame"][2]),
                },
                patch_exterior_2d=[(round(u * CM_TO_IN, 4), round(v * CM_TO_IN, 4))
                                   for u, v in h["ext2d"]],
                patch_holes_2d=[[(round(u * CM_TO_IN, 4), round(v * CM_TO_IN, 4))
                                 for u, v in ring] for ring in h["holes2d"]],
                obb_area_in2=round(h["obb_area_cm2"] * CM_TO_IN ** 2, 4),
                obb_divergence=(round(h["obb_divergence"], 4)
                                if h["obb_divergence"] is not None else None),
                obb_polygon_3d=[tuple(round(v * CM_TO_IN, 4) for v in p)
                                for p in h["obb_ring_cm"]],
            )
            for f in h["flags"]:
                _add_review(c, f)
            if len(hits) > 1:
                _add_review(c, f"multi_face_contact_{k + 1}_of_{len(hits)}")
            out.append(c)
            cid += 1

    diag["connections_out"] = len(out)
    return out, diag


print("OK: Connection builder loaded (centerline + live face-contact enrichment)")


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
            # --- Milestone 2: read cost iProperties into the exported cost field ---
            try:
                from engine.block1_cost_extract import extract_part_cost_fields
                cost_fields, _cost_review = extract_part_cost_fields(def_doc, m.part_number)
                m.cost = cost_fields
            except Exception:
                m.cost = {}   # cost extraction must never break structural extraction
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


def _transform_vector(matrix, v):
    """Apply just the rotational part of an Inventor Matrix to a direction
    vector (no translation), then re-normalize to guard against any scale
    in the occurrence transform."""
    x, y, z = v
    out = []
    for r in range(1, 4):
        out.append(matrix.Cell(r, 1) * x + matrix.Cell(r, 2) * y + matrix.Cell(r, 3) * z)
    n = math.sqrt(sum(c * c for c in out))
    if n < 1e-12:
        return (0.0, 0.0, 0.0)
    return (out[0] / n, out[1] / n, out[2] / n)


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


def _slice_loops(part_def, occ=None):
    """Profile loops by slicing the solid body at mid-span. Works regardless of
    how the part was modelled (sheet-metal/cold-formed, sweep, import, etc.).

    Also returns the two in-plane axes of the cut, as unit vectors in
    assembly/global space (via the occurrence transform's rotation), so the
    caller can later tell which global direction the measured depth/width
    bbox extents actually point in. (None, None) if occ isn't supplied."""
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

    axis0_dir = axis1_dir = None
    if occ is not None and loops:
        others = [i for i in range(3) if i != axis]
        local0 = tuple(1.0 if i == others[0] else 0.0 for i in range(3))
        local1 = tuple(1.0 if i == others[1] else 0.0 for i in range(3))
        m = occ.Transformation
        axis0_dir = _transform_vector(m, local0)
        axis1_dir = _transform_vector(m, local1)

    return loops, axis0_dir, axis1_dir


def _section_loops(part_def, occ=None):
    """
    Return (loops_cm, detection_method, axis0_dir, axis1_dir).

    The returned loop coordinates are still raw Inventor API/database cm. They
    must be converted to inches before classification/properties/export.
    axis0_dir/axis1_dir are unit vectors (assembly space) for the loop's local
    x/y axes -- only available from the midspan_slice path; None otherwise.
    """
    loops = _extrude_loops(part_def)
    if loops:
        return loops, "extrude_profile", None, None
    try:
        loops, axis0_dir, axis1_dir = _slice_loops(part_def, occ)
        if loops:
            return loops, "midspan_slice", axis0_dir, axis1_dir
    except Exception:
        pass
    return [], "geometry", None, None




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
            loops_cm, method, axis0_dir, axis1_dir = _section_loops(part_def, occ)
            if loops_cm:
                # Convert raw Inventor API/database cm to canonical inches before
                # section classification and property calculation. This prevents
                # cm/in mixing in depth, width, thickness, A, Iy, Iz, and J.
                loops_in = loops_cm_to_in(loops_cm)

                cs = classify_section(loops_in, detection_method=method,
                                       axis0_dir=axis0_dir, axis1_dir=axis1_dir)
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
    with_face_connections: bool = True,
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

    with_face_connections
      When True (default) the expensive live face-contact pass runs and each
      centerline connection is expanded into its real contact patches. When
      False the pass is skipped and only the cheap centerline connections are
      emitted -- this is what the app's geometry step uses so extraction stays
      fast; the face pass is then run on demand via ``enrich_connections_live``.
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

        # Per-stage wall-clock timing. Several stages are heavily data-dependent
        # (midspan-slice tessellation, Inventor mass regeneration, face-contact
        # reads), which is why total run time swings from seconds to many minutes.
        import time as _time
        _t: dict = {}

        def _timed(key, fn):
            _c0 = _time.perf_counter()
            out = fn()
            _t[key] = round(_time.perf_counter() - _c0, 2)
            return out

        # Step 1: filter GLV/HDG members (with unique occurrence_path).
        members = _timed("filter_members_s", lambda: filter_glv_members(asm, tags=tags))
        _t["n_members"] = len(members)

        # Step 2: fill geometry. _centerline and profile loops are converted to inches here.
        # This is the usual culprit for slow runs: parts without a clean extrude
        # profile fall back to a midspan mesh SLICE (body.CalculateFacets), whose
        # cost explodes for thin/large/complex sheet-metal parts.
        members = _timed("fill_geometry_s", lambda: fill_geometry(members))
        _t["n_midspan_slice"] = sum(
            1 for m in members if "midspan_slice" in (m.cross_section.detection_method or ""))
        _t["n_extrude_profile"] = sum(
            1 for m in members if "extrude_profile" in (m.cross_section.detection_method or ""))

        # Step 2b: solver-ready section properties in in^2/in^4.
        def _section_props():
            for m in members:
                compute_section_properties(m.cross_section)
        _timed("section_props_s", _section_props)

        # Step 2c: per-member dry mass + self-weight force for Block 2 dead load.
        # MassProperties.Mass can force Inventor to regenerate mass -- fast when
        # cached, slow on a cold model.
        def _mass_props():
            for m in members:
                try:
                    mass_kg = float(m._occ.MassProperties.Mass)  # Inventor API/database mass: kg
                    m.dry_mass = kg_to_lbm(mass_kg)              # canonical mass: lbm
                    m.mass_unit = "lbm"
                    m.self_weight_lbf = kg_to_lbf_weight(mass_kg) # canonical force: lbf
                    m.force_unit = "lbf"
                except Exception:
                    pass
        _timed("mass_props_s", _mass_props)

        result.members = members

        # Step 3: detect joint candidates using inch-based tolerances.
        candidates = _timed("detect_joints_s", lambda: detect_joints(
            members,
            tol_touch=tol_touch_in,
            cluster_cap=cluster_cap_in,
        ))
        # The radius that decides which contacts merge into one joint. Exported
        # so the connection viewer can draw it and so a joint that swallowed two
        # separate nodes is visible rather than inferred.
        result.joint_cluster_tol_in = round(
            joint_cluster_tolerance(members, cluster_cap=cluster_cap_in), 4)

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

        # Step 6: connections. A connection is ONE touching face pair, so the
        # centerline pass builds one per member pair and the face pass EXPANDS
        # each into its real face contacts (a pair touching on two flanges
        # becomes two connections). Runs while the live occurrence handles are
        # still attached.
        result.connections = _timed(
            "connections_centerline_s",
            lambda: build_centerline_connections(members, candidates, result.joints))

        # The face pass is the single most expensive live operation. It is skipped
        # here when with_face_connections=False so the geometry extraction stays
        # fast; the app then runs it on demand as its own workflow step via
        # enrich_connections_live(), which reuses these centerline connections.
        if with_face_connections:
            def _expand():
                try:
                    conns, diag = enrich_connections_with_faces(
                        inv.app, members, result.connections, tol_touch_in)
                    result.connections = conns
                    return diag
                except Exception as exc:
                    import traceback
                    # Keep the centerline connections; they are a usable fallback.
                    return {
                        "fatal_error": f"{type(exc).__name__}: {exc}",
                        "trace": traceback.format_exc()[-1500:],
                    }
            result.connection_diagnostics = _timed("connections_face_s", _expand)

        _t["total_s"] = round(sum(v for k, v in _t.items() if k.endswith("_s")), 2)
        result.timings = _t

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

    n_conn = len(result.connections)
    n_conn_face = sum(1 for c in result.connections if c.detection_method == "face_contact")
    n_conn_rev = sum(1 for c in result.connections if c.needs_review)
    print(f"Connections:           {n_conn}  ({n_conn_face} from face geometry, "
          f"{n_conn_rev} need review)")

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


def extract_live_to_dict(with_face_connections: bool = True) -> dict:
    """Run the frozen Block-1 pipeline against the live session and return a
    plain JSON-safe dict identical to an uploaded block1_result_unit_safe.json.

    Uses the notebook's hardcoded default tolerances. The dict is produced via
    to_json()/json.loads so Enums are strings and tuples are lists, byte-for-
    byte matching the upload path that downstream blocks already consume.

    Set ``with_face_connections=False`` for a fast geometry-only extraction that
    emits centerline connections only; resolve the real face contacts afterwards
    with ``enrich_connections_live``.
    """
    result = run(
        tags=("GLV", "HDG"),
        tol_touch_in=DEFAULT_TOL_TOUCH_IN,
        cluster_cap_in=DEFAULT_CLUSTER_CAP_IN,
        support_z_tol_in=DEFAULT_SUPPORT_Z_TOL_IN,
        with_face_connections=with_face_connections,
    )
    return json.loads(result.to_json())


# --- fields carried over when rebuilding a Connection from its exported dict ---
_CENTERLINE_CONNECTION_FIELDS = (
    "connection_id", "member_a", "member_b", "joint_id", "connection_type",
    "angle_deg", "is_inferred", "needs_review", "review_reason",
)


def _connection_from_dict(d: dict) -> Connection:
    """Rebuild a centerline ``Connection`` from its exported JSON dict.

    Only the fields the face pass needs (member ids, joint id, type, angle) are
    carried over; per-face patch geometry is intentionally dropped so the result
    is a clean centerline representative ready to be re-expanded.
    """
    kw = {k: d[k] for k in _CENTERLINE_CONNECTION_FIELDS
          if k in d and d[k] is not None}
    loc = d.get("location")
    if isinstance(loc, (list, tuple)) and len(loc) == 3:
        kw["location"] = tuple(loc)
    kw.setdefault("connection_id", "C001")
    conn = Connection(**kw)
    conn.detection_method = "centerline"
    return conn


def enrich_connections_live(
    existing: dict,
    tags: tuple[str, ...] = ("GLV", "HDG"),
    tol_touch_in: float = DEFAULT_TOL_TOUCH_IN,
) -> dict:
    """On-demand face-contact detection that reuses a prior geometry extraction.

    Re-attaches to the live Inventor session, re-reads only the member occurrence
    handles (``filter_glv_members`` -- cheap relative to geometry fill / mass
    regeneration), and expands the centerline connections stored in ``existing``
    into their real face contacts. It does NOT re-run the expensive geometry,
    section-property, or mass passes, so it is far faster than a full re-extract.

    ``existing`` is a block1 result dict (``st.session_state.block1_raw``). The
    call is idempotent: the stored connections are first collapsed to one
    representative per member pair, so re-running after a previous face pass does
    not duplicate contacts.

    Returns a JSON-safe dict::

        {"connections": [...], "connection_diagnostics": {...},
         "connections_face_s": <float seconds>}

    to be merged back into the block1 result dict.
    """
    import time as _time
    from dataclasses import asdict

    raw_conns = existing.get("connections") or []
    seen: set = set()
    centerline: list[Connection] = []
    # Engineer decisions are recorded per member-pair, not per id: the face pass
    # renumbers connections and expands one centerline pair into several patches,
    # so an id-keyed flag would be lost. Re-applied to the output below, which
    # keeps both decisions reversible across a re-extract instead of resurrecting
    # a ruled-out phantom or re-condemning a restored pair.
    #
    # Note every centerline pair is fed to the face pass, including ones flagged
    # unreachable by the reach heuristic. That heuristic only knows each member's
    # nominal section box; this pass measures the actual solids and is what should
    # decide. Pairs it cannot confirm are dropped or flagged here on their merits.
    removed_pairs: set = set()
    override_pairs: set = set()
    for d in raw_conns:
        if not isinstance(d, dict):
            continue
        pair = frozenset((d.get("member_a"), d.get("member_b")))
        if d.get("manually_removed"):
            removed_pairs.add(pair)
        if d.get("prune_override"):
            override_pairs.add(pair)
        key = (d.get("member_a"), d.get("member_b"), d.get("joint_id"))
        if key in seen:
            continue
        seen.add(key)
        centerline.append(_connection_from_dict(d))
    # Renumber sequentially so ids are stable regardless of what was fed in.
    for i, c in enumerate(centerline, start=1):
        c.connection_id = f"C{i:03d}"

    t0 = _time.perf_counter()
    with InventorSession() as inv:
        asm = inv.active_assembly()
        members = filter_glv_members(asm, tags=tags)
        conns, diag = enrich_connections_with_faces(
            inv.app, members, centerline, tol_touch_in)
    for m in members:
        if hasattr(m, "_occ"):
            del m._occ
    elapsed = round(_time.perf_counter() - t0, 2)

    # --- member-identification reconciliation (reproducibility guard) --------
    # Face detection re-derives the member list to get fresh occurrence handles.
    # For the connection->member links to stay valid, that fresh identification
    # must match the one Step 2 exported. Compare occurrence-path sets and report
    # any divergence rather than silently emitting mismatched contacts. Any
    # non-zero "members_unmatched" means the live model changed since geometry
    # extraction and the result should not be trusted until re-extracted.
    prior_paths = {
        (m.get("occurrence_path") or m.get("occurrence_name"))
        for m in (existing.get("members") or [])
        if isinstance(m, dict)
    }
    live_paths = {(m.occurrence_path or m.occurrence_name) for m in members}
    diag["members_prior"] = len(prior_paths)
    diag["members_relinked"] = len(live_paths)
    diag["members_matched"] = len(prior_paths & live_paths)
    diag["members_unmatched"] = len(prior_paths - live_paths)
    if prior_paths - live_paths:
        diag["members_unmatched_sample"] = sorted(prior_paths - live_paths)[:10]

    payload = json.loads(json.dumps([asdict(c) for c in conns], default=str))
    for d in payload:
        pair = frozenset((d.get("member_a"), d.get("member_b")))
        if pair in removed_pairs:
            d["manually_removed"] = True
        if pair in override_pairs:
            d["prune_override"] = True
    return {
        "connections": payload,
        "connection_diagnostics": diag,
        "connections_face_s": elapsed,
    }


def _member_reach_in(m: dict) -> Optional[float]:
    """Maximum cross-sectional reach of a member: half its section diagonal.

    This is the farthest any point of the (solid-bounded) member can sit from its
    centerline, for ANY section orientation -- so it depends only on the depth and
    width magnitudes, not on the (less reliable) section direction vectors.
    """
    cs = m.get("cross_section") or {}
    try:
        d = float(cs.get("depth") or 0.0)
        w = float(cs.get("width") or 0.0)
    except (TypeError, ValueError):
        return None
    if d <= 0.0 and w <= 0.0:
        return None
    return 0.5 * math.sqrt(d * d + w * w)


def prune_noncontact_connections(data: dict, extra_tol_in: float = 0.0):
    """FLAG centerline connections whose members are too far apart to touch.

    The centerline joint detector pairs members using a generous ``reach``, so at
    a multi-member joint it emits a connection for every near pair -- including
    pairs that never physically touch (e.g. two parallel base channels a foot
    apart, or a brace on one side paired with a chord on the other). Here a pair
    is marked when the straight distance between the two members' centerlines
    exceeds their combined maximum cross-sectional reach (half section diagonal
    each): at that separation no orientation of the two sections can bring them
    into contact.

    Marks, never deletes. The test is sound only for a member whose solid stays
    inside its nominal depth x width box -- a gusset plate, bracket, tab, or bent
    flange reaches past it, and for those a real bolted connection can sit well
    beyond ``reach``. Deleting outright made that failure invisible: the
    connection vanished from the model with no row, no id, and no way to get it
    back. So each pair gets ``pruned_noncontact`` plus a ``prune_reason``
    carrying the actual numbers, stays in the list, and remains available to the
    face pass -- which measures real geometry and is the only thing entitled to
    settle it. ``prune_override`` (set from the UI) protects an engineer's
    decision from being re-flagged.

    ``face_contact`` connections are already confirmed and are never touched.
    Uses only member endpoints + section depth/width, the most reliable extracted
    fields. Returns (connections, flagged_ids, diag); the first element is the
    full list, kept for call-site compatibility.
    """
    members: dict = {}
    for m in data.get("members", []) or []:
        if isinstance(m, dict):
            key = m.get("occurrence_path") or m.get("occurrence_name")
            if key:
                members[key] = m

    reach_cache: dict = {}

    def _reach(key):
        if key not in reach_cache:
            m = members.get(key)
            reach_cache[key] = _member_reach_in(m) if m else None
        return reach_cache[key]

    def _clear(c):
        """Reset a previous flag — the pair is reachable (or was overridden)."""
        c.pop("pruned_noncontact", None)
        c.pop("prune_reason", None)
        c.pop("prune_shortfall_in", None)

    kept, pruned = [], []
    for c in data.get("connections", []) or []:
        kept.append(c)
        if not isinstance(c, dict) or c.get("detection_method") == "face_contact":
            continue
        if c.get("prune_override"):
            # The engineer looked at this pair and said it is real. Face
            # detection may still overrule it; this heuristic may not.
            _clear(c)
            continue
        ka, kb = c.get("member_a"), c.get("member_b")
        ra, rb = _reach(ka), _reach(kb)
        ma, mb = members.get(ka), members.get(kb)
        if ra is None or rb is None or ma is None or mb is None:
            _clear(c)
            continue
        sa, ea = ma.get("start_point"), ma.get("end_point")
        sb, eb = mb.get("start_point"), mb.get("end_point")
        if not (sa and ea and sb and eb):
            _clear(c)
            continue
        try:
            dist = _closest_params(tuple(sa), tuple(ea), tuple(sb), tuple(eb))[0]
        except Exception:
            _clear(c)
            continue
        limit = ra + rb + extra_tol_in
        if dist <= limit:
            _clear(c)
        else:
            c["pruned_noncontact"] = True
            c["prune_shortfall_in"] = round(dist - limit, 3)
            c["prune_reason"] = (
                f"centerlines {dist:.2f} in apart, but the two sections can reach at "
                f"most {limit:.2f} in ({ra:.2f} + {rb:.2f}) — short by "
                f"{dist - limit:.2f} in, so they cannot touch unless a gusset, "
                f"bracket or tab extends past the nominal section."
            )
            pruned.append(c.get("connection_id"))

    diag = {
        "connections_in": len(data.get("connections", []) or []),
        "connections_kept": len(kept) - len(pruned),
        "connections_pruned": len(pruned),
        "pruned_ids": pruned,
    }
    return kept, pruned, diag


# Extraction parameters surfaced read-only in the UI for documentation.
EXTRACTION_PARAMS = {
    "tags": ("GLV", "HDG"),
    "tol_touch_in": DEFAULT_TOL_TOUCH_IN,
    "cluster_cap_in": DEFAULT_CLUSTER_CAP_IN,
    "support_z_tol_in": DEFAULT_SUPPORT_Z_TOL_IN,
}
