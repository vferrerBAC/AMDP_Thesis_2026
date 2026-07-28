"""
Contact geometry for connection detection (pure, no Inventor/COM).
=================================================================

Faithful Python port of the planar-face contact math in the standalone
``JointLocatorV16`` Inventor macro (VB). A *connection* is a face of two
different members touching; these helpers take the two contacting planar faces
(as vertex lists + plane normals in ANY consistent length unit) and return the
overlap metrics the macro produced per contact:

    contact area      <- Abs(SignedArea(clipped))        (unit^2)
    weld length       <- GetPolygonPerimeter(clipped)    (unit)   perimeter of overlap
    joint length      <- GetMaxDistanceInPolygon(clipped)(unit)   longest span of overlap
    centroid          <- ComputePolygonCentroid(clipped) -> world (unit)

Method (mirrors the macro):
  1. Build an oriented bounding box (OBB) in each face's own plane from its
     vertices, using the longest edge as the in-plane x-axis.
  2. Project face B's OBB corners onto face A's plane and express both in A's
     2D frame.
  3. Sutherland-Hodgman clip A's rectangle by B's rectangle.
  4. Area / perimeter / max-span / centroid of the clipped polygon; centroid is
     mapped back to 3D world coordinates in A's plane.

Everything is unit-agnostic: pass inches in, get inches/in^2 out; pass cm in,
get cm/cm^2 out. The caller converts at the Inventor cm boundary.
"""

from __future__ import annotations

import math
from typing import Optional, Sequence, Tuple

Vec3 = Tuple[float, float, float]
Vec2 = Tuple[float, float]


# --------------------------------------------------------------------------- #
# Small 3D vector helpers (match block1._sub/_dot/... but kept local so this
# module has no dependency on the big pipeline file).
# --------------------------------------------------------------------------- #
def _sub(a: Vec3, b: Vec3) -> Vec3:
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _add(a: Vec3, b: Vec3) -> Vec3:
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def _scale(a: Vec3, s: float) -> Vec3:
    return (a[0] * s, a[1] * s, a[2] * s)


def _dot(a: Vec3, b: Vec3) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a: Vec3, b: Vec3) -> Vec3:
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _norm(a: Vec3) -> float:
    return math.sqrt(_dot(a, a))


def _unit(a: Vec3) -> Vec3:
    n = _norm(a)
    return (a[0] / n, a[1] / n, a[2] / n) if n > 1e-12 else (0.0, 0.0, 0.0)


class FaceOBB:
    """Oriented bounding box of one planar face, in world coordinates.

    Mirrors the VB ``FaceOBB`` class: an origin (face centroid), an in-plane
    x/y frame, the plane normal, and the four rectangle corners in 3D.
    """

    __slots__ = ("origin", "x_axis", "y_axis", "normal", "corners")

    def __init__(self, origin: Vec3, x_axis: Vec3, y_axis: Vec3,
                 normal: Vec3, corners):
        self.origin = origin
        self.x_axis = x_axis
        self.y_axis = y_axis
        self.normal = normal
        self.corners = corners


def build_face_obb(vertices: Sequence[Vec3], normal: Vec3,
                   x_hint: Optional[Vec3] = None) -> Optional[FaceOBB]:
    """Build a face OBB from its vertices and plane normal (VB ``BuildFaceOBB``).

    ``x_hint`` is the direction of the face's longest edge; if omitted, the
    first non-degenerate vertex-to-vertex direction is used. Returns None when
    the inputs are degenerate (fewer than 3 vertices, zero-length axes).
    """
    verts = [tuple(float(c) for c in v) for v in vertices]
    if len(verts) < 3:
        return None

    n = _unit(normal)
    if _norm(n) < 1e-9:
        return None

    # Origin = centroid of the vertices.
    sx = sum(v[0] for v in verts)
    sy = sum(v[1] for v in verts)
    sz = sum(v[2] for v in verts)
    count = len(verts)
    origin = (sx / count, sy / count, sz / count)

    # In-plane x-axis: longest edge direction, projected into the plane.
    if x_hint is None or _norm(x_hint) < 1e-9:
        x_hint = None
        for i in range(len(verts)):
            d = _sub(verts[(i + 1) % len(verts)], verts[i])
            if _norm(d) > 1e-9:
                x_hint = d
                break
        if x_hint is None:
            return None

    x_axis = _unit(x_hint)
    y_axis = _unit(_cross(n, x_axis))
    if _norm(y_axis) < 1e-9:
        return None
    # Re-orthogonalize x against the normal/y (VB: xAxis = yAxis x normal).
    x_axis = _unit(_cross(y_axis, n))
    if _norm(x_axis) < 1e-9:
        return None

    min_x = min_y = float("inf")
    max_x = max_y = float("-inf")
    for v in verts:
        vec = _sub(v, origin)
        px = _dot(vec, x_axis)
        py = _dot(vec, y_axis)
        min_x = min(min_x, px)
        max_x = max(max_x, px)
        min_y = min(min_y, py)
        max_y = max(max_y, py)

    corners = [
        _local_to_world(origin, x_axis, y_axis, min_x, min_y),
        _local_to_world(origin, x_axis, y_axis, max_x, min_y),
        _local_to_world(origin, x_axis, y_axis, max_x, max_y),
        _local_to_world(origin, x_axis, y_axis, min_x, max_y),
    ]
    return FaceOBB(origin, x_axis, y_axis, n, corners)


def _local_to_world(origin: Vec3, x_axis: Vec3, y_axis: Vec3,
                    lx: float, ly: float) -> Vec3:
    return _add(origin, _add(_scale(x_axis, lx), _scale(y_axis, ly)))


def _to_2d(points: Sequence[Vec3], ref: FaceOBB, project: bool) -> list:
    """Express world points in ``ref``'s 2D frame (VB ``To2D``).

    When ``project`` is True each point is first dropped onto ref's plane along
    the plane normal (used for the *other* face's corners).
    """
    out = []
    for p in points:
        v = _sub(p, ref.origin)
        if project:
            dist = _dot(v, ref.normal)
            v = (v[0] - dist * ref.normal[0],
                 v[1] - dist * ref.normal[1],
                 v[2] - dist * ref.normal[2])
        out.append((_dot(v, ref.x_axis), _dot(v, ref.y_axis)))
    return out


def _signed_area(poly: Sequence[Vec2]) -> float:
    area = 0.0
    n = len(poly)
    for i in range(n):
        j = (i + 1) % n
        area += poly[i][0] * poly[j][1] - poly[j][0] * poly[i][1]
    return area * 0.5


def _ensure_ccw(poly: list) -> list:
    if _signed_area(poly) < 0:
        poly = list(reversed(poly))
    return poly


def _is_inside_edge(p: Vec2, a: Vec2, b: Vec2) -> bool:
    return (b[0] - a[0]) * (p[1] - a[1]) - (b[1] - a[1]) * (p[0] - a[0]) >= 0


def _edge_intersect(p: Vec2, q: Vec2, a: Vec2, b: Vec2) -> Vec2:
    a1 = q[1] - p[1]
    b1 = p[0] - q[0]
    c1 = a1 * p[0] + b1 * p[1]
    a2 = b[1] - a[1]
    b2 = a[0] - b[0]
    c2 = a2 * a[0] + b2 * a[1]
    det = a1 * b2 - a2 * b1
    if abs(det) < 1e-10:
        return (p[0], p[1])
    return ((b2 * c1 - b1 * c2) / det, (a1 * c2 - a2 * c1) / det)


def _clip_polygon(subject: list, clipper: list) -> list:
    """Sutherland-Hodgman clip of ``subject`` by convex CCW ``clipper``."""
    output = list(subject)
    for i in range(len(clipper)):
        if not output:
            return output
        a = clipper[i]
        b = clipper[(i + 1) % len(clipper)]
        input_poly = output
        output = []
        for j in range(len(input_poly)):
            p = input_poly[j]
            q = input_poly[(j + 1) % len(input_poly)]
            p_in = _is_inside_edge(p, a, b)
            q_in = _is_inside_edge(q, a, b)
            if p_in and q_in:
                output.append(q)
            elif p_in:
                output.append(_edge_intersect(p, q, a, b))
            elif q_in:
                output.append(_edge_intersect(p, q, a, b))
                output.append(q)
    return output


def _polygon_perimeter(poly: Sequence[Vec2]) -> float:
    if len(poly) < 2:
        return 0.0
    perim = 0.0
    for i in range(len(poly)):
        p1 = poly[i]
        p2 = poly[(i + 1) % len(poly)]
        perim += math.hypot(p2[0] - p1[0], p2[1] - p1[1])
    return perim


def _max_distance(poly: Sequence[Vec2]) -> float:
    if len(poly) < 2:
        return 0.0
    max_d = 0.0
    for i in range(len(poly)):
        for j in range(i + 1, len(poly)):
            d = math.hypot(poly[i][0] - poly[j][0], poly[i][1] - poly[j][1])
            if d > max_d:
                max_d = d
    return max_d


def _polygon_centroid(poly: Sequence[Vec2]) -> Optional[Vec2]:
    if not poly:
        return None
    if len(poly) == 1:
        return poly[0]
    if len(poly) == 2:
        return ((poly[0][0] + poly[1][0]) / 2.0, (poly[0][1] + poly[1][1]) / 2.0)
    cx = cy = area = 0.0
    for i in range(len(poly)):
        j = (i + 1) % len(poly)
        cross = poly[i][0] * poly[j][1] - poly[j][0] * poly[i][1]
        area += cross
        cx += (poly[i][0] + poly[j][0]) * cross
        cy += (poly[i][1] + poly[j][1]) * cross
    area *= 0.5
    if abs(area) < 1e-10:
        return ((poly[0][0] + poly[1][0]) / 2.0, (poly[0][1] + poly[1][1]) / 2.0)
    return (cx / (6.0 * area), cy / (6.0 * area))


def contact_metrics(obb_a: FaceOBB, obb_b: FaceOBB):
    """Overlap metrics for two planar-face OBBs, computed in A's plane.

    Returns ``(area, perimeter, max_span, centroid_world)`` in the input length
    unit (area in unit^2), or None when the faces do not overlap. Mirrors the
    combination of the VB ``GetContactArea`` / ``GetWeldedJointLength`` /
    ``GetJointLength`` / ``GetContactCentroid`` functions, which all share the
    same clip.
    """
    if obb_a is None or obb_b is None:
        return None

    poly1 = _ensure_ccw(_to_2d(obb_a.corners, obb_a, project=False))
    poly2 = _ensure_ccw(_to_2d(obb_b.corners, obb_a, project=True))

    clipped = _clip_polygon(poly1, poly2)
    if len(clipped) < 3:
        return None

    area = abs(_signed_area(clipped))
    if area < 1e-9:
        return None

    perimeter = _polygon_perimeter(clipped)
    max_span = _max_distance(clipped)
    centroid_2d = _polygon_centroid(clipped)
    if centroid_2d is None:
        return None
    centroid_world = _local_to_world(
        obb_a.origin, obb_a.x_axis, obb_a.y_axis, centroid_2d[0], centroid_2d[1])

    return area, perimeter, max_span, centroid_world


def _smoke_test() -> None:
    """Two unit squares in the z=0 plane, offset so they overlap on a 0.5x1 rect."""
    a = build_face_obb(
        [(0, 0, 0), (1, 0, 0), (1, 1, 0), (0, 1, 0)], (0, 0, 1))
    b = build_face_obb(
        [(0.5, 0, 0), (1.5, 0, 0), (1.5, 1, 0), (0.5, 1, 0)], (0, 0, 1))
    area, perim, span, cen = contact_metrics(a, b)
    assert abs(area - 0.5) < 1e-6, area
    assert abs(perim - 3.0) < 1e-6, perim              # 2*(0.5+1)
    assert abs(span - math.hypot(0.5, 1.0)) < 1e-6, span
    assert abs(cen[0] - 0.75) < 1e-6 and abs(cen[1] - 0.5) < 1e-6, cen
    # Non-overlapping squares -> None.
    c = build_face_obb(
        [(5, 5, 0), (6, 5, 0), (6, 6, 0), (5, 6, 0)], (0, 0, 1))
    assert contact_metrics(a, c) is None


if __name__ == "__main__":
    _smoke_test()
    print("connection_geometry smoke test passed")


# =========================================================================== #
# EXACT CONTACT PATCHES  (added for the connection-based IJET model)
# =========================================================================== #
# The OBB path above is a faithful port of JointLocatorV16 and is RETAINED as an
# automatic cross-check. It is not, however, correct for real parts: it clips a
# rectangle against a rectangle, so any coped tube end, notched flange, mitred
# brace, or bolt hole makes its contact area an OVERESTIMATE -- and an
# overestimated bearing area makes a capacity check unconservative.
#
# The functions below clip the TRUE face outline (exterior loop + interior loops,
# i.e. bolt holes) instead. block1 runs both and flags any connection where they
# disagree by more than OBB_DIVERGENCE_FLAG. Same pattern as the Block 1
# section-property cross-check: exact result primary, closed form as validator,
# divergence -> needs_review.

OBB_DIVERGENCE_FLAG = 0.20

try:
    from shapely.geometry import Polygon as _ShPoly
    from shapely.validation import make_valid as _shp_make_valid
    HAVE_SHAPELY = True
except Exception:                                     # pragma: no cover
    HAVE_SHAPELY = False


class FaceLoops:
    """A planar face as its true boundary: one exterior ring + N hole rings,
    expressed in the face's own 2D frame."""

    __slots__ = ("origin", "normal", "x_axis", "y_axis", "exterior_2d",
                 "holes_2d", "flags")

    def __init__(self, origin, normal, x_axis, y_axis, exterior_2d, holes_2d,
                 flags=None):
        self.origin = origin
        self.normal = normal
        self.x_axis = x_axis
        self.y_axis = y_axis
        self.exterior_2d = exterior_2d
        self.holes_2d = holes_2d
        self.flags = flags or []


def _ring_area_3d(ring, normal) -> float:
    """Area of a planar 3D ring via the projected shoelace."""
    n = len(ring)
    if n < 3:
        return 0.0
    cx = sum(p[0] for p in ring) / n
    cy = sum(p[1] for p in ring) / n
    cz = sum(p[2] for p in ring) / n
    tx = ty = tz = 0.0
    for i in range(n):
        a = (ring[i][0] - cx, ring[i][1] - cy, ring[i][2] - cz)
        b = (ring[(i + 1) % n][0] - cx, ring[(i + 1) % n][1] - cy,
             ring[(i + 1) % n][2] - cz)
        c = _cross(a, b)
        tx += c[0]; ty += c[1]; tz += c[2]
    return abs(_dot((tx, ty, tz), normal)) * 0.5


def build_face_loops(rings: Sequence[Sequence[Vec3]], normal: Vec3,
                     x_hint: Optional[Vec3] = None) -> Optional["FaceLoops"]:
    """Assemble a FaceLoops from a face's loop rings (largest = exterior)."""
    rings = [r for r in rings if r is not None and len(r) >= 3]
    if not rings:
        return None
    n = _unit(normal)
    if _norm(n) < 1e-9:
        return None

    order = sorted(range(len(rings)), key=lambda i: _ring_area_3d(rings[i], n),
                   reverse=True)
    ext3 = list(rings[order[0]])
    holes3 = [list(rings[i]) for i in order[1:]]

    k = len(ext3)
    origin = (sum(p[0] for p in ext3) / k,
              sum(p[1] for p in ext3) / k,
              sum(p[2] for p in ext3) / k)

    # In-plane x-axis: caller's longest-edge hint, else the longest exterior span.
    xa = x_hint
    if xa is None:
        best, xa = 0.0, None
        for i in range(k):
            d = _sub(ext3[(i + 1) % k], ext3[i])
            L = _norm(d)
            if L > best:
                best, xa = L, d
    if xa is None:
        return None
    xa = _sub(xa, _scale(n, _dot(xa, n)))          # project into the plane
    if _norm(xa) < 1e-9:
        return None
    xa = _unit(xa)
    ya = _unit(_cross(n, xa))

    def to2d(ring):
        out = []
        for p in ring:
            v = _sub(p, origin)
            out.append((_dot(v, xa), _dot(v, ya)))
        return out

    return FaceLoops(origin, n, xa, ya, to2d(ext3), [to2d(h) for h in holes3])


def _project_loops_into(src: "FaceLoops", ref: "FaceLoops"):
    """Express src's rings in ref's 2D frame (drops the out-of-plane component)."""
    def to2d(ring2d):
        out = []
        for u, v in ring2d:
            p3 = _add(src.origin,
                      _add(_scale(src.x_axis, u), _scale(src.y_axis, v)))
            d = _sub(p3, ref.origin)
            out.append((_dot(d, ref.x_axis), _dot(d, ref.y_axis)))
        return out
    return to2d(src.exterior_2d), [to2d(h) for h in src.holes_2d]


def _is_convex(poly) -> bool:
    n = len(poly)
    if n < 4:
        return True
    pos = neg = False
    for i in range(n):
        ax, ay = poly[i]
        bx, by = poly[(i + 1) % n]
        cx, cy = poly[(i + 2) % n]
        cr = (bx - ax) * (cy - by) - (by - ay) * (cx - bx)
        if cr > 1e-9:
            pos = True
        elif cr < -1e-9:
            neg = True
    return not (pos and neg)


def clip_loops(ext_a, holes_a, ext_b, holes_b):
    """Intersect two face outlines already in the SAME 2D frame.
    Returns (exterior_ring, hole_rings, flags)."""
    flags = []
    if HAVE_SHAPELY:
        try:
            pa, pb = _ShPoly(ext_a, holes_a), _ShPoly(ext_b, holes_b)
            if not pa.is_valid:
                pa = _shp_make_valid(pa)
            if not pb.is_valid:
                pb = _shp_make_valid(pb)
            inter = pa.intersection(pb)
            if inter.is_empty:
                return [], [], flags
            if inter.geom_type == "MultiPolygon":
                parts = sorted(inter.geoms, key=lambda g: g.area, reverse=True)
                flags.append(f"contact_patch_disjoint_{len(parts)}_regions")
                inter = parts[0]
            if inter.geom_type != "Polygon":
                return [], [], flags
            ext = [tuple(c) for c in inter.exterior.coords[:-1]]
            holes = [[tuple(c) for c in r.coords[:-1]] for r in inter.interiors]
            return ext, holes, flags
        except Exception as exc:                       # pragma: no cover
            flags.append(f"shapely_clip_failed_{type(exc).__name__}")

    # Fallback: Sutherland-Hodgman. Convex clipper only, cannot carry holes.
    if not _is_convex(ext_a) or not _is_convex(ext_b):
        flags.append("nonconvex_face_no_shapely")
    if holes_a or holes_b:
        flags.append("holes_ignored_no_shapely")
    a = _ensure_ccw(list(ext_a))
    b = _ensure_ccw(list(ext_b))
    return _clip_polygon(a, b), [], flags


def patch_props(ext, holes):
    """Area / centroid / exterior perimeter / max span of a clipped patch."""
    if len(ext) < 3:
        return 0.0, (0.0, 0.0), 0.0, 0.0
    a_ext = abs(_signed_area(ext))
    cen = _polygon_centroid(ext) or (0.0, 0.0)
    mx, my, m = cen[0] * a_ext, cen[1] * a_ext, a_ext
    for h in holes:
        ah = abs(_signed_area(h))
        hc = _polygon_centroid(h) or (0.0, 0.0)
        mx -= hc[0] * ah
        my -= hc[1] * ah
        m -= ah
    centroid = (mx / m, my / m) if abs(m) > 1e-12 else cen
    return max(m, 0.0), centroid, _polygon_perimeter(ext), _max_distance(ext)


def contact_metrics_exact(fa: "FaceLoops", fb: "FaceLoops"):
    """True contact patch between two touching planar faces.

    Returns (area, perimeter, span, centroid_world, n_holes, exterior_2d,
             holes_2d, frame, flags) in the caller's length unit, or None.
    ``frame`` is (origin, x_axis, y_axis) of face A -- the plane a bolt pattern
    is laid out in and a weld path runs around.
    """
    ext_b, holes_b = _project_loops_into(fb, fa)
    ext, holes, flags = clip_loops(fa.exterior_2d, fa.holes_2d, ext_b, holes_b)
    if len(ext) < 3:
        return None
    area, cen2, perim, span = patch_props(ext, holes)
    if area <= 0.0:
        return None
    world = _local_to_world(fa.origin, fa.x_axis, fa.y_axis, cen2[0], cen2[1])
    frame = (fa.origin, fa.x_axis, fa.y_axis)
    return area, perim, span, world, len(holes), ext, holes, frame, flags


def obb_area_of(ext_2d) -> float:
    """Area of the axis-aligned box of a ring in its own frame -- what the V16
    OBB path would have reported for this face."""
    if len(ext_2d) < 2:
        return 0.0
    xs = [p[0] for p in ext_2d]
    ys = [p[1] for p in ext_2d]
    return (max(xs) - min(xs)) * (max(ys) - min(ys))


def obb_ring_world(oa: "FaceOBB", ob: "FaceOBB") -> list:
    """The V16 box-clip ring, in WORLD coordinates.

    This is exactly the polygon JointLocatorV16 measures its contact area from:
    face A's bounding rectangle clipped by face B's, both in A's plane. Returned
    in 3D so a viewer can draw it on top of the true contact patch -- the visible
    gap between the two IS the error the OBB method introduces.
    """
    rect_a = _ensure_ccw(_to_2d(oa.corners, oa, project=False))
    rect_b = _ensure_ccw(_to_2d(ob.corners, oa, project=True))
    clipped = _clip_polygon(rect_a, rect_b)
    if len(clipped) < 3:
        return []
    return [_local_to_world(oa.origin, oa.x_axis, oa.y_axis, u, v)
            for (u, v) in clipped]
