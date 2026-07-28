"""
ASCE 7-22 LRFD loads + Block-3 FE bridge for the Streamlit app.
================================================================
Turns the raw Block-1 geometry dict into the two tables the existing engine
already consumes:

  * joint_forces  -> columns  joint_id, combo_id, Fx_lbf..Mz_lbf_in
                     Joint-level display/balance only. The CHECKS no longer run
                     off this table: connection_demand.connection_demands takes
                     the solver's per-member global end actions and distributes
                     them across contact patches, because capacity lives on the
                     connection, not on the joint.
  * combo_table   -> columns  combo_id, combo_name, expression, plain_language_notes
                     (replaces load_combinations.default_load_combinations so
                      joint_forces.combo_id always matches the combo table)

Heavy lifting (PyNite frame solve) lives in engine.block3_solver; this file only
adapts its output to the app's schema and derives a member-connectivity table
for display from the real Block-1 schema (joints[].member_names), which the
app's generic parser does not pick up.

Units throughout: lb, in, lb*in (imperial).
"""
from __future__ import annotations

import math
from typing import Any, Dict, List, Optional

import pandas as pd

from engine.block3_solver import analyze_joints, DEFAULT_CONFIG

FORCE_COLUMNS = ["Fx_lbf", "Fy_lbf", "Fz_lbf", "Mx_lbf_in", "My_lbf_in", "Mz_lbf_in"]

_COMBO_META = {
    "1.4D": ("Dead only", "1.4D", "Gravity-only strength case."),
    "W1+":  ("Wind (1.2D)",              "1.2D + 1.0W",            "Wind, full gravity, +direction."),
    "W1-":  ("Wind (1.2D), rev.",        "1.2D - 1.0W",            "Wind, full gravity, reversed."),
    "W2+":  ("Uplift wind (0.9D)",       "0.9D + 1.0W",            "Wind, reduced stabilizing dead, +direction."),
    "W2-":  ("Uplift wind (0.9D), rev.", "0.9D - 1.0W",            "Wind, reduced stabilizing dead, reversed."),
    "E1+":  ("Seismic (1.2D)",           "(1.2+0.2Sds)D + rho*QE", "Seismic, full gravity + vertical Ev, +direction."),
    "E1-":  ("Seismic (1.2D), rev.",     "(1.2+0.2Sds)D - rho*QE", "Seismic, full gravity + vertical Ev, reversed."),
    "E2+":  ("Uplift seismic (0.9D)",    "(0.9-0.2Sds)D + rho*QE", "Seismic, reduced dead - Ev, +direction."),
    "E2-":  ("Uplift seismic (0.9D), rev.", "(0.9-0.2Sds)D - rho*QE", "Seismic, reduced dead - Ev, reversed."),
}


def members_dataframe_from_block1(block1_data: Dict[str, Any]) -> pd.DataFrame:
    """Build the app-schema member table from joints[].member_names.

    The app's generic parser only links members carrying start_joint/end_joint
    ids. The real Block-1 schema lists member_names on each joint, so we invert
    that to recover start_joint / end_joint for display.
    """
    cols = ["member_id", "start_joint", "end_joint", "section_type", "material"]
    joints = block1_data.get("joints", [])
    members = {m.get("occurrence_name"): m for m in block1_data.get("members", [])}

    incidence: Dict[str, List[str]] = {}
    for j in joints:
        for mname in j.get("member_names", []):
            incidence.setdefault(mname, []).append(j.get("joint_id"))
    jloc = {j.get("joint_id"): j.get("location") for j in joints}

    def _order(mname: str, jids: List[str]) -> List[str]:
        m = members.get(mname)
        if not m or len(jids) < 2:
            return jids
        s, e = m.get("start_point"), m.get("end_point")
        ax = [e[i] - s[i] for i in range(3)]
        L2 = sum(c * c for c in ax) or 1.0
        return sorted(jids, key=lambda jid: sum((jloc[jid][i] - s[i]) * ax[i]
                                                 for i in range(3)) / L2)

    rows = []
    for mname, jids in incidence.items():
        m = members.get(mname, {})
        cs = m.get("cross_section", {}) if isinstance(m, dict) else {}
        ordered = _order(mname, jids)
        rows.append({
            "member_id": mname,
            "start_joint": ordered[0] if ordered else "",
            "end_joint": ordered[-1] if len(ordered) > 1 else "",
            "section_type": str(cs.get("section_type", "unknown")),
            "material": str(cs.get("family", m.get("material", "unknown"))),
        })
    return pd.DataFrame(rows, columns=cols)


def run_lrfd_joint_analysis(
    block1_data: Dict[str, Any],
    *,
    S_DS: float = 1.0,
    rho: float = 1.3,
    wind_loads: Optional[List[Dict[str, Any]]] = None,
    seismic_h_loads: Optional[List[Dict[str, Any]]] = None,
    support_nodes: Optional[List[str]] = None,
    support_fixity: str = "fixed",
    source_length_unit: str = "in",
    source_mass_unit: str = "lb",
    include_1_4D: bool = True,
) -> Dict[str, Any]:
    """Wrapper over engine.block3_solver.analyze_joints with app defaults.

    wind_loads / seismic_h_loads: lists of {"node","FX","FY","FZ"} point loads (lb).
    Returns the solver dict (status 'ok' or 'error').
    """
    config = {
        **DEFAULT_CONFIG,
        "S_DS": float(S_DS), "rho": float(rho),
        "source_length_unit": source_length_unit,
        "source_mass_unit": source_mass_unit,
        "support_fixity": support_fixity,
        "include_1_4D": include_1_4D,
    }
    if support_nodes:
        config["support_nodes"] = support_nodes
    loads = {"wind": wind_loads or [], "seismic_h": seismic_h_loads or []}
    return analyze_joints(block1_data, config=config, loads=loads)


def _resultant(row: Dict[str, float]) -> float:
    shear = math.sqrt(row["N"] ** 2 + row["Vy"] ** 2 + row["Vz"] ** 2)
    moment = math.sqrt(row["T"] ** 2 + row["My"] ** 2 + row["Mz"] ** 2)
    return shear + moment


def joint_forces_table(result: Dict[str, Any]) -> pd.DataFrame:
    """Reduce per-member end forces to one governing demand per (joint, combo).

    DISPLAY ONLY. This collapses a joint to its single worst member end, which is
    the right summary for a reactions-style table and the wrong input for a
    capacity check — it cannot say which contact patch carries what. The checks
    read connection_demand.connection_demands instead.

    Member-local components map to the app's columns (the summary metric uses
    only invariant shear/moment magnitudes, so the local->global naming does not
    affect it):
        N -> Fx_lbf, Vy -> Fy_lbf, Vz -> Fz_lbf,
        T -> Mx_lbf_in, My -> My_lbf_in, Mz -> Mz_lbf_in
    """
    out: List[Dict[str, Any]] = []
    if result.get("status") != "ok":
        return pd.DataFrame(columns=["joint_id", "combo_id"] + FORCE_COLUMNS)
    for combo, joints in result["joint_member_forces"].items():
        for jid, rows in joints.items():
            gov = max(rows, key=_resultant) if rows else \
                {"N": 0, "Vy": 0, "Vz": 0, "T": 0, "My": 0, "Mz": 0}
            out.append({
                "joint_id": jid, "combo_id": combo,
                "Fx_lbf": gov["N"], "Fy_lbf": gov["Vy"], "Fz_lbf": gov["Vz"],
                "Mx_lbf_in": gov["T"], "My_lbf_in": gov["My"], "Mz_lbf_in": gov["Mz"],
            })
    return pd.DataFrame(out, columns=["joint_id", "combo_id"] + FORCE_COLUMNS)


def combos_table(result: Dict[str, Any]) -> pd.DataFrame:
    """Combo table matching the combo_ids present in the solved forces."""
    combos = result.get("meta", {}).get("combos", []) if result.get("status") == "ok" else []
    rows = []
    for cid in combos:
        name, expr, note = _COMBO_META.get(cid, (cid, cid, ""))
        rows.append({"combo_id": cid, "combo_name": name,
                     "expression": expr, "plain_language_notes": note})
    return pd.DataFrame(rows, columns=["combo_id", "combo_name", "expression", "plain_language_notes"])


# --------------------------------------------------------------------------- #
#  Parametric loads:  psf / g  ->  nodal point loads
# --------------------------------------------------------------------------- #
# Equipment qualification presets (wind pressure, SDS).  At a rigid mount and
# x/h = 1.0 the horizontal seismic coefficient is taken equal to SDS (the full
# top-of-structure rigid case); swap in the ASCE 7-22 Ch.13 Fp(ap,Rp,Ip,z/h)
# equation here if a code-rigorous component force is required.
PARAMETRIC_PRESETS = {
    "30 psf / 0.3g  (SDS 0.3, rigid, x/h=1.0)": {"wind_psf": 30.0, "sds": 0.3},
    "70 psf / 0.7g  (SDS 0.7, rigid, x/h=1.0)": {"wind_psf": 70.0, "sds": 0.7},
    "Custom": None,
}
_LEN_TO_FT = {"in": 1.0 / 12.0, "ft": 1.0, "cm": 1.0 / 30.48, "mm": 1.0 / 304.8}
_COMP = {"X (+)": "FX", "Z (+)": "FZ", "Y (+)": "FY"}


def nodal_dead_weights(block1_data: Dict[str, Any],
                       source_mass_unit: str = "lb") -> Dict[str, float]:
    """Lbf of self-weight lumped to each joint - mirrors the solver's lumping
    (half of each member's weight to its two extreme end joints)."""
    from engine.block3_solver import _MASS_TO_LBF
    kmass = _MASS_TO_LBF[source_mass_unit]
    joints = block1_data.get("joints", [])
    members = {m.get("occurrence_name"): m for m in block1_data.get("members", [])}
    jloc = {j.get("joint_id"): j.get("location") for j in joints}
    incidence: Dict[str, List[str]] = {}
    for j in joints:
        for mn in j.get("member_names", []):
            incidence.setdefault(mn, []).append(j.get("joint_id"))
    w = {j.get("joint_id"): 0.0 for j in joints}
    for mn, jids in incidence.items():
        m = members.get(mn)
        if not m or len(jids) < 2:
            continue
        wt = float(m.get("dry_mass", 0.0)) * kmass
        s, e = m.get("start_point"), m.get("end_point")
        ax = [e[i] - s[i] for i in range(3)]
        L2 = sum(c * c for c in ax) or 1.0
        ends = sorted(jids, key=lambda jid: sum((jloc[jid][i] - s[i]) * ax[i]
                                                 for i in range(3)) / L2)
        w[ends[0]] += wt / 2.0
        w[ends[-1]] += wt / 2.0
    return w


# Coordinate resolution.  block1_raw is stored as the RAW uploaded dict (before
# joint_io normalization), so it may arrive in either schema:
#   * native Block-1 output : joint["location"], member["start_point"/"end_point"]
#   * authoring / upload     : joint x/y/z,       member start_joint/end_joint (id refs)
# The bounding box below needs coordinates from whichever schema is present, so
# these mirror joint_io's alias tables locally (keeping engine.loads dependency-free).
_X_KEYS = ("x", "X", "x_in", "X_in", "global_x", "coord_x")
_Y_KEYS = ("y", "Y", "y_in", "Y_in", "global_y", "coord_y")
_Z_KEYS = ("z", "Z", "z_in", "Z_in", "global_z", "coord_z")
_POINT_KEYS = ("location", "point", "coord", "coords", "coordinate",
               "coordinates", "centroid", "origin")
_ID_KEYS = ("joint_id", "node_id", "id", "name", "label", "guid")
_START_KEYS = ("start_point", "start_joint", "i_joint", "node_i", "joint_i",
               "start_node", "from", "start")
_END_KEYS = ("end_point", "end_joint", "j_joint", "node_j", "joint_j",
             "end_node", "to", "end")


def _first(d: Dict[str, Any], keys) -> Any:
    for k in keys:
        if k in d and d[k] is not None:
            return d[k]
    return None


def _as_xyz(value: Any) -> Optional[tuple]:
    """Coerce a list/tuple/dict coordinate into a 3-float tuple, else None."""
    if isinstance(value, (list, tuple)) and len(value) >= 3:
        try:
            return (float(value[0]), float(value[1]), float(value[2]))
        except (TypeError, ValueError):
            return None
    if isinstance(value, dict):
        x, y, z = _first(value, _X_KEYS), _first(value, _Y_KEYS), _first(value, _Z_KEYS)
        if None not in (x, y, z):
            try:
                return (float(x), float(y), float(z))
            except (TypeError, ValueError):
                return None
    return None


def _joint_xyz(j: Dict[str, Any]) -> Optional[tuple]:
    """Joint coordinate from either schema: location/point object, or x/y/z."""
    pt = _as_xyz(_first(j, _POINT_KEYS))
    if pt is not None:
        return pt
    x, y, z = _first(j, _X_KEYS), _first(j, _Y_KEYS), _first(j, _Z_KEYS)
    if None not in (x, y, z):
        try:
            return (float(x), float(y), float(z))
        except (TypeError, ValueError):
            return None
    return None


def _joint_locmap(block1_data: Dict[str, Any]) -> Dict[str, tuple]:
    out: Dict[str, tuple] = {}
    for j in block1_data.get("joints", []):
        jid = _first(j, _ID_KEYS)
        xyz = _joint_xyz(j)
        if jid is not None and xyz is not None:
            out[str(jid)] = xyz
    return out


def _member_endpoints(m: Dict[str, Any], locmap: Dict[str, tuple]) -> tuple:
    """Return (start_xyz, end_xyz), resolving id refs through locmap. Either may be None."""
    s_raw, e_raw = _first(m, _START_KEYS), _first(m, _END_KEYS)
    s = _as_xyz(s_raw)
    if s is None and s_raw is not None:
        s = locmap.get(str(s_raw))
    e = _as_xyz(e_raw)
    if e is None and e_raw is not None:
        e = locmap.get(str(e_raw))
    return s, e


def estimate_projected_area_ft2(block1_data: Dict[str, Any], wind_dir: str = "X (+)",
                                source_length_unit: str = "in") -> float:
    """Bounding-box silhouette area (ft^2) on the plane normal to the wind.
    Y is vertical.

    Uses only joint locations and member centerline endpoints - a member's own
    flange/cross-section footprint is deliberately excluded from the envelope
    (per convention: flange overhang doesn't count toward projected wind area;
    the frame's own member spacing - e.g. two parallel chords straddling the
    silhouette - already defines the true width).

    The one place cross-section size still matters is the minimal fallback:
    if an axis collapses to zero spread on its own (e.g. a single member with
    no other joint/member offsetting that axis), that axis is padded by half
    of the lone member's own max(width, depth) so the silhouette isn't
    degenerate. This is never applied to axes that already have real spread.
    """
    k = _LEN_TO_FT.get(source_length_unit, 1.0 / 12.0)
    locmap = _joint_locmap(block1_data)
    pts_ft: List[tuple] = []
    for j in block1_data.get("joints", []):
        loc = _joint_xyz(j)
        if loc:
            pts_ft.append(tuple(c * k for c in loc))

    members = block1_data.get("members", [])
    axis_fallback_half_ft = [0.0, 0.0, 0.0]   # per global axis, only used if that axis collapses
    for m in members:
        s, e = _member_endpoints(m, locmap)
        if not s or not e:
            continue
        pts_ft.append(tuple(c * k for c in s))
        pts_ft.append(tuple(c * k for c in e))

        cs = m.get("cross_section") or {}
        depth, width = cs.get("depth") or 0.0, cs.get("width") or 0.0
        dim = max(width, depth)
        if dim <= 0:
            continue
        kcs = _LEN_TO_FT.get(cs.get("length_unit", "in"), 1.0 / 12.0)
        half_ft = (dim * kcs) / 2.0
        axis = [e[i] - s[i] for i in range(3)]
        amag = math.sqrt(sum(c * c for c in axis)) or 1.0
        dominant = max(range(3), key=lambda i: abs(axis[i]) / amag)
        for i in range(3):
            if i != dominant:
                axis_fallback_half_ft[i] = max(axis_fallback_half_ft[i], half_ft)

    if not pts_ft:
        return 0.0

    tol_ft = 1e-6
    extents = []
    for i in range(3):
        span = max(p[i] for p in pts_ft) - min(p[i] for p in pts_ft)
        if span <= tol_ft:
            span = 2.0 * axis_fallback_half_ft[i]
        extents.append(span)
    dx, dy, dz = extents

    if wind_dir.startswith("X"):
        return abs(dy * dz)
    if wind_dir.startswith("Z"):
        return abs(dy * dx)
    return abs(dx * dz)


# Plane basis for each wind direction: (e1, e2) are the two GLOBAL axes spanning
# the plane normal to the wind, so a 3D point projects to (p.e1, p.e2). Y is
# vertical, so for the (fixed) X wind the silhouette lands on the Y-Z plane.
_WIND_PLANE_BASIS = {
    "X": ((0.0, 1.0, 0.0), (0.0, 0.0, 1.0)),   # project onto Y, Z
    "Y": ((1.0, 0.0, 0.0), (0.0, 0.0, 1.0)),   # project onto X, Z
    "Z": ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0)),   # project onto X, Y
}


def _flagp(nr, ident, field, issue):
    nr.append({"scope": "member", "identifier": ident, "field": field, "issue": issue})


def member_projected_area_ft2(block1_data: Dict[str, Any], wind_dir: str = "X (+)",
                              source_length_unit: str = "in") -> Dict[str, Any]:
    """True wind projected (silhouette) area A_s, in ft^2, on the plane normal
    to the wind.

    Each member is projected as its oriented section box (centerline x
    ``depth``.``depth_dir`` x ``width``.``width_dir``); the convex hull of the
    projected corners is that member's exact silhouette, and all silhouettes are
    UNIONed so overlapping shadows -- members that cross on the windward face, or
    sit one behind another along the wind -- are counted once. That union is the
    quantity ``wind_load = q * A_s`` needs, and is what the bounding-box
    ``estimate_projected_area_ft2`` billboard over-states for an open frame.

    Returns a dict (not a bare float) so coverage is auditable:
        area_ft2      union silhouette area  (the value to use)
        sum_area_ft2  naive per-member sum   (>= area_ft2; the gap is shielding)
        n_used        members that contributed a silhouette
        n_total       members seen
        needs_review  per-member flags for anything skipped or approximated
        method        "member_silhouette"

    NOTE (flag, not applied): this is the gross geometric solid area. Any ASCE 7
    open-frame shielding/solidity coefficient between windward and leeward faces
    is a separate empirical factor and is deliberately NOT applied here.
    """
    try:
        from shapely.geometry import MultiPoint
        from shapely.ops import unary_union
    except Exception as exc:                       # shapely is a project dep; be explicit if absent
        raise RuntimeError(
            "member_projected_area_ft2 needs shapely (already in requirements.txt): "
            f"{exc}") from exc

    axis = (wind_dir or "X")[0].upper()
    e1, e2 = _WIND_PLANE_BASIS.get(axis, _WIND_PLANE_BASIS["X"])
    k_pt = _LEN_TO_FT.get(source_length_unit, 1.0 / 12.0)
    locmap = _joint_locmap(block1_data)

    def _proj_ft(pt3, k):
        p = (pt3[0] * k, pt3[1] * k, pt3[2] * k)
        return (p[0] * e1[0] + p[1] * e1[1] + p[2] * e1[2],
                p[0] * e2[0] + p[1] * e2[1] + p[2] * e2[2])

    members = block1_data.get("members", [])
    needs_review: List[Dict[str, Any]] = []
    polys = []
    sum_area_ft2 = 0.0
    n_used = 0

    for m in members:
        ident = m.get("occurrence_path") or m.get("occurrence_name") or m.get("part_number") or "?"
        s, e = _member_endpoints(m, locmap)
        if not s or not e:
            _flagp(needs_review, ident, "endpoints", "no start/end point; excluded from wind area")
            continue

        cs = m.get("cross_section") or {}
        D = cs.get("depth") or 0.0
        W = cs.get("width") or 0.0
        if D <= 0 or W <= 0:
            _flagp(needs_review, ident, "cross_section", "no depth/width; excluded from wind area (A_s undercounts)")
            continue
        k_cs = _LEN_TO_FT.get(cs.get("length_unit", "in"), 1.0 / 12.0)
        ud, uw = cs.get("depth_dir"), cs.get("width_dir")

        if ud and uw:
            # Exact: 8 corners of the oriented section box, projected then hulled.
            hd = [D / 2.0 * c for c in ud]
            hw = [W / 2.0 * c for c in uw]
            pts2d = []
            for P in (s, e):
                for sd in (-1.0, 1.0):
                    for sw in (-1.0, 1.0):
                        c3 = (P[0] + sd * hd[0] * (k_cs / k_pt) + sw * hw[0] * (k_cs / k_pt),
                              P[1] + sd * hd[1] * (k_cs / k_pt) + sw * hw[1] * (k_cs / k_pt),
                              P[2] + sd * hd[2] * (k_cs / k_pt) + sw * hw[2] * (k_cs / k_pt))
                        pts2d.append(_proj_ft(c3, k_pt))
        else:
            # Orientation unknown: conservative axis strip of face width max(D,W)
            # so the member's windward area is not dropped (dropping = unsafe).
            _flagp(needs_review, ident, "section_dirs",
                   "no depth_dir/width_dir; used conservative max(depth,width) strip")
            a = _proj_ft(s, k_pt)
            b = _proj_ft(e, k_pt)
            dxu, dyu = b[0] - a[0], b[1] - a[1]
            L = math.hypot(dxu, dyu) or 1e-9
            nx, ny = -dyu / L, dxu / L                 # in-plane perpendicular
            hw_ft = (max(D, W) * k_cs) / 2.0
            pts2d = [(a[0] + s2 * nx * hw_ft, a[1] + s2 * ny * hw_ft) for s2 in (-1, 1)]
            pts2d += [(b[0] + s2 * nx * hw_ft, b[1] + s2 * ny * hw_ft) for s2 in (-1, 1)]

        hull = MultiPoint(pts2d).convex_hull
        if getattr(hull, "geom_type", "") == "Polygon" and hull.area > 0:
            polys.append(hull)
            sum_area_ft2 += hull.area
            n_used += 1

    area_ft2 = float(unary_union(polys).area) if polys else 0.0
    return {
        "area_ft2": area_ft2,
        "sum_area_ft2": float(sum_area_ft2),
        "n_used": n_used,
        "n_total": len(members),
        "needs_review": needs_review,
        "method": "member_silhouette",
    }


def wind_projected_area(block1_data: Dict[str, Any], wind_dir: str = "X (+)",
                        source_length_unit: str = "in") -> Dict[str, Any]:
    """Primary entry: true member-silhouette A_s when sections are available,
    else the bounding-box billboard (clearly labelled). Always returns a dict."""
    try:
        res = member_projected_area_ft2(block1_data, wind_dir, source_length_unit)
        if res["n_used"] > 0:
            return res
        reason = "no member carried a usable cross-section"
    except Exception as exc:
        reason = f"silhouette method unavailable ({exc})"

    bbox = estimate_projected_area_ft2(block1_data, wind_dir, source_length_unit)
    return {
        "area_ft2": bbox,
        "sum_area_ft2": bbox,
        "n_used": 0,
        "n_total": len(block1_data.get("members", [])),
        "needs_review": [{"scope": "model", "identifier": "-", "field": "cross_section",
                          "issue": f"Fell back to bounding-box silhouette: {reason}. "
                                   "This treats the frame as a solid billboard and "
                                   "over-states A for an open frame."}],
        "method": "bounding_box_fallback",
    }


# ---------------------------------------------------------------------------
# BAC tributary-area wind method (1-way system)
# ---------------------------------------------------------------------------
# Per the BAC procedure, wind is collected by the members that span the face in
# the chosen 1-way direction (the vertical columns / mullions in BAC's examples):
#
#     tributary_area   = tributary_width * tributary_height
#     tributary_height = the member's own length (its longest side)
#     tributary_width  = half the distance to the adjacent collector on each side
#     wind_force        = wind_pressure(psf) * tributary_area
#
# Because each band runs halfway to its neighbours, the bands TILE the face: for
# a fully-populated frame the tributary areas sum back to the gross windward
# face, and the load is apportioned per collector by spacing. This is the model
# for a clad BAC unit (the casing catches the wind); it is NOT the open-frame
# silhouette in member_projected_area_ft2 above.
_AXIS_IDX = {"X": 0, "Y": 1, "Z": 2}


def tributary_wind_area(block1_data: Dict[str, Any], wind_dir: str = "auto",
                        source_length_unit: str = "in", *,
                        vertical_axis: str = "Y",
                        adjacency_gap_in: Optional[float] = None) -> Dict[str, Any]:
    """Per-collector tributary areas (ft^2) for the BAC 1-way wind method.

    Collectors are the members whose dominant axis is ``vertical_axis`` (the
    columns/mullions). Wind must blow PERPENDICULAR to the frame plane to strike
    the broad face; the tributary width is then measured along the in-plane
    horizontal axis the collectors spread along.

    ``wind_dir``:
        "auto"  -> derive the windward face from geometry: the collectors spread
                   along the wider horizontal axis (that becomes the spacing
                   axis), so wind is pointed along the narrower/thin horizontal
                   axis, perpendicular to the frame plane. This is orientation-
                   independent and is the safe default.
        "X (+)"/"Z (+)"/... -> force the wind axis; spacing is then the other
                   horizontal axis. (Wind IN the frame plane hits it edge-on and
                   yields ~0 area -- that mismatch is what "auto" avoids.)

    Edge collectors' tributary width runs out to the frame edge on the open side
    when ``edge_to_frame`` is True (so the bands tile the full face).

    Returns:
        total_area_ft2  sum over collectors (the value for the A field)
        per_member      list of {member, height_ft, width_ft, area_ft2, pos_in}
        span_axis       "X"|"Y"|"Z" the spacing was measured on
        wind_axis       "X"|"Y"|"Z" the wind was pointed along
        auto            True if the axes were derived from geometry
        n_collectors    how many collecting members were found
        needs_review    flags (no collectors, single collector, non-collectors, ...)
        method          "tributary"
    """
    vert = vertical_axis.strip().upper()
    vi = _AXIS_IDX[vert]
    horiz = [a for a in ("X", "Y", "Z") if a != vert]     # the two non-vertical axes
    review: List[Dict[str, Any]] = []

    k = _LEN_TO_FT.get(source_length_unit, 1.0 / 12.0)
    locmap = _joint_locmap(block1_data)

    # Pre-pass: collect vertical members (with full midpoints) + all coordinates,
    # so we can auto-pick the face and get frame extents on any axis.
    all_coords = {0: [], 1: [], 2: []}
    for j in block1_data.get("joints", []):
        xyz = _joint_xyz(j)
        if xyz:
            for i in range(3):
                all_coords[i].append(xyz[i])
    members = block1_data.get("members", [])
    raw_collectors = []          # (mid3, length_in, ident)
    n_noncollector = 0
    for m in members:
        s, e = _member_endpoints(m, locmap)
        if not s or not e:
            continue
        for i in range(3):
            all_coords[i].append(s[i]); all_coords[i].append(e[i])
        axis = [e[i] - s[i] for i in range(3)]
        amag = math.sqrt(sum(c * c for c in axis)) or 1.0
        dominant = max(range(3), key=lambda i: abs(axis[i]))
        ident = m.get("occurrence_path") or m.get("occurrence_name") or m.get("part_number") or "?"
        if dominant == vi:
            mid = [(s[i] + e[i]) / 2.0 for i in range(3)]
            vlo, vhi = (s[vi], e[vi]) if s[vi] <= e[vi] else (e[vi], s[vi])
            raw_collectors.append((mid, amag, ident, m.get("cross_section") or {}, vlo, vhi))
        else:
            n_noncollector += 1

    if n_noncollector:
        review.append({"scope": "model", "identifier": "-", "field": "collectors",
                       "issue": f"{n_noncollector} member(s) do not run along the vertical "
                                f"axis '{vert}' and are treated as non-collectors in this "
                                "1-way system (their face area is carried by the collectors)."})

    if not raw_collectors:
        review.append({"scope": "model", "identifier": "-", "field": "collectors",
                       "issue": f"no members run along the vertical axis '{vert}'; "
                                "no tributary bands could be formed."})
        return {"total_area_ft2": 0.0, "per_member": [], "span_axis": None,
                "wind_axis": None, "auto": False, "n_collectors": 0,
                "needs_review": review, "method": "tributary"}

    # Resolve wind & spacing axes.
    wtxt = (wind_dir or "").strip().upper()
    if wtxt.startswith("AUTO") or wtxt == "":
        def _spread(letter):
            idx = _AXIS_IDX[letter]
            ps = [c[0][idx] for c in raw_collectors]
            return (max(ps) - min(ps)) if ps else 0.0
        h0, h1 = horiz
        span_ax = h0 if _spread(h0) >= _spread(h1) else h1   # collectors spread wider here
        wind_ax = h1 if span_ax == h0 else h0                # thin axis -> perpendicular to face
        auto = True
        review.append({"scope": "model", "identifier": "-", "field": "wind_dir",
                       "issue": f"auto: frame spreads on {span_ax}, so wind was pointed along "
                                f"{wind_ax} (perpendicular to the broad face). "
                                f"spread {span_ax}={_spread(span_ax):.1f}, "
                                f"{wind_ax}={_spread(wind_ax):.1f} (source units)."})
    else:
        wind_ax = wtxt[0]
        if wind_ax == vert or wind_ax not in _AXIS_IDX:
            review.append({"scope": "model", "identifier": "-", "field": "wind_dir",
                           "issue": f"wind '{wind_ax}' is not a valid horizontal axis with "
                                    f"vertical '{vert}'."})
            return {"total_area_ft2": 0.0, "per_member": [], "span_axis": None,
                    "wind_axis": None, "auto": False, "n_collectors": len(raw_collectors),
                    "needs_review": review, "method": "tributary"}
        span_ax = [a for a in horiz if a != wind_ax][0]
        auto = False

    si = _AXIS_IDX[span_ax]

    def _own_width_ft(cs):
        """Member's own face width along the in-face horizontal axis (span_ax),
        in feet. Returns (width_ft, issue)."""
        D = cs.get("depth") or 0.0
        W = cs.get("width") or 0.0
        kcs = _LEN_TO_FT.get(cs.get("length_unit", source_length_unit), k)
        if D <= 0 and W <= 0:
            return 0.0, "no cross-section depth/width; width unknown (area under-counts)"
        ud, uw = cs.get("depth_dir"), cs.get("width_dir")
        if ud and uw:
            # projected extent of the oriented section box onto the span axis
            w_src = D * abs(ud[si]) + W * abs(uw[si])
            return w_src * kcs, None
        return max(D, W) * kcs, "no section orientation; used max(depth,width) as face width"

    # One entry per collector: lateral position (span axis), vertical extent
    # [vlo, vhi] on the vertical axis, own face width (all in ft).
    entries = []
    for (mid, length_in, ident, cs, vlo, vhi) in raw_collectors:
        wft, issue = _own_width_ft(cs)
        if issue:
            review.append({"scope": "member", "identifier": ident,
                           "field": "tributary_width", "issue": issue})
        entries.append({"ident": ident, "pos": mid[si] * k,
                        "vlo": min(vlo, vhi) * k, "vhi": max(vlo, vhi) * k, "own_w": wft})

    # --- Project onto the wind-normal (span x vertical, e.g. Z-Y) plane -------
    # The tributary area is a SILHOUETTE on that plane. Collectors that share a
    # lateral position but differ along the WIND axis -- the front and back of a
    # 3-D weldment, or a single line split into stacked segments -- project onto
    # the SAME strip of the face. The wind sees that strip once, so summing those
    # members double-counts area. Merge collectors into one line per lateral
    # strip; the strip's silhouette height is the UNION of its members' vertical
    # extents (never their sum), and its width is that strip's own face width.
    pos_tol = 1e-6
    if entries:
        wmax = max((e["own_w"] for e in entries), default=0.0)
        pos_tol = max(pos_tol, 0.10 * wmax)      # coincident lines within 1/10 face-width
    entries.sort(key=lambda e: e["pos"])
    lines: List[Dict[str, Any]] = []
    for e in entries:
        if lines and abs(e["pos"] - lines[-1]["pos_sum"] / lines[-1]["n"]) <= pos_tol:
            g = lines[-1]
            g["pos_sum"] += e["pos"]; g["n"] += 1
            g["intervals"].append((e["vlo"], e["vhi"]))
            g["own_w"] = max(g["own_w"], e["own_w"])
            g["idents"].append(e["ident"])
        else:
            lines.append({"pos_sum": e["pos"], "n": 1,
                          "intervals": [(e["vlo"], e["vhi"])],
                          "own_w": e["own_w"], "idents": [e["ident"]]})

    def _union_len(intervals):
        """Length covered by the union of 1-D [lo, hi] intervals."""
        segs = sorted(intervals)
        clo, chi = segs[0]
        tot = 0.0
        for lo, hi in segs[1:]:
            if lo <= chi:                        # overlapping / touching -> extend
                chi = max(chi, hi)
            else:
                tot += chi - clo; clo, chi = lo, hi
        return tot + (chi - clo)

    rows = []
    for g in lines:
        rows.append({"pos": g["pos_sum"] / g["n"],
                     "height": _union_len(g["intervals"]),
                     "own_w": g["own_w"],
                     "n_members": g["n"],
                     "ident": (g["idents"][0] if g["n"] == 1
                               else f"{g['idents'][0]} (+{g['n'] - 1} in-line)")})
    rows.sort(key=lambda r: r["pos"])

    merged = sum(g["n"] - 1 for g in lines)
    if merged:
        review.append({"scope": "model", "identifier": "-", "field": "projection",
                       "issue": f"{merged} collector(s) share a lateral line with another "
                                "(stacked segments, or front/back across the wind axis) and were "
                                f"merged into the {len(rows)} face strip(s) so the wind silhouette "
                                "is not double-counted on the projected plane."})

    # Adjacency: a neighbour on a given side is "adjacent" (share half the distance
    # to it) only when the centre-to-centre spacing is within adjacency_gap_in.
    # Otherwise that side stops at the member's own face edge. adjacency_gap_in is
    # None -> no sharing at all: every member gets its own width (the open-frame /
    # standalone case, e.g. the N0GNGVM5 columns). The total is always floored at
    # the member's own width, since a member at least catches its own footprint.
    gap_ft = None if adjacency_gap_in is None else float(adjacency_gap_in) * k
    if adjacency_gap_in is not None:
        review.append({"scope": "model", "identifier": "-", "field": "adjacency_gap",
                       "issue": f"members spaced within {adjacency_gap_in:g} in share the gap "
                                "(half-distance to the nearest neighbour); wider spacing -> "
                                "standalone (own width)."})

    per_member = []
    total_area_ft2 = 0.0
    ncnt = len(rows)
    for i, r in enumerate(rows):
        p, w = r["pos"], r["own_w"]
        # left boundary
        if i > 0 and gap_ft is not None and (p - rows[i - 1]["pos"]) <= gap_ft:
            left_bound = (p + rows[i - 1]["pos"]) / 2.0
        else:
            left_bound = p - w / 2.0
        # right boundary
        if i < ncnt - 1 and gap_ft is not None and (rows[i + 1]["pos"] - p) <= gap_ft:
            right_bound = (p + rows[i + 1]["pos"]) / 2.0
        else:
            right_bound = p + w / 2.0
        width_ft = max(right_bound - left_bound, w)     # never less than own footprint
        area_ft2 = r["height"] * width_ft
        total_area_ft2 += area_ft2
        per_member.append({"member": r["ident"], "height_ft": round(r["height"], 3),
                           "width_ft": round(width_ft, 3), "area_ft2": round(area_ft2, 3),
                           "own_width_ft": round(w, 3), "pos_in": round(p / k, 3),
                           "n_members": r["n_members"]})
    n = ncnt

    return {"total_area_ft2": float(total_area_ft2), "per_member": per_member,
            "span_axis": span_ax, "wind_axis": wind_ax, "auto": auto,
            "n_collectors": n, "n_strips": len(rows),
            "n_members": len(raw_collectors), "needs_review": review,
            "method": "tributary"}


def tributary_wind_loads(block1_data: Dict[str, Any], *, wind_psf: float,
                         wind_dir: str = "auto", source_length_unit: str = "in",
                         vertical_axis: str = "Y",
                         adjacency_gap_in: Optional[float] = None) -> tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Per-node wind loads from the tributary method: each collector's force
    ``q * tributary_area`` is split 50/50 to its two end joints. Returns
    ``(wind_loads, info)`` with wind_loads in the ``{"node", "FX"}`` schema the
    solver already consumes, so it drops in where parametric wind did."""
    trib = tributary_wind_area(block1_data, wind_dir, source_length_unit,
                               vertical_axis=vertical_axis, adjacency_gap_in=adjacency_gap_in)
    wcomp = "F" + (trib.get("wind_axis") or "X")      # resolved axis (handles "auto")
    locmap = _joint_locmap(block1_data)

    # member ident -> its two end joint ids (nearest joints to each endpoint)
    def _end_joints(m):
        ids = m.get("member_names") or []
        if len(ids) >= 2:
            return ids[0], ids[-1]
        # fall back to nearest joint by coordinate
        s, e = _member_endpoints(m, locmap)
        def nearest(pt):
            best, bd = None, None
            for jid, xyz in locmap.items():
                d = sum((xyz[i] - pt[i]) ** 2 for i in range(3))
                if bd is None or d < bd:
                    best, bd = jid, d
            return best
        return (nearest(s) if s else None), (nearest(e) if e else None)

    by_ident = {}
    for m in block1_data.get("members", []):
        ident = m.get("occurrence_path") or m.get("occurrence_name") or m.get("part_number") or "?"
        by_ident[ident] = m

    node_force: Dict[str, float] = {}
    total = 0.0
    for rec in trib["per_member"]:
        f = float(wind_psf) * rec["area_ft2"]
        total += f
        m = by_ident.get(rec["member"])
        if not m:
            continue
        j0, j1 = _end_joints(m)
        for jid in (j0, j1):
            if jid is not None:
                node_force[str(jid)] = node_force.get(str(jid), 0.0) + f / 2.0

    wind_loads = [{"node": jid, wcomp: val} for jid, val in node_force.items()]
    info = {"wind_total_lb": round(total, 1), "total_area_ft2": round(trib["total_area_ft2"], 2),
            "n_collectors": trib["n_collectors"], "n_loaded_nodes": len(wind_loads),
            "span_axis": trib["span_axis"], "wind_axis": trib.get("wind_axis"),
            "auto": trib.get("auto"), "needs_review": trib["needs_review"]}
    return wind_loads, info


def parametric_to_point_loads(
    block1_data: Dict[str, Any], *,
    wind_psf: float, ref_area_ft2: float, wind_dir: str,
    seismic_coeff_g: float, seismic_dir: str,
    source_mass_unit: str = "lb",
    operating_weight_lb: Optional[float] = None,
) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    """Convert (wind pressure, seismic coefficient) into nodal point loads.

    Wind:    total = wind_psf * ref_area_ft2,  distributed to joints by weight
             fraction (proxy for tributary area), applied along wind_dir.
    Seismic: per-joint horizontal QE = seismic_coeff_g * (joint self-weight)
             i.e. total = coeff * operating weight, mass-proportional (F = m*a),
             applied along seismic_dir.  Returned QE is the reference vector;
             the load combinations apply +/- and the rho factor.
    """
    w = nodal_dead_weights(block1_data, source_mass_unit)
    Wp_dead = sum(w.values())
    Wp = float(operating_weight_lb) if operating_weight_lb else (Wp_dead or 1.0)
    scale = (Wp / Wp_dead) if Wp_dead else 1.0       # scale weights to Wp

    wind_total = float(wind_psf) * float(ref_area_ft2)
    seismic_total = float(seismic_coeff_g) * Wp
    wcomp = _COMP.get(wind_dir, "FX")
    scomp = _COMP.get(seismic_dir, "FX")

    wind_loads, seismic_loads = [], []
    for node, wn in w.items():
        wn_s = wn * scale
        frac = wn_s / Wp if Wp else 0.0
        wind_loads.append({"node": node, wcomp: wind_total * frac})
        seismic_loads.append({"node": node, scomp: float(seismic_coeff_g) * wn_s})

    info = {"operating_weight_lb": round(Wp, 1), "dead_weight_lb": round(Wp_dead, 1),
            "wind_total_lb": round(wind_total, 1), "seismic_total_lb": round(seismic_total, 1),
            "n_loaded_nodes": len(w)}
    return wind_loads, seismic_loads, info
