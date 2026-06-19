from __future__ import annotations

import json
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

JOINT_LIST_NAMES = {
    "joints", "nodes", "joint_table", "node_table", "points", "vertices", "joint_coordinates"
}
MEMBER_LIST_NAMES = {
    "members", "frames", "elements", "beams", "frame_table", "member_table", "connectivity"
}
ID_KEYS = ["joint_id", "node_id", "id", "name", "label", "guid"]
MEMBER_ID_KEYS = ["member_id", "frame_id", "element_id", "id", "name", "label", "guid"]
X_KEYS = ["x", "X", "x_in", "X_in", "global_x", "coord_x"]
Y_KEYS = ["y", "Y", "y_in", "Y_in", "global_y", "coord_y"]
Z_KEYS = ["z", "Z", "z_in", "Z_in", "global_z", "coord_z"]
POINT_KEYS = ["point", "coord", "coords", "coordinate", "coordinates", "centroid", "origin", "location"]
START_KEYS = ["start_joint", "i_joint", "node_i", "joint_i", "start_node", "from", "start"]
END_KEYS = ["end_joint", "j_joint", "node_j", "joint_j", "end_node", "to", "end"]
SECTION_KEYS = ["section", "section_type", "profile", "profile_type", "shape", "cross_section"]
MATERIAL_KEYS = ["material", "material_name", "mat", "coating"]


def load_json_from_upload(uploaded_file: Any) -> Dict[str, Any]:
    raw = uploaded_file.getvalue().decode("utf-8")
    return json.loads(raw)


def _norm_key(key: str) -> str:
    return str(key).lower().replace(" ", "_").replace("-", "_")


def _get_first(d: Dict[str, Any], keys: Iterable[str], default: Any = None) -> Any:
    lookup = {_norm_key(k): k for k in d.keys()}
    for key in keys:
        nk = _norm_key(key)
        if nk in lookup:
            return d[lookup[nk]]
    return default


def _as_float(value: Any, default: Optional[float] = None) -> Optional[float]:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except Exception:
        return default


def _extract_point(d: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    if isinstance(d, dict):
        x = _get_first(d, X_KEYS)
        y = _get_first(d, Y_KEYS)
        z = _get_first(d, Z_KEYS)
        if x is not None and y is not None:
            return _as_float(x), _as_float(y), _as_float(z, 0.0)
        for key in POINT_KEYS:
            p = _get_first(d, [key])
            if p is not None:
                return _extract_point(p)
    if isinstance(d, (list, tuple)) and len(d) >= 2:
        return _as_float(d[0]), _as_float(d[1]), _as_float(d[2] if len(d) > 2 else 0.0)
    return None, None, None


def _find_named_lists(obj: Any, candidate_names: set[str]) -> List[List[Dict[str, Any]]]:
    found: List[List[Dict[str, Any]]] = []
    if isinstance(obj, dict):
        for key, value in obj.items():
            if _norm_key(key) in candidate_names and isinstance(value, list):
                if all(isinstance(item, dict) for item in value):
                    found.append(value)
            found.extend(_find_named_lists(value, candidate_names))
    elif isinstance(obj, list):
        for item in obj:
            found.extend(_find_named_lists(item, candidate_names))
    return found


def _find_candidate_joint_dicts(obj: Any) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if isinstance(obj, dict):
        x, y, z = _extract_point(obj)
        has_id = any(_get_first(obj, [key]) is not None for key in ID_KEYS)
        if has_id and x is not None and y is not None:
            rows.append(obj)
        for value in obj.values():
            rows.extend(_find_candidate_joint_dicts(value))
    elif isinstance(obj, list):
        for item in obj:
            rows.extend(_find_candidate_joint_dicts(item))
    return rows


def _find_candidate_member_dicts(obj: Any) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if isinstance(obj, dict):
        has_id = any(_get_first(obj, [key]) is not None for key in MEMBER_ID_KEYS)
        has_start = any(_get_first(obj, [key]) is not None for key in START_KEYS)
        has_end = any(_get_first(obj, [key]) is not None for key in END_KEYS)
        if has_id and has_start and has_end:
            rows.append(obj)
        for value in obj.values():
            rows.extend(_find_candidate_member_dicts(value))
    elif isinstance(obj, list):
        for item in obj:
            rows.extend(_find_candidate_member_dicts(item))
    return rows


def _normalize_joint_rows(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    out = []
    seen = set()
    for idx, row in enumerate(rows, start=1):
        jid = _get_first(row, ID_KEYS, f"J{idx:04d}")
        x, y, z = _extract_point(row)
        if x is None or y is None:
            continue
        key = str(jid)
        if key in seen:
            continue
        seen.add(key)
        out.append({"joint_id": key, "x": x, "y": y, "z": z if z is not None else 0.0})
    return pd.DataFrame(out, columns=["joint_id", "x", "y", "z"])


def _joint_ref(value: Any) -> str:
    if isinstance(value, dict):
        ref = _get_first(value, ID_KEYS + ["joint", "node", "ref"])
        if ref is not None:
            return str(ref)
        x, y, z = _extract_point(value)
        if x is not None and y is not None:
            return f"POINT({x:.4f},{y:.4f},{(z or 0.0):.4f})"
    return str(value)


def _normalize_member_rows(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    out = []
    seen = set()
    for idx, row in enumerate(rows, start=1):
        mid = _get_first(row, MEMBER_ID_KEYS, f"M{idx:04d}")
        start = _get_first(row, START_KEYS)
        end = _get_first(row, END_KEYS)
        if start is None or end is None:
            continue
        key = str(mid)
        if key in seen:
            continue
        seen.add(key)
        out.append({
            "member_id": key,
            "start_joint": _joint_ref(start),
            "end_joint": _joint_ref(end),
            "section_type": str(_get_first(row, SECTION_KEYS, "unknown")),
            "material": str(_get_first(row, MATERIAL_KEYS, "unknown")),
        })
    return pd.DataFrame(out, columns=["member_id", "start_joint", "end_joint", "section_type", "material"])


def extract_joints_and_members(data: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    joint_lists = _find_named_lists(data, JOINT_LIST_NAMES)
    member_lists = _find_named_lists(data, MEMBER_LIST_NAMES)
    joint_rows: List[Dict[str, Any]] = []
    member_rows: List[Dict[str, Any]] = []
    for lst in joint_lists:
        joint_rows.extend(lst)
    for lst in member_lists:
        member_rows.extend(lst)
    if not joint_rows:
        joint_rows = _find_candidate_joint_dicts(data)
    if not member_rows:
        member_rows = _find_candidate_member_dicts(data)
    return _normalize_joint_rows(joint_rows), _normalize_member_rows(member_rows)


def geometry_metrics(joints: pd.DataFrame, members: pd.DataFrame) -> Dict[str, Any]:
    metrics = {"n_joints": int(len(joints)), "n_members": int(len(members)), "x_span": 0.0, "y_span": 0.0, "z_span": 0.0}
    if not joints.empty:
        for axis in ["x", "y", "z"]:
            metrics[f"{axis}_span"] = float(joints[axis].max() - joints[axis].min())
    return metrics


def connected_members_summary(joints: pd.DataFrame, members: pd.DataFrame) -> pd.DataFrame:
    if joints.empty:
        return pd.DataFrame(columns=["joint_id", "connected_member_count", "connected_members", "sections", "materials"])
    rows = []
    for jid in joints["joint_id"].astype(str).tolist():
        if members.empty:
            conn = pd.DataFrame()
        else:
            conn = members[(members["start_joint"].astype(str) == jid) | (members["end_joint"].astype(str) == jid)]
        rows.append({
            "joint_id": jid,
            "connected_member_count": int(len(conn)),
            "connected_members": ", ".join(conn.get("member_id", pd.Series(dtype=str)).astype(str).tolist()),
            "sections": ", ".join(sorted(set(conn.get("section_type", pd.Series(dtype=str)).astype(str).tolist()))),
            "materials": ", ".join(sorted(set(conn.get("material", pd.Series(dtype=str)).astype(str).tolist()))),
        })
    return pd.DataFrame(rows)
