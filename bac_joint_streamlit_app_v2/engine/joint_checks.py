from __future__ import annotations

import math
from typing import Any, Dict, List

import pandas as pd

FORCE_COLUMNS = ["Fx_lbf", "Fy_lbf", "Fz_lbf", "Mx_lbf_in", "My_lbf_in", "Mz_lbf_in"]
REQUIRED_CONNECTION_COLUMNS = [
    "joint_id", "template_name", "connection_type", "fastener_type", "n_fasteners",
    "diameter_in", "sheet_t_in", "edge_dist_in", "spacing_in", "weld_size_in",
    "weld_length_in", "custom_capacity_lbf", "material_grade", "notes"
]


def _num(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default


def ensure_force_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in ["joint_id", "combo_id"]:
        if col not in out.columns:
            out[col] = ""
    for col in FORCE_COLUMNS:
        if col not in out.columns:
            out[col] = 0.0
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out[["joint_id", "combo_id"] + FORCE_COLUMNS]


def ensure_connection_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in REQUIRED_CONNECTION_COLUMNS:
        if col not in out.columns:
            out[col] = "" if col not in ["n_fasteners", "diameter_in", "sheet_t_in", "edge_dist_in", "spacing_in", "weld_size_in", "weld_length_in", "custom_capacity_lbf"] else 0.0
    numeric_cols = ["n_fasteners", "diameter_in", "sheet_t_in", "edge_dist_in", "spacing_in", "weld_size_in", "weld_length_in", "custom_capacity_lbf"]
    for col in numeric_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out[REQUIRED_CONNECTION_COLUMNS]


def default_joint_forces(joint_ids: List[str], combo_ids: List[str]) -> pd.DataFrame:
    if not joint_ids:
        joint_ids = ["J001", "J002", "J003"]
    if not combo_ids:
        combo_ids = ["LRFD-01", "LRFD-02"]
    rows = []
    for idx, jid in enumerate(joint_ids[:12], start=1):
        for cidx, combo in enumerate(combo_ids[:5], start=1):
            rows.append({
                "joint_id": str(jid),
                "combo_id": str(combo),
                "Fx_lbf": 70.0 * idx * cidx,
                "Fy_lbf": -35.0 * idx,
                "Fz_lbf": 18.0 * cidx,
                "Mx_lbf_in": 8.0 * idx,
                "My_lbf_in": 10.0 * cidx,
                "Mz_lbf_in": 4.0 * idx * cidx,
            })
    return ensure_force_columns(pd.DataFrame(rows))


def force_demand_lbf(row: Dict[str, Any], lever_arm_in: float = 1.0) -> float:
    shear = math.sqrt(_num(row.get("Fx_lbf")) ** 2 + _num(row.get("Fy_lbf")) ** 2 + _num(row.get("Fz_lbf")) ** 2)
    moment = math.sqrt(_num(row.get("Mx_lbf_in")) ** 2 + _num(row.get("My_lbf_in")) ** 2 + _num(row.get("Mz_lbf_in")) ** 2)
    return shear + moment / max(lever_arm_in, 1.0)


def missing_data_reason(conn: Dict[str, Any]) -> str:
    ctype = str(conn.get("connection_type", "")).lower()
    if ctype == "ignore":
        return ""
    if ctype == "custom_capacity":
        if _num(conn.get("custom_capacity_lbf")) <= 0:
            return "Custom capacity is missing."
    elif ctype in ["screwed_sheet_joint", "bolted_bracket_joint"]:
        if _num(conn.get("n_fasteners")) <= 0:
            return "Number of fasteners is missing."
        if _num(conn.get("diameter_in")) <= 0:
            return "Fastener diameter is missing."
        if _num(conn.get("sheet_t_in")) <= 0:
            return "Sheet/bracket thickness is missing."
        if _num(conn.get("edge_dist_in")) <= 0:
            return "Edge distance is missing."
    elif ctype == "welded_joint":
        if _num(conn.get("weld_size_in")) <= 0:
            return "Weld size is missing."
        if _num(conn.get("weld_length_in")) <= 0:
            return "Weld length is missing."
    else:
        return "Connection type is missing or unsupported."
    return ""


def estimate_screening_capacity_lbf(conn: Dict[str, Any]) -> tuple[float, str]:
    """Demo screening capacity only. Replace with approved AISC/AISI/AWS/BAC equations."""
    ctype = str(conn.get("connection_type", "")).lower()
    n = _num(conn.get("n_fasteners"))
    d = _num(conn.get("diameter_in"))
    t = _num(conn.get("sheet_t_in"))
    edge = _num(conn.get("edge_dist_in"))
    spacing = _num(conn.get("spacing_in"))
    weld_size = _num(conn.get("weld_size_in"))
    weld_len = _num(conn.get("weld_length_in"))
    custom = _num(conn.get("custom_capacity_lbf"))

    if ctype == "ignore":
        return float("inf"), "Ignored by user."
    if ctype == "custom_capacity":
        return custom, "Engineer-supplied/tested capacity placeholder."
    if ctype == "screwed_sheet_joint":
        # Interface prototype only. This is not a code equation.
        capacity = n * max(d, 0.0) * max(t, 0.0) * 18500.0 * min(max(edge / max(2.0 * d, 0.001), 0.50), 1.20)
        return capacity, "Demo screw-group screening placeholder. Replace with AISI/BAC check."
    if ctype == "bolted_bracket_joint":
        gross_area = math.pi * (max(d, 0.0) ** 2) / 4.0
        edge_factor = min(max(edge / max(2.0 * d, 0.001), 0.50), 1.20)
        capacity = n * gross_area * 22000.0 * edge_factor
        return capacity, "Demo bolt-group screening placeholder. Replace with AISC/BAC check."
    if ctype == "welded_joint":
        capacity = max(weld_size, 0.0) * max(weld_len, 0.0) * 12500.0
        return capacity, "Demo weld screening placeholder. Replace with AWS/AISC/BAC check."
    return 0.0, "Unsupported connection type."


def suggestion_for(conn: Dict[str, Any], demand: float, capacity: float, status: str) -> str:
    ctype = str(conn.get("connection_type", "")).lower()
    if status == "Ignored":
        return "No action. User marked this joint as non-structural/ignored."
    if status == "Incomplete":
        return "Complete the missing connection inputs, then re-run the check."
    if status == "OK":
        return "No immediate action. Keep the joint in the design record."
    if status == "Warning":
        prefix = "Close to limit. Consider adding margin: "
    else:
        prefix = "Revise joint: "
    if ctype == "screwed_sheet_joint":
        if capacity > 0:
            needed = max(1, math.ceil(_num(conn.get("n_fasteners"), 1) * demand / capacity))
            return prefix + f"add screws, increase screw diameter, increase sheet thickness, or improve edge distance. Approx. fastener count target: {needed}."
        return prefix + "add screws, increase screw diameter, increase sheet thickness, or provide custom tested capacity."
    if ctype == "bolted_bracket_joint":
        return prefix + "increase bolt count/diameter, increase bracket thickness, improve edge distance, or add a reinforcing plate."
    if ctype == "welded_joint":
        return prefix + "increase weld length/size, use a different weld layout, or add a bracket/gusset."
    if ctype == "custom_capacity":
        return prefix + "increase approved capacity or use a stronger tested/engineer-approved detail."
    return prefix + "review the joint type and provide a supported connection template."


def status_from_utilization(utilization: float, missing_reason: str, ctype: str) -> str:
    if ctype == "ignore":
        return "Ignored"
    if missing_reason:
        return "Incomplete"
    if math.isinf(utilization):
        return "Incomplete"
    if utilization <= 0.80:
        return "OK"
    if utilization <= 1.00:
        return "Warning"
    return "Not OK"


def evaluate_joint_demands(joint_forces: pd.DataFrame, connection_schedule: pd.DataFrame) -> pd.DataFrame:
    forces = ensure_force_columns(joint_forces)
    conns = ensure_connection_columns(connection_schedule)
    merged = forces.merge(conns, on="joint_id", how="left", suffixes=("", "_conn"))
    rows = []
    for _, r in merged.iterrows():
        conn = r.to_dict()
        ctype = str(conn.get("connection_type", "")).lower()
        spacing = max(_num(conn.get("spacing_in")), 1.0)
        n = max(_num(conn.get("n_fasteners")), 1.0)
        lever = max(spacing * math.sqrt(n), 1.0)
        demand = force_demand_lbf(conn, lever)
        capacity, basis = estimate_screening_capacity_lbf(conn)
        missing = missing_data_reason(conn)
        utilization = demand / capacity if capacity and not math.isinf(capacity) else (0.0 if math.isinf(capacity) else float("inf"))
        status = status_from_utilization(utilization, missing, ctype)
        issue = missing if status == "Incomplete" else ("Demand exceeds screening capacity." if status == "Not OK" else ("Utilization is close to the limit." if status == "Warning" else "No issue detected in screening."))
        rows.append({
            "joint_id": conn.get("joint_id", ""),
            "combo_id": conn.get("combo_id", ""),
            "connection_template": conn.get("template_name", ""),
            "connection_type": conn.get("connection_type", ""),
            "force_demand_lbf": round(demand, 2),
            "screening_capacity_lbf": round(capacity, 2) if not math.isinf(capacity) else float("inf"),
            "percent_used": round(utilization * 100.0, 1) if not math.isinf(utilization) else None,
            "status": status,
            "plain_language_issue": issue,
            "suggested_fix": suggestion_for(conn, demand, capacity, status),
            "capacity_basis": basis,
            **{col: conn.get(col, 0.0) for col in FORCE_COLUMNS},
        })
    result = pd.DataFrame(rows)
    if result.empty:
        return result
    status_rank = {"Not OK": 4, "Warning": 3, "Incomplete": 2, "OK": 1, "Ignored": 0}
    result["status_rank"] = result["status"].map(status_rank).fillna(0).astype(int)
    result = result.sort_values(["status_rank", "percent_used"], ascending=[False, False], na_position="last").drop(columns=["status_rank"])
    return result


def critical_joint_summary(results: pd.DataFrame) -> pd.DataFrame:
    if results.empty:
        return pd.DataFrame(columns=["joint_id", "worst_status", "max_percent_used", "governing_combo", "suggested_fix"])
    status_rank = {"Not OK": 4, "Warning": 3, "Incomplete": 2, "OK": 1, "Ignored": 0}
    rows = []
    for jid, g in results.groupby("joint_id"):
        gg = g.copy()
        gg["rank"] = gg["status"].map(status_rank).fillna(0).astype(int)
        gg["pct_sort"] = pd.to_numeric(gg["percent_used"], errors="coerce").fillna(-1)
        worst = gg.sort_values(["rank", "pct_sort"], ascending=[False, False]).iloc[0]
        rows.append({
            "joint_id": jid,
            "worst_status": worst["status"],
            "max_percent_used": worst["percent_used"],
            "governing_combo": worst["combo_id"],
            "plain_language_issue": worst["plain_language_issue"],
            "suggested_fix": worst["suggested_fix"],
        })
    return pd.DataFrame(rows).sort_values(["worst_status", "max_percent_used"], ascending=[False, False], na_position="last")


def joint_balance_table(joint_forces: pd.DataFrame) -> pd.DataFrame:
    forces = ensure_force_columns(joint_forces)
    if forces.empty:
        return pd.DataFrame(columns=["joint_id", "combo_id", "sum_Fx_lbf", "sum_Fy_lbf", "sum_Fz_lbf", "sum_Mx_lbf_in", "sum_My_lbf_in", "sum_Mz_lbf_in", "reaction_resultant_lbf"])
    grouped = forces.groupby(["joint_id", "combo_id"], as_index=False)[FORCE_COLUMNS].sum()
    grouped = grouped.rename(columns={
        "Fx_lbf": "sum_Fx_lbf", "Fy_lbf": "sum_Fy_lbf", "Fz_lbf": "sum_Fz_lbf",
        "Mx_lbf_in": "sum_Mx_lbf_in", "My_lbf_in": "sum_My_lbf_in", "Mz_lbf_in": "sum_Mz_lbf_in",
    })
    grouped["reaction_resultant_lbf"] = (grouped["sum_Fx_lbf"] ** 2 + grouped["sum_Fy_lbf"] ** 2 + grouped["sum_Fz_lbf"] ** 2) ** 0.5
    return grouped
