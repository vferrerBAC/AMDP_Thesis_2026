from __future__ import annotations

from typing import Dict, Tuple

import pandas as pd

from .joint_checks import ensure_connection_columns, ensure_force_columns, missing_data_reason


def _row(item: str, status: str, message: str, action: str = "") -> Dict[str, str]:
    return {"check": item, "status": status, "message": message, "recommended_action": action}


def validation_checklist(joints: pd.DataFrame, members: pd.DataFrame, combos: pd.DataFrame, connections: pd.DataFrame, forces: pd.DataFrame, units: str) -> pd.DataFrame:
    rows = []
    if joints is not None and not joints.empty:
        rows.append(_row("Geometry", "Passed", f"{len(joints)} joints detected.", ""))
    else:
        rows.append(_row("Geometry", "Blocked", "No joints are loaded.", "Upload Block 1 JSON or load sample geometry."))

    if members is not None and not members.empty:
        rows.append(_row("Member connectivity", "Passed", f"{len(members)} members linked to joints.", ""))
    else:
        rows.append(_row("Member connectivity", "Warning", "No member connectivity was detected.", "Check Block 1 JSON if joint/member preview is needed."))

    if str(units) == "inch-lbf":
        rows.append(_row("Units", "Passed", "inch-lbf is supported by the current prototype.", ""))
    else:
        rows.append(_row("Units", "Warning", f"Selected units are {units}; current demo check columns use lbf and inches.", "Add full unit conversion before design use."))

    if combos is not None and not combos.empty and "combo_id" in combos.columns:
        rows.append(_row("Load combinations", "Passed", f"{len(combos)} load combinations available.", ""))
    else:
        rows.append(_row("Load combinations", "Blocked", "No load combinations are available.", "Create or upload load combinations."))

    if connections is not None and not connections.empty:
        conns = ensure_connection_columns(connections)
        # Coverage is checked per JOINT (every joint must have at least one
        # connection row) but completeness is checked per CONNECTION, because a
        # joint with three patches needs three complete rows, not one.
        joint_ids = set(joints["joint_id"].astype(str).tolist()) if joints is not None and not joints.empty else set()
        scheduled_joints = set(conns["joint_id"].astype(str).tolist())
        missing_conn = sorted(joint_ids - scheduled_joints)
        dupes = conns["connection_id"].astype(str)
        repeated = sorted(dupes[dupes.duplicated()].unique().tolist())
        if missing_conn:
            rows.append(_row("Connection schedule", "Blocked", f"{len(missing_conn)} joints have no connection rows: {', '.join(missing_conn[:5])}.", "Extract connections (Step 3), then complete the schedule."))
        elif repeated:
            rows.append(_row("Connection schedule", "Blocked", f"Duplicate connection IDs: {', '.join(repeated[:5])}. Each row must be one contact patch.", "Give every row a unique Connection ID."))
        else:
            missing_data = []
            for _, r in conns.iterrows():
                reason = missing_data_reason(r.to_dict())
                if reason:
                    missing_data.append((r.get("connection_id", ""), reason))
            if missing_data:
                examples = "; ".join([f"{cid}: {reason}" for cid, reason in missing_data[:3]])
                rows.append(_row("Connection schedule", "Blocked", f"{len(missing_data)} rows have missing connection data. {examples}", "Open the bolted-connections step and complete missing values."))
            else:
                n_multi = int((conns.groupby("joint_id")["connection_id"].count() > 1).sum())
                extra = f" {n_multi} joint(s) have more than one contact patch." if n_multi else ""
                rows.append(_row("Connection schedule", "Passed", f"{len(conns)} connection rows are complete enough for screening.{extra}", ""))
    else:
        rows.append(_row("Connection schedule", "Blocked", "No connection schedule is available.", "Create one from a template."))

    if forces is not None and not forces.empty:
        f = ensure_force_columns(forces)
        bad_joints = set(f["joint_id"].astype(str)) - (set(joints["joint_id"].astype(str)) if joints is not None and not joints.empty else set())
        combo_ids = set(combos["combo_id"].astype(str)) if combos is not None and not combos.empty and "combo_id" in combos.columns else set()
        bad_combos = set(f["combo_id"].astype(str)) - combo_ids if combo_ids else set()
        if bad_joints:
            rows.append(_row("Joint forces", "Blocked", f"Forces reference unknown joints: {', '.join(sorted(list(bad_joints))[:5])}.", "Fix joint_id values in the force table."))
        elif bad_combos:
            rows.append(_row("Joint forces", "Warning", f"Some force rows reference combos not in the combo table: {', '.join(sorted(list(bad_combos))[:5])}.", "Add those combinations or rename combo_id values."))
        else:
            rows.append(_row("Joint forces", "Passed", f"{len(f)} joint force rows available.", ""))
    else:
        rows.append(_row("Joint forces", "Blocked", "No joint force demand table is available.", "Upload force results or generate demo loads."))

    df = pd.DataFrame(rows)
    df["ready_blocker"] = df["status"].eq("Blocked")
    return df


def is_ready_to_run(checklist: pd.DataFrame) -> Tuple[bool, str]:
    blockers = checklist[checklist["status"].eq("Blocked")]
    if blockers.empty:
        return True, "Ready to run screening analysis."
    return False, f"Not ready: fix {len(blockers)} blocked item(s)."
