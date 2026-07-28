from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Dict, List, Tuple

import pandas as pd


@dataclass(frozen=True)
class ConnectionTemplate:
    template_name: str
    connection_type: str
    user_friendly_name: str
    when_to_use: str
    fastener_type: str = ""
    n_fasteners: int = 0
    diameter_in: float = 0.0
    sheet_t_in: float = 0.0
    edge_dist_in: float = 0.0
    spacing_in: float = 0.0
    weld_size_in: float = 0.0
    weld_length_in: float = 0.0
    custom_capacity_lbf: float = 0.0
    material_grade: str = "Unknown"
    recommended_data_needed: str = ""


TEMPLATES: Dict[str, ConnectionTemplate] = {
    "Screw: sheet lap joint": ConnectionTemplate(
        template_name="Screw: sheet lap joint",
        connection_type="screwed_sheet_joint",
        user_friendly_name="Screwed sheet-metal lap joint",
        when_to_use="Two thin sheet-metal parts overlap and are joined with self-drilling or sheet-metal screws.",
        fastener_type="self_drilling_screw",
        n_fasteners=4,
        diameter_in=0.190,
        sheet_t_in=0.050,
        edge_dist_in=0.50,
        spacing_in=1.00,
        material_grade="Galvanized steel",
        recommended_data_needed="number of screws, screw size, sheet thickness, edge distance, spacing",
    ),
    "Screw: lipped C to frame": ConnectionTemplate(
        template_name="Screw: lipped C to frame",
        connection_type="screwed_sheet_joint",
        user_friendly_name="Lipped C / channel section screwed to frame",
        when_to_use="A thin formed channel, C-lipped, Z, or hat section is attached to a supporting frame using screws.",
        fastener_type="self_drilling_screw",
        n_fasteners=6,
        diameter_in=0.190,
        sheet_t_in=0.060,
        edge_dist_in=0.50,
        spacing_in=1.25,
        material_grade="HDG / GLV steel",
        recommended_data_needed="number of screws, section thickness, screw line spacing, edge distance",
    ),
    "Bolt: bracket to frame": ConnectionTemplate(
        template_name="Bolt: bracket to frame",
        connection_type="bolted_bracket_joint",
        user_friendly_name="Bolted bracket or plate joint",
        when_to_use="A bracket, plate, or clip angle is attached using bolts rather than sheet-metal screws.",
        fastener_type="bolt",
        n_fasteners=2,
        diameter_in=0.250,
        sheet_t_in=0.075,
        edge_dist_in=0.75,
        spacing_in=1.50,
        material_grade="Steel bracket",
        recommended_data_needed="bolt count, bolt diameter, bracket thickness, edge distance, spacing",
    ),
    "Weld: tube to plate": ConnectionTemplate(
        template_name="Weld: tube to plate",
        connection_type="welded_joint",
        user_friendly_name="Welded tube or frame joint",
        when_to_use="A tube, frame, or heavy member is attached to another member or plate with welds.",
        fastener_type="fillet_weld",
        n_fasteners=0,
        diameter_in=0.0,
        sheet_t_in=0.075,
        edge_dist_in=0.0,
        spacing_in=0.0,
        weld_size_in=0.125,
        weld_length_in=3.0,
        material_grade="Weldable steel",
        recommended_data_needed="weld size, effective weld length, base metal thickness, weld layout",
    ),
    "Custom tested capacity": ConnectionTemplate(
        template_name="Custom tested capacity",
        connection_type="custom_capacity",
        user_friendly_name="Custom / tested joint capacity",
        when_to_use="Use this when BAC has a test result, validated supplier data, or an engineer-approved capacity for this joint.",
        fastener_type="engineer_capacity",
        n_fasteners=0,
        diameter_in=0.0,
        sheet_t_in=0.0,
        edge_dist_in=0.0,
        spacing_in=0.0,
        custom_capacity_lbf=1000.0,
        material_grade="Engineer supplied",
        recommended_data_needed="approved capacity and notes explaining the source",
    ),
    "Ignore / non-structural": ConnectionTemplate(
        template_name="Ignore / non-structural",
        connection_type="ignore",
        user_friendly_name="Ignore this joint",
        when_to_use="Use for cosmetic, reference, duplicate, or non-load-bearing joints that should not be checked.",
        fastener_type="not_applicable",
        recommended_data_needed="reason for ignoring the joint",
    ),
}


def template_names() -> List[str]:
    return list(TEMPLATES.keys())


def templates_table() -> pd.DataFrame:
    return pd.DataFrame([asdict(t) for t in TEMPLATES.values()])


def is_active_connection(c) -> bool:
    """True when a connection should be checked, costed, and shown downstream.

    The single definition of "active", so the schedule, the demand table, the 3D
    view, and the counts can never disagree about what the model contains. Two
    things deactivate a connection, both reversible and both leaving the record
    in place:

    ``manually_removed``  the engineer unchecked it in Step 3.
    ``pruned_noncontact`` the members cannot reach each other (see
                          ``block1.prune_noncontact_connections``) — unless the
                          engineer overrode that with ``prune_override``, or the
                          face pass has since confirmed a real contact, which
                          outranks the centerline heuristic outright.
    """
    if not isinstance(c, dict):
        return False
    if c.get("manually_removed"):
        return False
    if c.get("detection_method") == "face_contact":
        return True
    if c.get("pruned_noncontact") and not c.get("prune_override"):
        return False
    return True


def connection_keys(block1_data: Dict) -> List[Tuple[str, str]]:
    """[(connection_id, joint_id)] for every ACTIVE connection in a Block 1 model.

    The schedule is built from CONNECTIONS, not joints: one contact patch is one
    checkable thing, and a joint can own several of them with different
    thicknesses, bolt counts, and materials.

    Inactive connections (unchecked by the engineer, or flagged unreachable) are
    skipped: they are not real contacts, so they must not appear as blank rows in
    the bolted schedule. See ``is_active_connection``.
    """
    out: List[Tuple[str, str]] = []
    for c in (block1_data or {}).get("connections") or []:
        if not is_active_connection(c):
            continue
        cid = str(c.get("connection_id") or "").strip()
        if cid:
            out.append((cid, str(c.get("joint_id") or "").strip()))
    return out


def default_connection_schedule(connection_keys_: List[Tuple[str, str]],
                                template_name: str = "Screw: lipped C to frame") -> pd.DataFrame:
    if not connection_keys_:
        connection_keys_ = [("C001", "J001"), ("C002", "J002"), ("C003", "J003")]
    t = TEMPLATES.get(template_name, TEMPLATES["Screw: lipped C to frame"])
    rows = []
    for cid, jid in connection_keys_:
        row = asdict(t)
        row.update({"connection_id": str(cid), "joint_id": str(jid),
                    "notes": "Review template defaults before final use."})
        rows.append(row)
    return pd.DataFrame(rows)


def apply_template_to_rows(df: pd.DataFrame, selected_connection_ids: List[str],
                           template_name: str) -> pd.DataFrame:
    out = df.copy()
    t = TEMPLATES[template_name]
    values = asdict(t)
    for cid in selected_connection_ids:
        mask = out["connection_id"].astype(str) == str(cid)
        if mask.any():
            for key, value in values.items():
                out.loc[mask, key] = value
        else:
            row = values.copy()
            row["connection_id"] = str(cid)
            row["notes"] = "Added from template."
            out = pd.concat([out, pd.DataFrame([row])], ignore_index=True)
    return out


def connections_in_joint(df: pd.DataFrame, joint_id: str) -> List[str]:
    """Every connection_id belonging to one joint — backs the UI's
    'apply to all patches in this joint' bulk action."""
    if df is None or df.empty or "joint_id" not in df.columns:
        return []
    mask = df["joint_id"].astype(str) == str(joint_id)
    return df.loc[mask, "connection_id"].astype(str).tolist()


def connection_schedule_template() -> pd.DataFrame:
    return default_connection_schedule([("C001", "J001"), ("C002", "J001"),
                                        ("C003", "J002")])
