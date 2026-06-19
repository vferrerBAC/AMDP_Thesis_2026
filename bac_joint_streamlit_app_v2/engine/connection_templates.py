from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Dict, List

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


def default_connection_schedule(joint_ids: List[str], template_name: str = "Screw: lipped C to frame") -> pd.DataFrame:
    if not joint_ids:
        joint_ids = ["J001", "J002", "J003"]
    t = TEMPLATES.get(template_name, TEMPLATES["Screw: lipped C to frame"])
    rows = []
    for jid in joint_ids:
        row = asdict(t)
        row.update({"joint_id": str(jid), "notes": "Review template defaults before final use."})
        rows.append(row)
    return pd.DataFrame(rows)


def apply_template_to_rows(df: pd.DataFrame, selected_joint_ids: List[str], template_name: str) -> pd.DataFrame:
    out = df.copy()
    t = TEMPLATES[template_name]
    values = asdict(t)
    for jid in selected_joint_ids:
        mask = out["joint_id"].astype(str) == str(jid)
        if mask.any():
            for key, value in values.items():
                out.loc[mask, key] = value
        else:
            row = values.copy()
            row["joint_id"] = str(jid)
            row["notes"] = "Added from template."
            out = pd.concat([out, pd.DataFrame([row])], ignore_index=True)
    return out


def connection_schedule_template() -> pd.DataFrame:
    return default_connection_schedule(["J001", "J002", "J003"])
