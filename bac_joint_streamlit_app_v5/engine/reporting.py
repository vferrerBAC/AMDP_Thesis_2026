from __future__ import annotations

import io
import json
from datetime import datetime
from typing import Dict

import pandas as pd


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = str(name)[:31]
            if isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=safe_name, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def to_json_bytes(config: dict) -> bytes:
    return json.dumps(config, indent=2).encode("utf-8")


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def to_pdf_bytes(config: dict, summary: pd.DataFrame, results: pd.DataFrame, checklist: pd.DataFrame) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    project = config.get("project", {})
    story.append(Paragraph("BAC Joint Check Assistant — Screening Report", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Project: {project.get('project_name', 'Unnamed project')}", styles["Heading2"]))
    story.append(Paragraph(f"Generated: {datetime.now().isoformat(timespec='seconds')}", styles["Normal"]))
    story.append(Paragraph(f"Mode: {config.get('app_mode', 'Simple/Advanced not recorded')}", styles["Normal"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Important Limitation", styles["Heading2"]))
    story.append(Paragraph("This prototype uses demo screening capacities for workflow development. It is not a final AISC/AISI/AWS code-check engine and should not be used for design release without approved equations and professional review.", styles["Normal"]))
    story.append(Spacer(1, 12))

    def add_table(title: str, df: pd.DataFrame, max_rows: int = 12):
        story.append(Paragraph(title, styles["Heading2"]))
        if df is None or df.empty:
            story.append(Paragraph("No data available.", styles["Normal"]))
            story.append(Spacer(1, 8))
            return
        dd = df.head(max_rows).copy().fillna("")
        data = [list(dd.columns)] + dd.astype(str).values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B5CAD")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(table)
        story.append(Spacer(1, 12))

    add_table("Validation Checklist", checklist[["check", "status", "message"]] if checklist is not None and not checklist.empty else pd.DataFrame())
    add_table("Joint Summary", summary if summary is not None else pd.DataFrame())
    result_cols = [c for c in ["joint_id", "combo_id", "status", "percent_used", "plain_language_issue", "suggested_fix"] if results is not None and c in results.columns]
    add_table("Governing Joint Results", results[result_cols] if results is not None and result_cols else pd.DataFrame())
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
