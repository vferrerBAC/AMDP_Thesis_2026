"""
engine/cost_page.py — Streamlit UI for the Milestone 2 cost engine run.

Two entry points:
  render_cost_run(cost_parts, cost_joints)  -> engine-only: method picks + Excel
      run + results. Use this when the caller ALREADY built the inputs with
      build_cost_inputs (e.g. the Financial Analysis tab in app.py).
  render_cost_analysis(block1_data)         -> convenience: builds inputs then
      calls render_cost_run. For a standalone page.

COM notes: the Excel run is button-gated (Excel isn't launched on every rerun),
CoInitialize-wrapped (Streamlit's worker thread isn't COM-initialized), and the
result is cached in session_state. pythoncom is imported defensively so this
module still imports on non-Windows hosts (the run then errors clearly).
"""
from __future__ import annotations
import os
from pathlib import Path

import streamlit as st
import pandas as pd

from engine.cost_adapter import build_cost_inputs
from engine.cost import (
    run_cost_analysis, read_method_options,
    FORMING_METHOD_DEFAULT, CUTTING_METHOD_DEFAULT_TUBE, CUTTING_METHOD_DEFAULT_SHEET_METAL,
)

try:
    import pythoncom
    _PYWIN32 = True
except Exception:
    _PYWIN32 = False

TEMPLATE_NAME = "cost_calculator - Clean - 12FEB26.xlsx"

# Display labels for the fabrication-region codes. The workbook still receives the
# code (NA/EU/CN); only the dropdown label is expanded for the user.
REGION_LABELS = {"NA": "North America", "EU": "Europe", "CN": "China"}


def _engine_template_path() -> Path:
    return Path(__file__).parent / TEMPLATE_NAME


def _default_cut(part_class):
    return CUTTING_METHOD_DEFAULT_TUBE if part_class == "tube" else CUTTING_METHOD_DEFAULT_SHEET_METAL


def render_cost_run(cost_parts, cost_joints, template_path: str | None = None, output_dir: str | None = None):
    """Engine-only: takes inputs already built by build_cost_inputs."""
    if not cost_parts:
        return

    template_path = template_path or str(_engine_template_path())
    output_dir = output_dir or str(Path(template_path).parent.parent / "outputs" / "cost")

    st.markdown("### Run cost engine")
    if not os.path.exists(template_path):
        st.error(f"Cost template not found at {template_path}. "
                 f"Place {TEMPLATE_NAME} in the engine/ folder.")
        return

    # Method options straight from the template's named ranges (cached).
    if "cost_method_options" not in st.session_state:
        st.session_state["cost_method_options"] = read_method_options(template_path)
    opts = st.session_state["cost_method_options"]

    c1, c2 = st.columns(2)
    region = c1.selectbox("Fabrication region", opts["fabrication_region"],
                          index=opts["fabrication_region"].index("NA") if "NA" in opts["fabrication_region"] else 0,
                          format_func=lambda r: REGION_LABELS.get(r, r))
    length_uom = c2.selectbox("Length units", opts["length_uom"],
                              index=opts["length_uom"].index("in") if "in" in opts["length_uom"] else 0)

    # Per-part method picks; defaults pre-filled (Manual Press Brake / tube->Auto Tube Laser).
    df = pd.DataFrame([{
        "Part": p["part_identifier"], "Qty": p["part_quantity"], "Class": p["part_class"],
        "Cutting Method": p["cutting_method"] or _default_cut(p["part_class"]),
        "Forming Method": p["forming_method"] or FORMING_METHOD_DEFAULT,
    } for p in cost_parts])
    edited = st.data_editor(
        df, hide_index=True, width="stretch", disabled=["Part", "Qty", "Class"],
        column_config={
            "Cutting Method": st.column_config.SelectboxColumn(options=opts["cutting_method"]),
            "Forming Method": st.column_config.SelectboxColumn(options=opts["forming_method"]),
        },
        key="cost_method_editor",
    )

    if not _PYWIN32:
        st.info("Excel COM (pywin32) isn't available on this host — the cost run needs Windows + Excel. "
                "You can still review inputs and method picks here.")

    if st.button("Run cost analysis", type="primary", width="stretch", disabled=not _PYWIN32):
        picks = {r["Part"]: (r["Cutting Method"], r["Forming Method"]) for _, r in edited.iterrows()}
        for p in cost_parts:
            p["cutting_method"], p["forming_method"] = picks.get(p["part_identifier"], (None, None))
        ready_joints = [j for j in cost_joints if j.get("joint_length_inches") is not None]

        os.makedirs(output_dir, exist_ok=True)
        pythoncom.CoInitialize()
        try:
            with st.spinner("Recalculating in Excel… (large SAP tables, give it a moment)"):
                st.session_state["cost_result"] = run_cost_analysis(
                    cost_parts, ready_joints, template_path=template_path,
                    output_dir=output_dir, region=region, length_uom=length_uom)
        except Exception as e:
            st.error(f"Cost run failed: {e}")
        finally:
            pythoncom.CoUninitialize()

    result = st.session_state.get("cost_result")
    if result:
        st.markdown("#### Summary")
        st.dataframe(pd.DataFrame(result["summary"]), width="stretch", hide_index=True)
        if result.get("parts"):
            st.markdown("#### Cost per part")
            st.dataframe(
                pd.DataFrame(result["parts"]), width="stretch", hide_index=True,
                column_config={
                    "Part Quantity": st.column_config.NumberColumn("Qty", format="%d"),
                    "Unit Material Cost": st.column_config.NumberColumn(format="$%.2f"),
                    "Unit Fully Burdened Labor Cost": st.column_config.NumberColumn(format="$%.2f"),
                    "Unit Fully Burdened Total Cost": st.column_config.NumberColumn(format="$%.2f"),
                    "Extended Material Cost": st.column_config.NumberColumn(format="$%.2f"),
                    "Extended Fully Burdened Labor Cost": st.column_config.NumberColumn(format="$%.2f"),
                    "Extended Fully Burdened Total Cost": st.column_config.NumberColumn(format="$%.2f"),
                },
            )
        if result["needs_review"]:
            with st.expander(f"{len(result['needs_review'])} flags from the run"):
                st.dataframe(pd.DataFrame(result["needs_review"]), width="stretch", hide_index=True)
        wb = result["output_workbook"]
        with open(wb, "rb") as f:
            st.download_button(
                "Download cost workbook", f.read(), file_name=os.path.basename(wb),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch")


def render_cost_analysis(block1_data: dict, template_path: str | None = None, output_dir: str | None = None):
    """Convenience: build inputs + run, for a standalone page."""
    cost_parts, cost_joints, _ = build_cost_inputs(block1_data)
    render_cost_run(cost_parts, cost_joints, template_path, output_dir)
