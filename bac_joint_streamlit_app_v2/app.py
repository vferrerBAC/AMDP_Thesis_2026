from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

import pandas as pd
import streamlit as st

from engine.connection_templates import (
    apply_template_to_rows,
    connection_schedule_template,
    default_connection_schedule,
    template_names,
    templates_table,
)
from engine.joint_checks import (
    critical_joint_summary,
    default_joint_forces,
    ensure_connection_columns,
    ensure_force_columns,
    evaluate_joint_demands,
    joint_balance_table,
)
from engine.joint_io import (
    connected_members_summary,
    extract_joints_and_members,
    geometry_metrics,
    load_json_from_upload,
)
from engine.load_combinations import default_load_combinations, load_combination_template
from engine.reporting import to_csv_bytes, to_excel_bytes, to_json_bytes, to_pdf_bytes
from engine.validation import is_ready_to_run, validation_checklist

st.set_page_config(
    page_title="BAC Integrated Joint Evaluation Tool (IJET)",
    page_icon="🔩",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
.block-container {padding-top: 1.1rem;}
div[data-testid="stMetric"] {background: #F4F7FB; border: 1px solid #E5EAF2; padding: 12px; border-radius: 12px;}
.small-card {background: #F4F7FB; padding: 14px 16px; border-radius: 12px; border: 1px solid #E5EAF2;}
.ok {color: #138A36; font-weight: 700;}
.warn {color: #A66F00; font-weight: 700;}
.fail {color: #C62828; font-weight: 700;}
.gray {color: #667085; font-weight: 700;}
</style>
""",
    unsafe_allow_html=True,
)


def empty_joints() -> pd.DataFrame:
    return pd.DataFrame(columns=["joint_id", "x", "y", "z"])


def empty_members() -> pd.DataFrame:
    return pd.DataFrame(columns=["member_id", "start_joint", "end_joint", "section_type", "material"])


def init_state() -> None:
    defaults = {
        "project_name": "BAC Joint Analysis",
        "analyst": "",
        "product_family": "Cooling tower / BAC assembly",
        "units": "inch-lbf",
        "design_method": "LRFD",
        "steel_code": "AISI S100 / AISC 360 / Custom",
        "code_edition": "Confirm with project requirements",
        "app_mode": "Simple",
        "risk_category": "II",
        "environment_preset": "Outdoor / normal wind",
        "exposure_category": "C",
        "basic_wind_speed_mph": 115.0,
        "topographic_factor_kzt": 1.0,
        "directionality_factor_kd": 0.85,
        "seismic_sds": 0.0,
        "seismic_sd1": 0.0,
        "corrosion_environment": "Galvanized / outdoor typical",
        "storage_mode": "Local / memory only",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)
    st.session_state.setdefault("joints", empty_joints())
    st.session_state.setdefault("members", empty_members())
    st.session_state.setdefault("combo_table", default_load_combinations(st.session_state.design_method))
    st.session_state.setdefault("connection_schedule", default_connection_schedule(["J001", "J002", "J003"]))
    st.session_state.setdefault("joint_forces", default_joint_forces(["J001", "J002", "J003"], ["LRFD-01", "LRFD-02"]))
    st.session_state.setdefault("analysis_results", pd.DataFrame())
    st.session_state.setdefault("critical_summary", pd.DataFrame())
    st.session_state.setdefault("balance_table", pd.DataFrame())


init_state()


def joint_ids() -> list[str]:
    if st.session_state.joints.empty:
        return []
    return st.session_state.joints["joint_id"].astype(str).tolist()


def combo_ids() -> list[str]:
    if st.session_state.combo_table.empty or "combo_id" not in st.session_state.combo_table.columns:
        return []
    return st.session_state.combo_table["combo_id"].astype(str).tolist()


def build_project_config() -> Dict[str, Any]:
    return {
        "app_mode": st.session_state.app_mode,
        "project": {
            "project_name": st.session_state.project_name,
            "analyst": st.session_state.analyst,
            "product_family": st.session_state.product_family,
            "units": st.session_state.units,
            "design_method": st.session_state.design_method,
            "steel_code": st.session_state.steel_code,
            "code_edition": st.session_state.code_edition,
            "created_at": datetime.now().isoformat(timespec="seconds"),
        },
        "environmental_inputs": {
            "preset": st.session_state.environment_preset,
            "risk_category": st.session_state.risk_category,
            "exposure_category": st.session_state.exposure_category,
            "basic_wind_speed_mph": st.session_state.basic_wind_speed_mph,
            "topographic_factor_kzt": st.session_state.topographic_factor_kzt,
            "directionality_factor_kd": st.session_state.directionality_factor_kd,
            "seismic_sds": st.session_state.seismic_sds,
            "seismic_sd1": st.session_state.seismic_sd1,
            "corrosion_environment": st.session_state.corrosion_environment,
        },
        "privacy": {
            "storage_mode": st.session_state.storage_mode,
            "raw_geometry_saved_by_app": False,
            "note": "Local prototype. Uploaded files are used in memory unless a developer adds saving logic.",
        },
        "load_combinations": st.session_state.combo_table.to_dict("records"),
        "connection_schedule": st.session_state.connection_schedule.to_dict("records"),
    }


def header(title: str, caption: str) -> None:
    st.subheader(title)
    st.caption(caption)


def apply_environment_preset(preset: str) -> None:
    if preset == "Indoor / dry":
        st.session_state.basic_wind_speed_mph = 0.0
        st.session_state.exposure_category = "B"
        st.session_state.seismic_sds = 0.0
        st.session_state.seismic_sd1 = 0.0
        st.session_state.corrosion_environment = "Interior / dry"
    elif preset == "Outdoor / normal wind":
        st.session_state.basic_wind_speed_mph = 115.0
        st.session_state.exposure_category = "C"
        st.session_state.seismic_sds = 0.0
        st.session_state.seismic_sd1 = 0.0
        st.session_state.corrosion_environment = "Galvanized / outdoor typical"
    elif preset == "Coastal / corrosive":
        st.session_state.basic_wind_speed_mph = 130.0
        st.session_state.exposure_category = "C"
        st.session_state.seismic_sds = 0.0
        st.session_state.seismic_sd1 = 0.0
        st.session_state.corrosion_environment = "Coastal / corrosive — review coating and stainless options"
    elif preset == "High wind or seismic review":
        st.session_state.basic_wind_speed_mph = 145.0
        st.session_state.exposure_category = "D"
        st.session_state.seismic_sds = 0.75
        st.session_state.seismic_sd1 = 0.30
        st.session_state.corrosion_environment = "Project-specific"


def status_badge(status: str) -> str:
    mapping = {
        "Passed": "✅ Passed",
        "Warning": "⚠️ Warning",
        "Blocked": "⛔ Blocked",
        "OK": "🟢 OK",
        "WarningResult": "🟡 Warning",
        "Not OK": "🔴 Not OK",
        "Incomplete": "⚪ Incomplete",
        "Ignored": "⚫ Ignored",
    }
    return mapping.get(status, status)


def draw_validation_summary() -> pd.DataFrame:
    checklist = validation_checklist(
        st.session_state.joints,
        st.session_state.members,
        st.session_state.combo_table,
        st.session_state.connection_schedule,
        st.session_state.joint_forces,
        st.session_state.units,
    )
    ready, message = is_ready_to_run(checklist)
    if ready:
        st.success(message, icon="✅")
    else:
        st.error(message, icon="⛔")
    return checklist


# Sidebar
st.sidebar.title("🔩 BAC Integrated Joint Evaluation Tool (IJET)")
st.sidebar.caption("Design-engineer-friendly joint screening")
st.sidebar.radio("User mode", ["Simple", "Advanced"], key="app_mode", help="Simple hides code-level inputs. Advanced exposes load combinations and engineering details.")
step = st.sidebar.radio(
    "Workflow",
    [
        "1️⃣ Project",
        "2️⃣ Geometry",
        "3️⃣ Environment",
        "4️⃣ Connections",
        "5️⃣ Loads",
        "6️⃣ Validate",
        "7️⃣ Results",
        "8️⃣ Export",
    ],
)
st.sidebar.divider()
with st.sidebar.expander("What is safe to do here?", expanded=False):
    st.write("Use sample files freely. For proprietary CAD, run locally and avoid cloud deployment. The prototype does not save uploaded geometry unless you add file-writing code.")

# Create tabs at the top level, before any step content
tab1, tab2, tab3, tab4 = st.tabs(["🏗️ Structural Analysis", "💰 Financial Analysis", "🏭 Manufacturability Analysis", "📋 Summary"])


def draw_structural_analysis_tab() -> None:
    st.title("BAC Integrated Joint Evaluation Tool (IJET)")
    st.caption("A guided Streamlit dashboard for joint-only structural screening — built for design engineers, with advanced details available when needed.")
    metrics = geometry_metrics(st.session_state.joints, st.session_state.members)
    checklist_sidebar = validation_checklist(
        st.session_state.joints,
        st.session_state.members,
        st.session_state.combo_table,
        st.session_state.connection_schedule,
        st.session_state.joint_forces,
        st.session_state.units,
    )
    ready, ready_msg = is_ready_to_run(checklist_sidebar)
    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
    mc1.metric("Joints", metrics["n_joints"])
    mc2.metric("Members", metrics["n_members"])
    mc3.metric("Combos", len(st.session_state.combo_table))
    mc4.metric("Connection rows", len(st.session_state.connection_schedule))
    mc5.metric("Ready", "Yes" if ready else "No")
    st.divider()

    # STEP 1
    if step.startswith("1"):
        header("1. Project setup", "Tell the app what project this is and how much structural detail to expose.")
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Project name", key="project_name", help="Used in the exported report header.")
            st.text_input("Analyst / owner", key="analyst", help="Optional. Name of the person preparing this run.")
            st.selectbox("Product family", ["Cooling tower / BAC assembly", "Heat exchanger support", "Sheet metal frame", "Tube frame", "Custom product"], key="product_family")
            st.selectbox("Project units", ["inch-lbf", "ft-kip", "mm-N", "m-kN"], key="units", help="The current prototype is strongest in inch-lbf. Add full unit conversion before design release.")
        with c2:
            st.markdown("### What the user sees")
            st.write("Simple Mode hides code equations and focuses on geometry, connection templates, loads, and results.")
            if st.session_state.app_mode == "Advanced":
                old_method = st.session_state.design_method
                st.selectbox("Design method", ["LRFD", "ASD"], key="design_method")
                st.text_input("Primary code/check family", key="steel_code")
                st.text_input("Code edition / project note", key="code_edition")
                if st.session_state.design_method != old_method:
                    st.session_state.combo_table = default_load_combinations(st.session_state.design_method)
                    st.toast("Load-combination table updated for selected design method.")
            else:
                st.selectbox("Design intent", ["Preliminary screening", "Engineering review package", "Manufacturing design support"], help="This does not change calculations yet; it controls report language.")
                with st.expander("Engineering details used in the background"):
                    st.write(f"Design method: **{st.session_state.design_method}**")
                    st.write(f"Check family: **{st.session_state.steel_code}**")
                    st.write(f"Note: **{st.session_state.code_edition}**")
        st.warning("This tool is meant for preliminary structural screening only, designs should stil be approved by a structural engineer.", icon="⚠️")

    # STEP 2
    elif step.startswith("2"):
        header("2. Upload geometry", "Upload the CAD-derived Block 1 JSON. The app extracts joint coordinates and member connectivity.")
        left, right = st.columns([1, 1])
        with left:
            uploaded_json = st.file_uploader("Upload Block 1 JSON", type=["json"], help="Use the JSON produced by your Inventor/CAD joint identification code.")
            if uploaded_json is not None:
                try:
                    data = load_json_from_upload(uploaded_json)
                    joints, members = extract_joints_and_members(data)
                    st.session_state.joints = joints
                    st.session_state.members = members
                    ids = joint_ids()
                    st.session_state.connection_schedule = default_connection_schedule(ids)
                    st.session_state.joint_forces = default_joint_forces(ids, combo_ids())
                    st.success(f"Loaded {len(joints)} joints and {len(members)} members.")
                except Exception as exc:
                    st.error(f"Could not read JSON: {exc}")
            if st.button("Load sample geometry", type="primary", use_container_width=True):
                data = json.loads(Path("sample_data/block1_sample_geometry.json").read_text())
                joints, members = extract_joints_and_members(data)
                st.session_state.joints = joints
                st.session_state.members = members
                ids = joint_ids()
                st.session_state.connection_schedule = default_connection_schedule(ids)
                st.session_state.joint_forces = default_joint_forces(ids, combo_ids())
                st.success("Sample geometry loaded.")
        with right:
            st.markdown("### Geometry preview")
            if st.session_state.joints.empty:
                st.info("No geometry loaded yet.")
            else:
                plot_df = st.session_state.joints.rename(columns={"x": "X", "y": "Y"})
                st.scatter_chart(plot_df, x="X", y="Y", size=80, color=None)
        st.markdown("### Detected joints and connected members")
        summary = connected_members_summary(st.session_state.joints, st.session_state.members)
        st.dataframe(summary, use_container_width=True, hide_index=True)
        with st.expander("Raw joint coordinate table"):
            st.dataframe(st.session_state.joints, use_container_width=True, hide_index=True)
        with st.expander("Raw member connectivity table"):
            st.dataframe(st.session_state.members, use_container_width=True, hide_index=True)

    # STEP 3
    elif step.startswith("3"):
        header("3. Environment", "Simple choices set default environmental factors; advanced users can edit the technical values.")
        previous = st.session_state.environment_preset
        st.selectbox("Environment preset", ["Indoor / dry", "Outdoor / normal wind", "Coastal / corrosive", "High wind or seismic review", "Custom"], key="environment_preset", help="Choose the option closest to where the equipment will operate.")
        if st.session_state.environment_preset != previous and st.session_state.environment_preset != "Custom":
            apply_environment_preset(st.session_state.environment_preset)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("### Plain-language meaning")
            st.write("The environment affects which loads and durability assumptions should be reviewed.")
            st.write(f"Selected preset: **{st.session_state.environment_preset}**")
            st.write(f"Corrosion note: **{st.session_state.corrosion_environment}**")
        with c2:
            st.markdown("### Wind")
            disabled = st.session_state.app_mode != "Advanced" and st.session_state.environment_preset != "Custom"
            st.selectbox("Risk category", ["I", "II", "III", "IV"], key="risk_category", disabled=disabled)
            st.number_input("Basic wind speed, V (mph)", min_value=0.0, step=5.0, key="basic_wind_speed_mph", disabled=disabled, help="Use project-specified wind speed when available.")
            st.selectbox("Exposure category", ["B", "C", "D"], key="exposure_category", disabled=disabled)
        with c3:
            st.markdown("### Seismic / other")
            st.number_input("SDS", min_value=0.0, step=0.05, key="seismic_sds", disabled=disabled)
            st.number_input("SD1", min_value=0.0, step=0.05, key="seismic_sd1", disabled=disabled)
            st.text_input("Corrosion / coating note", key="corrosion_environment", disabled=disabled)
        if st.session_state.app_mode == "Advanced":
            with st.expander("Advanced wind factors"):
                st.number_input("Topographic factor, Kzt", min_value=0.0, step=0.05, key="topographic_factor_kzt")
                st.number_input("Directionality factor, Kd", min_value=0.0, step=0.05, key="directionality_factor_kd")

    # STEP 4
    elif step.startswith("4"):
        header("4. Connection templates", "Assign each joint a connection type using design-engineer-friendly templates.")
        st.markdown("### Template guide")
        guide = templates_table()[["template_name", "user_friendly_name", "when_to_use", "recommended_data_needed"]]
        st.dataframe(guide, use_container_width=True, hide_index=True)
        if st.session_state.joints.empty:
            st.warning("Load geometry first so the app knows which joints need connection data.")
        c1, c2 = st.columns([1, 2])
        with c1:
            selected_template = st.selectbox("Template to apply", template_names(), help="Choose the physical connection detail that best matches the joint.")
            selected_joints = st.multiselect("Apply to joints", joint_ids(), default=joint_ids()[: min(5, len(joint_ids()))])
            if st.button("Apply template", type="primary", use_container_width=True, disabled=not bool(selected_joints)):
                st.session_state.connection_schedule = apply_template_to_rows(st.session_state.connection_schedule, selected_joints, selected_template)
                st.success(f"Applied {selected_template} to {len(selected_joints)} joint(s).")
            if st.button("Reset schedule from default template", use_container_width=True):
                st.session_state.connection_schedule = default_connection_schedule(joint_ids(), selected_template)
                st.success("Connection schedule reset.")
        with c2:
            st.markdown("### Editable connection schedule")
            st.caption("Design engineers can edit this like a spreadsheet. Advanced equation details stay hidden inside the engine.")
            st.session_state.connection_schedule = ensure_connection_columns(st.data_editor(
                ensure_connection_columns(st.session_state.connection_schedule),
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "joint_id": st.column_config.TextColumn("Joint ID", help="Must match a detected joint ID."),
                    "template_name": st.column_config.SelectboxColumn("Template", options=template_names(), help="Friendly connection template."),
                    "connection_type": st.column_config.SelectboxColumn("Check type", options=["screwed_sheet_joint", "bolted_bracket_joint", "welded_joint", "custom_capacity", "ignore"], help="Calculation path used by the engine."),
                    "n_fasteners": st.column_config.NumberColumn("# Fasteners", min_value=0, step=1),
                    "diameter_in": st.column_config.NumberColumn("Diameter (in)", min_value=0.0, format="%.3f"),
                    "sheet_t_in": st.column_config.NumberColumn("Thickness (in)", min_value=0.0, format="%.4f"),
                    "edge_dist_in": st.column_config.NumberColumn("Edge dist. (in)", min_value=0.0, format="%.3f"),
                    "spacing_in": st.column_config.NumberColumn("Spacing (in)", min_value=0.0, format="%.3f"),
                    "weld_size_in": st.column_config.NumberColumn("Weld size (in)", min_value=0.0, format="%.3f"),
                    "weld_length_in": st.column_config.NumberColumn("Weld length (in)", min_value=0.0, format="%.3f"),
                    "custom_capacity_lbf": st.column_config.NumberColumn("Custom capacity (lbf)", min_value=0.0, format="%.1f"),
                },
            ))
        st.download_button("Download blank connection schedule template", to_csv_bytes(connection_schedule_template()), "connection_schedule_template.csv", "text/csv")

    # STEP 5
    elif step.startswith("5"):
        header("5. Loads and combinations", "In Simple Mode, upload or edit joint force demands. In Advanced Mode, edit load combinations too.")
        c1, c2 = st.columns(2)
        with c1:
            uploaded_forces = st.file_uploader("Upload joint force CSV", type=["csv"], help="Expected columns: joint_id, combo_id, Fx_lbf, Fy_lbf, Fz_lbf, Mx_lbf_in, My_lbf_in, Mz_lbf_in.")
            if uploaded_forces is not None:
                try:
                    st.session_state.joint_forces = ensure_force_columns(pd.read_csv(uploaded_forces))
                    st.success("Joint force table uploaded.")
                except Exception as exc:
                    st.error(f"Could not read force CSV: {exc}")
            if st.button("Generate demo joint loads", use_container_width=True):
                st.session_state.joint_forces = default_joint_forces(joint_ids(), combo_ids())
                st.success("Demo force demands generated.")
        with c2:
            st.download_button("Download force input template", to_csv_bytes(default_joint_forces(["J001", "J002"], ["LRFD-01", "LRFD-02"])), "joint_force_template.csv", "text/csv", use_container_width=True)
            st.download_button("Download load-combination template", to_csv_bytes(load_combination_template()), "load_combination_template.csv", "text/csv", use_container_width=True)
            uploaded_combos = st.file_uploader("Upload load combinations CSV", type=["csv"], help="Advanced input. Expected columns: combo_id, combo_name, expression, plain_language_notes.")
            if uploaded_combos is not None:
                try:
                    st.session_state.combo_table = pd.read_csv(uploaded_combos)
                    st.success("Load-combination table uploaded.")
                except Exception as exc:
                    st.error(f"Could not read combo CSV: {exc}")
        if st.session_state.app_mode == "Advanced":
            st.markdown("### Editable load combinations")
            st.session_state.combo_table = st.data_editor(st.session_state.combo_table, use_container_width=True, hide_index=True, num_rows="dynamic")
        else:
            with st.expander("View load combinations used in the background"):
                st.dataframe(st.session_state.combo_table, use_container_width=True, hide_index=True)
        st.markdown("### Editable joint force demands")
        st.session_state.joint_forces = ensure_force_columns(st.data_editor(
            ensure_force_columns(st.session_state.joint_forces),
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "joint_id": st.column_config.SelectboxColumn("Joint ID", options=joint_ids() or ["J001", "J002", "J003"]),
                "combo_id": st.column_config.SelectboxColumn("Combo ID", options=combo_ids() or ["LRFD-01", "LRFD-02"]),
            },
        ))

    # STEP 6
    elif step.startswith("6"):
        header("6. Validate before running", "This prevents crashes and tells the user what needs to be fixed in plain language.")
        checklist = draw_validation_summary()
        st.dataframe(checklist[["check", "status", "message", "recommended_action"]], use_container_width=True, hide_index=True)
        st.markdown("### Why validation matters")
        st.write("A design engineer should not have to debug Python errors. The app should identify missing geometry, incomplete connection rows, inconsistent units, or force rows that reference the wrong joint.")

    # STEP 7
    elif step.startswith("7"):
        header("7. Run and review results", "The first output is plain-language status. Technical tables are available underneath.")
        checklist = validation_checklist(st.session_state.joints, st.session_state.members, st.session_state.combo_table, st.session_state.connection_schedule, st.session_state.joint_forces, st.session_state.units)
        ready, message = is_ready_to_run(checklist)
        if not ready:
            st.error(message)
            st.dataframe(checklist[["check", "status", "message", "recommended_action"]], use_container_width=True, hide_index=True)
        run_anyway = st.checkbox("Advanced: run even with warnings/blockers", value=False, disabled=ready or st.session_state.app_mode != "Advanced")
        if st.button("Run joint screening", type="primary", use_container_width=True, disabled=(not ready and not run_anyway)):
            st.session_state.analysis_results = evaluate_joint_demands(st.session_state.joint_forces, st.session_state.connection_schedule)
            st.session_state.critical_summary = critical_joint_summary(st.session_state.analysis_results)
            st.session_state.balance_table = joint_balance_table(st.session_state.joint_forces)
            st.success("Joint screening complete.")
        results = st.session_state.analysis_results
        if results.empty:
            st.info("No results yet. Run the joint screening after validation passes.")
        else:
            counts = results["status"].value_counts().to_dict()
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("OK", counts.get("OK", 0))
            c2.metric("Warnings", counts.get("Warning", 0))
            c3.metric("Not OK", counts.get("Not OK", 0))
            c4.metric("Incomplete", counts.get("Incomplete", 0))
            c5.metric("Ignored", counts.get("Ignored", 0))
            if counts.get("Not OK", 0) > 0:
                st.error("Overall result: review required. At least one joint demand exceeds the demo screening capacity.", icon="🔴")
            elif counts.get("Warning", 0) > 0:
                st.warning("Overall result: close to limit. Review yellow joints before release.", icon="🟡")
            elif counts.get("Incomplete", 0) > 0:
                st.warning("Overall result: incomplete. Some joints need more input data.", icon="⚪")
            else:
                st.success("Overall result: all checked joints are green in the current screening model.", icon="🟢")
            st.markdown("### Critical joint summary")
            st.dataframe(st.session_state.critical_summary, use_container_width=True, hide_index=True)
            st.markdown("### Detailed results")
            status_filter = st.multiselect("Filter by status", sorted(results["status"].dropna().unique().tolist()), default=sorted(results["status"].dropna().unique().tolist()))
            view = results[results["status"].isin(status_filter)] if status_filter else results
            friendly_cols = ["joint_id", "combo_id", "status", "percent_used", "plain_language_issue", "suggested_fix", "force_demand_lbf", "screening_capacity_lbf"]
            st.dataframe(view[friendly_cols], use_container_width=True, hide_index=True)
            with st.expander("Engineering details: forces and capacity basis"):
                st.dataframe(view, use_container_width=True, hide_index=True)
            with st.expander("SAP-like joint balance / reaction-style table"):
                st.dataframe(st.session_state.balance_table, use_container_width=True, hide_index=True)

    # STEP 8
    elif step.startswith("8"):
        header("8. Export", "Create files that a design engineer can share with a reviewer or keep as a design record.")
        checklist = validation_checklist(st.session_state.joints, st.session_state.members, st.session_state.combo_table, st.session_state.connection_schedule, st.session_state.joint_forces, st.session_state.units)
        config = build_project_config()
        sheets = {
            "Validation": checklist[["check", "status", "message", "recommended_action"]],
            "Joints": st.session_state.joints,
            "Members": st.session_state.members,
            "Connection Schedule": st.session_state.connection_schedule,
            "Load Combinations": st.session_state.combo_table,
            "Joint Forces": st.session_state.joint_forces,
            "Results": st.session_state.analysis_results,
            "Critical Summary": st.session_state.critical_summary,
            "Joint Balance": st.session_state.balance_table,
        }
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("Download Excel workbook", to_excel_bytes(sheets), "bac_joint_screening_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            st.download_button("Download project config JSON", to_json_bytes(config), "project_joint_config.json", "application/json", use_container_width=True)
        with c3:
            try:
                pdf_bytes = to_pdf_bytes(config, st.session_state.critical_summary, st.session_state.analysis_results, checklist)
                st.download_button("Download PDF summary", pdf_bytes, "bac_joint_screening_report.pdf", "application/pdf", use_container_width=True)
            except Exception as exc:
                st.warning(f"PDF export unavailable: {exc}")
        st.markdown("### Templates")
        t1, t2, t3 = st.columns(3)
        t1.download_button("Connection schedule CSV", to_csv_bytes(connection_schedule_template()), "connection_schedule_template.csv", "text/csv", use_container_width=True)
        t2.download_button("Joint force CSV", to_csv_bytes(default_joint_forces(["J001", "J002"], ["LRFD-01", "LRFD-02"])), "joint_force_template.csv", "text/csv", use_container_width=True)
        t3.download_button("Load combinations CSV", to_csv_bytes(load_combination_template()), "load_combination_template.csv", "text/csv", use_container_width=True)
        st.info("Exported reports include a clear limitation statement because the current capacities are placeholders.")

with tab1:
    draw_structural_analysis_tab()

with tab2:
    st.header("Financial Analysis")
    st.write("Financial Analysis content will be added here.")

with tab3:
    st.header("Manufacturability Analysis")
    st.write("Manufacturability Analysis content will be added here.")

with tab4:
    st.header("Summary")
    st.write("This summary tab can be used to present final observations and export notes.")

