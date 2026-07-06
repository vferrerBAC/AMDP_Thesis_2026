from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

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
from engine.bolted_capacity import MATERIALS, BOLTS, default_bolted_row, map_material, map_bolt

# Block 3 (PyNite FE solve) integration. Imported defensively so the rest of the
# app still loads even if the optional dependency is missing.
try:
    from engine.loads.asce7_lrfd_loads import (
        combos_table,
        joint_forces_table,
        members_dataframe_from_block1,
        run_lrfd_joint_analysis,
    )
    BLOCK3_AVAILABLE = True
    BLOCK3_IMPORT_ERROR = ""
except Exception as _exc:  # pragma: no cover
    BLOCK3_AVAILABLE = False
    BLOCK3_IMPORT_ERROR = str(_exc)

# Block 1 (live Inventor extraction) integration. Imported defensively: the
# module is platform-safe (win32 is guarded), so this import succeeds on any
# host. Only run()/extract_live_to_dict() require Windows + a live session.
try:
    from engine import block1 as block1_pipeline
    BLOCK1_AVAILABLE = True
    BLOCK1_IMPORT_ERROR = ""
except Exception as _exc:  # pragma: no cover
    BLOCK1_AVAILABLE = False
    BLOCK1_IMPORT_ERROR = str(_exc)

st.set_page_config(
    page_title="BAC Integrated Joint Evaluation Tool (IJET)",
    page_icon="🔩",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
/* ── Layout ── */
.block-container {padding-top: 0.75rem;}

/* ── Sidebar dark theme ── */
[data-testid="stSidebar"] {background-color: #1a2332 !important;}
[data-testid="stSidebar"] > div:first-child {background-color: #1a2332 !important;}
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
[data-testid="stSidebar"] label {color: #c8d3e8 !important;}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {color: #ffffff !important;}
[data-testid="stSidebar"] small, [data-testid="stSidebar"] .stCaption p {
    color: #7b8fb3 !important;}
[data-testid="stSidebar"] hr {border-color: #2d3a52 !important;}
[data-testid="stSidebar"] [data-testid="stExpander"] {
    background-color: #243040 !important; border: 1px solid #2d3a52 !important;}
[data-testid="stSidebar"] [data-testid="stRadio"] > div {gap: 2px;}
[data-testid="stSidebar"] [data-testid="stRadio"] label {
    font-size: 13.5px !important; padding: 6px 8px !important;
    border-radius: 6px !important;}

/* ── Stage banner ── */
.stage-banner {
    background: linear-gradient(135deg, #1e3a8a 0%, #1d4ed8 100%);
    color: white; padding: 20px 28px; border-radius: 12px; margin-bottom: 20px;}
.stage-banner h2 {margin: 0 0 5px; font-size: 26px; font-weight: 700; color: white !important;}
.stage-banner p {margin: 0; opacity: 0.88; font-size: 14px; color: white !important;}

/* ── Metric cards ── */
.mc-row {display: flex; gap: 14px; margin-bottom: 20px;}
.mc {
    flex: 1; background: white; border: 1px solid #e5eaf2;
    border-radius: 10px; padding: 16px 20px; min-width: 0;}
.mc-label {
    font-size: 11px; font-weight: 700; color: #6b7280;
    letter-spacing: 0.06em; text-transform: uppercase; margin-bottom: 6px;}
.mc-value {font-size: 30px; font-weight: 700; color: #111827; line-height: 1;}
.mc-blue   {border-left: 4px solid #2563eb;}
.mc-green  {border-left: 4px solid #16a34a;}
.mc-teal   {border-left: 4px solid #0d9488;}
.mc-yellow {border-left: 4px solid #d97706;}
.mc-red    {border-left: 4px solid #dc2626;}
.mc-gray   {border-left: 4px solid #6b7280;}

/* ── Source badge (sidebar footer) ── */
.src-badge {
    background: #243040; border-radius: 8px; padding: 10px 14px; margin-top: 6px;}
.src-badge .src-label {
    font-size: 10px; color: #7b8fb3; text-transform: uppercase;
    letter-spacing: 0.05em; margin-bottom: 3px;}
.src-badge .src-value {font-size: 13px; font-weight: 600; color: #c8d3e8;}

/* ── Legacy status colours ── */
div[data-testid="stMetric"] {background: #F4F7FB; border: 1px solid #E5EAF2; padding: 12px; border-radius: 12px;}
.small-card {background: #F4F7FB; padding: 14px 16px; border-radius: 12px; border: 1px solid #E5EAF2;}
.ok   {color: #138A36; font-weight: 700;}
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
        "steel_code": "AISI S100 / ASCE 7-22 / Custom",
        "code_edition": "Confirm with project requirements",
        "app_mode": "Advanced",
        "seismic_sds": 0.0,
        "storage_mode": "Local / memory only",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)
    st.session_state.setdefault("joints", empty_joints())
    st.session_state.setdefault("members", empty_members())
    st.session_state.setdefault("block1_raw", None)
    st.session_state.setdefault("block3_result", None)
    st.session_state.setdefault("wind_loads_df", pd.DataFrame([{"node": "", "FX": 0.0, "FY": 0.0, "FZ": 0.0}]))
    st.session_state.setdefault("seismic_loads_df", pd.DataFrame([{"node": "", "FX": 0.0, "FY": 0.0, "FZ": 0.0}]))
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


def _thickness_map_from_block1(block1_raw: dict, members_df: pd.DataFrame) -> dict:
    """Return {joint_id: min_wall_thickness_in} from Block 1 member cross-sections.

    Tries the real Block 1 schema first (joints[].member_names → member.cross_section
    .wall_thickness), then falls back to the simple schema where members carry
    start_joint / end_joint directly.  Returns an empty dict when no thickness
    data is present (e.g., the minimal sample geometry file).
    """
    try:
        result: dict = {}
        raw_members = block1_raw.get("members", [])

        # Real Block 1 schema: each joint lists connected member occurrence paths.
        # _name_from_contact prefers occurrence_path over occurrence_name, so index
        # by both to match regardless of assembly nesting depth.
        member_lookup: dict = {}
        for m in raw_members:
            if not isinstance(m, dict):
                continue
            path = m.get("occurrence_path", "")
            name = m.get("occurrence_name", "")
            if path:
                member_lookup[path] = m
            if name and name not in member_lookup:
                member_lookup[name] = m

        for j in block1_raw.get("joints", []):
            if not isinstance(j, dict):
                continue
            jid = str(j.get("joint_id", ""))
            if not jid:
                continue
            thicknesses = []
            for mname in (j.get("member_names") or []):
                m = member_lookup.get(mname)
                if not isinstance(m, dict):
                    continue
                cs = m.get("cross_section")
                if not isinstance(cs, dict):
                    continue
                try:
                    t = cs.get("wall_thickness")
                    if t is not None and float(t) > 0:
                        thicknesses.append(float(t))
                except (TypeError, ValueError):
                    pass
            if thicknesses:
                result[jid] = min(thicknesses)

        # Fallback: simple schema — members carry start_joint / end_joint + cross_section.
        if not result and not members_df.empty:
            for raw_m in raw_members:
                if not isinstance(raw_m, dict):
                    continue
                cs = raw_m.get("cross_section")
                if not isinstance(cs, dict):
                    continue
                try:
                    thickness = cs.get("wall_thickness")
                    if thickness is None:
                        continue
                    thickness = float(thickness)
                    if thickness <= 0:
                        continue
                except (TypeError, ValueError):
                    continue
                mid = str(raw_m.get("member_id", raw_m.get("occurrence_name", "")))
                matches = members_df[members_df["member_id"].astype(str) == mid]
                for _, mrow in matches.iterrows():
                    for jid in (str(mrow["start_joint"]), str(mrow["end_joint"])):
                        if jid and (jid not in result or thickness < result[jid]):
                            result[jid] = thickness

        return result
    except Exception:
        return {}


def ingest_block1_dict(data: Dict[str, Any]) -> tuple[int, int]:
    """Single ingest path shared by upload, sample, and live-Inventor extraction.

    Keeps all three sources byte-identical downstream: parse joints/members,
    recover connectivity from the real Block-1 schema when the generic parser
    finds no members, then seed the connection schedule and force table.
    Returns (n_joints, n_members).
    """
    joints, members = extract_joints_and_members(data)
    st.session_state.block1_raw = data
    if members.empty and BLOCK3_AVAILABLE:
        members = members_dataframe_from_block1(data)
    st.session_state.joints = joints
    st.session_state.members = members
    ids = joint_ids()
    # Seed with bolted defaults (not the screw template) so Step 3 shows the
    # right fastener/diameter/edge-distance starting values.
    sched = ensure_connection_columns(pd.DataFrame([default_bolted_row(j) for j in ids]))

    # Autopopulate sheet_t_in from Block 1 member wall thicknesses.
    t_map = _thickness_map_from_block1(data, members)
    if t_map:
        for i, row in sched.iterrows():
            jid = str(row["joint_id"])
            if jid in t_map:
                sched.at[i, "sheet_t_in"] = t_map[jid]

    st.session_state.connection_schedule = sched

    st.session_state.joint_forces = default_joint_forces(ids, combo_ids())
    return len(joints), len(members)


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
            "seismic_sds": st.session_state.seismic_sds,
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
    parts = title.split(". ", 1)
    num, name = (parts[0], parts[1]) if len(parts) == 2 else ("", title)
    st.markdown(
        f'<div class="stage-banner"><h2>Stage {num} · {name}</h2><p>{caption}</p></div>',
        unsafe_allow_html=True,
    )


def metric_cards_html(cards: list) -> str:
    """Build a row of colored-border metric cards. cards = [(label, value, color_class), ...]"""
    items = "".join(
        f'<div class="mc mc-{c}"><div class="mc-label">{lbl}</div>'
        f'<div class="mc-value">{val}</div></div>'
        for lbl, val, c in cards
    )
    return f'<div class="mc-row">{items}</div>'


def geometry_figure(joints_df: pd.DataFrame, members_df: pd.DataFrame):
    """Plotly 3D scatter with joint labels and member lines."""
    if not PLOTLY_AVAILABLE or joints_df.empty:
        return None

    coord = {
        str(r["joint_id"]): (float(r["x"]), float(r["y"]), float(r["z"]))
        for _, r in joints_df.iterrows()
    }

    # Compute per-axis spans to detect the flat frame's orientation.
    x_span = float(joints_df["x"].max() - joints_df["x"].min()) or 1.0
    y_span = float(joints_df["y"].max() - joints_df["y"].min()) or 1.0
    z_span = float(joints_df["z"].max() - joints_df["z"].min()) or 1.0
    spans = {"x": x_span, "y": y_span, "z": z_span}
    thin_axis = min(spans, key=spans.get)
    max_span = max(x_span, y_span, z_span)

    # Aspect ratios: proportional to actual spans, thin axis gets at least 8%.
    ar_x = max(x_span / max_span, 0.08)
    ar_y = max(y_span / max_span, 0.08)
    ar_z = max(z_span / max_span, 0.08)

    # Camera: look face-on at the frame by positioning the eye along the thin
    # axis.  "up" is determined by convention, not by max span — the max-span
    # axis is typically the horizontal width, which should not be screen-up.
    # XY-plane frames (thin=Z): Y is the vertical height (Inventor/CAD Y-up).
    # XZ-plane frames (thin=Y): Z is the vertical height (structural Z-up).
    # YZ-plane frames (thin=X): Z is the vertical height (structural Z-up).
    thin_to_up = {"z": "y", "y": "z", "x": "z"}
    up_axis = thin_to_up[thin_axis]
    up_map = {"x": dict(x=1, y=0, z=0), "y": dict(x=0, y=1, z=0), "z": dict(x=0, y=0, z=1)}
    eye_map = {
        "x": dict(x=2.2, y=0.6, z=0.4),
        "y": dict(x=0.6, y=2.2, z=0.4),
        "z": dict(x=0.6, y=0.4, z=2.2),
    }
    camera = dict(up=up_map[up_axis], eye=eye_map[thin_axis])

    fig = go.Figure()

    # Member lines
    for _, m in members_df.iterrows():
        si = coord.get(str(m.get("start_joint", "")))
        ej = coord.get(str(m.get("end_joint", "")))
        if si and ej:
            fig.add_trace(go.Scatter3d(
                x=[si[0], ej[0], None], y=[si[1], ej[1], None], z=[si[2], ej[2], None],
                mode="lines", line=dict(color="#1d4ed8", width=5),
                showlegend=False, hoverinfo="skip",
            ))

    # Joint markers + labels
    fig.add_trace(go.Scatter3d(
        x=joints_df["x"].tolist(), y=joints_df["y"].tolist(), z=joints_df["z"].tolist(),
        mode="markers+text",
        marker=dict(size=9, color="#1abc9c", line=dict(color="white", width=1.5)),
        text=joints_df["joint_id"].astype(str).tolist(),
        textposition="top center",
        textfont=dict(size=12, color="#111827"),
        name="Joints",
        hovertemplate="%{text}<br>X=%{x:.2f} Y=%{y:.2f} Z=%{z:.2f}<extra></extra>",
    ))

    unit = "in" if (st.session_state.get("units") or "inch-lbf").startswith("inch") else "m"
    fig.update_layout(
        scene=dict(
            xaxis=dict(title=f"X ({unit})", gridcolor="#e5e7eb"),
            yaxis=dict(title=f"Y ({unit})", gridcolor="#e5e7eb"),
            zaxis=dict(title=f"Z ({unit})", gridcolor="#e5e7eb"),
            bgcolor="white",
            aspectmode="manual",
            aspectratio=dict(x=ar_x, y=ar_y, z=ar_z),
            camera=camera,
        ),
        margin=dict(l=0, r=0, t=8, b=0),
        paper_bgcolor="white",
        height=430,
    )
    return fig



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
st.sidebar.title("🔩 IJET")
st.sidebar.caption("Design-engineer-friendly joint screening")
step = st.sidebar.radio(
    "Workflow",
    [
        "1 · Project",
        "2 · Geometry",
        "3 · Connections",
        "4 · Loads",
        "5 · Validate",
        "6 · Results",
        "7 · Export",
    ],
    label_visibility="collapsed",
)
st.sidebar.divider()
with st.sidebar.expander("What is safe to do here?", expanded=False):
    st.write(
        "Use sample files freely. For proprietary CAD, run locally and avoid cloud deployment. "
        "The prototype does not save uploaded geometry unless you add file-writing code."
    )
_src_label = "Inventor" if (BLOCK1_AVAILABLE and st.session_state.get("block1_raw") is not None) else "JSON / notebook mode"
st.sidebar.markdown(
    f'<div class="src-badge"><div class="src-label">Source</div>'
    f'<div class="src-value">{_src_label}</div></div>',
    unsafe_allow_html=True,
)

st.title("BAC Integrated Joint Evaluation Tool (IJET)")
st.caption("Streamlit dashboard for joint-only structural screening.")

tab1, tab2, tab3, tab4 = st.tabs(["🏗️ Structural Analysis", "💰 Financial Analysis", "🏭 Manufacturability Analysis", "📋 Summary"])


def draw_structural_analysis_tab() -> None:
    metrics = geometry_metrics(st.session_state.joints, st.session_state.members)
    checklist_sidebar = validation_checklist(
        st.session_state.joints,
        st.session_state.members,
        st.session_state.combo_table,
        st.session_state.connection_schedule,
        st.session_state.joint_forces,
        st.session_state.units,
    )
    ready, _ = is_ready_to_run(checklist_sidebar)
    st.markdown(metric_cards_html([
        ("Joints", metrics["n_joints"], "blue"),
        ("Members", metrics["n_members"], "green"),
        ("Combos", len(st.session_state.combo_table), "teal"),
        ("Connections", len(st.session_state.connection_schedule), "yellow"),
        ("Ready", "Yes" if ready else "No", "green" if ready else "red"),
    ]), unsafe_allow_html=True)

    # STEP 1
    if step.startswith("1"):
        header("1. Project setup", "Configure project metadata and engineering details.")
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Project name", key="project_name", help="Used in the exported report header.")
            st.text_input("Analyst / owner", key="analyst", help="Optional. Name of the person preparing this run.")
            st.selectbox("Product family", ["Closed Circuit Cooling Towers", "Cooling Tower", "Hybrid and Adiabatic Product", "Evaporative Condensers", "Ice Thermal Storage"], key="product_family")
            st.selectbox("Project units", ["inch-lbf"], key="units", help="The current prototype is strongest in inch-lbf. Add full unit conversion before design release.")
        with c2:
            st.text_input("Design method", value="LRFD", disabled=True, help="LRFD only by project scope. ASD is excluded from the analysis workflow.")
            st.text_input("Primary code/check family", key="steel_code")
            st.text_input("Code edition / project note", key="code_edition")
        st.warning("This tool is meant for preliminary structural screening only, designs should still be approved by a structural engineer.", icon="⚠️")

    # STEP 2
    elif step.startswith("2"):
        header("2. Get geometry", "Extract data directly from an active Inventor session, or upload the CAD-derived Block 1 JSON. For the Inventor path only have one assembly (.iam) open to avoid runtime delays.")

        # ── Metric cards ──────────────────────────────────────────────────────
        _nj = len(st.session_state.joints)
        _nm = len(st.session_state.members)
        _raw = st.session_state.block1_raw or {}
        _n_sup = sum(1 for j in _raw.get("joints", []) if j.get("is_support_candidate", False))
        _meta = (st.session_state.block3_result or {}).get("meta", {})
        _n_flags = len([n for n in _meta.get("build_notes", []) if n.get("needs_review")])
        st.markdown(metric_cards_html([
            ("Nodes", _nj, "blue"),
            ("Elements", _nm, "green"),
            ("Supports", _n_sup, "teal"),
            ("Flags", _n_flags, "yellow"),
        ]), unsafe_allow_html=True)

        status = block1_pipeline.live_inventor_status() if BLOCK1_AVAILABLE else {"available": False, "reason": BLOCK1_IMPORT_ERROR}
        st.markdown("### Live extraction (Block 1 · Inventor)")
        if status["available"]:
            st.success(f"Inventor detected. {status['reason']}", icon="🟢")
        else:
            st.info(f"Live extraction unavailable: {status['reason']} Use the JSON upload below instead.", icon="ℹ️")
        if st.button("Extract from live Inventor", type="primary", use_container_width=True, disabled=not status["available"]):
            try:
                with st.spinner("Running Block 1 against the active assembly…"):
                    data = block1_pipeline.extract_live_to_dict()
                nj, nm = ingest_block1_dict(data)
                u = data.get("units", {})
                st.success(f"Extracted {nj} joints and {nm} members from {Path(str(data.get('source_document',''))).name or 'the active assembly'}.")
                st.caption(f"Canonical units: length={u.get('length','in')}, mass={u.get('mass','lbm')}, force={u.get('force','lbf')}.")
            except Exception as exc:
                st.error(f"Live extraction failed: {exc}")
        with st.expander("Extraction parameters used (read-only)"):
            p = block1_pipeline.EXTRACTION_PARAMS if BLOCK1_AVAILABLE else {}
            st.caption("Frozen Block-1 defaults, tuned on the source model. Shown for the design record; not user-editable here.")
            st.dataframe(pd.DataFrame([
                {"parameter": "Material tags", "value": ", ".join(p.get("tags", ())) or "—", "meaning": "Members kept (GLV/HDG finish)."},
                {"parameter": "tol_touch_in", "value": p.get("tol_touch_in", "—"), "meaning": "Max centerline gap counted as a joint contact."},
                {"parameter": "cluster_cap_in", "value": "None (geometry-derived)" if p.get("cluster_cap_in") is None else p.get("cluster_cap_in"), "meaning": "Absolute cap on merging nearby contacts."},
                {"parameter": "support_z_tol_in", "value": p.get("support_z_tol_in", "—"), "meaning": "Band above lowest Z flagged as support candidate."},
            ]), use_container_width=True, hide_index=True)

        st.divider()
        left, right = st.columns([1, 1])
        with left:
            st.markdown("### Upload / sample (fallback)")
            uploaded_json = st.file_uploader("Upload Block 1 JSON", type=["json"], help="Use the block1_result_unit_safe.json produced by your Inventor/CAD joint identification code.")
            if uploaded_json is not None:
                try:
                    data = load_json_from_upload(uploaded_json)
                    nj, nm = ingest_block1_dict(data)
                    st.success(f"Loaded {nj} joints and {nm} members.")
                except Exception as exc:
                    st.error(f"Could not read JSON: {exc}")
            if st.button("Load sample geometry", use_container_width=True):
                data = json.loads(Path("sample_data/block1_sample_geometry.json").read_text())
                nj, nm = ingest_block1_dict(data)
                st.success("Sample geometry loaded.")
            if st.session_state.block1_raw is not None:
                st.download_button(
                    "Download Block 1 JSON",
                    data=json.dumps(st.session_state.block1_raw, indent=2, default=str).encode("utf-8"),
                    file_name="block1_result.json",
                    mime="application/json",
                    use_container_width=True,
                    help="Download the currently loaded Block 1 geometry as JSON.",
                )
        with right:
            st.markdown("### 3D model preview")
            if st.session_state.joints.empty:
                st.info("No geometry loaded yet.")
            elif PLOTLY_AVAILABLE:
                fig = geometry_figure(st.session_state.joints, st.session_state.members)
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
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
        header("3. Bolted connections",
               "Every joint is checked as a bolted sheet-metal connection (AISC Chapter J). "
               "Two inputs come from the connected parts; the rest follow the BAC standard bolt detail.")

        if st.session_state.joints.empty:
            st.warning("Load geometry in Step 2 first so the app knows which joints need connection data.")
        else:
            ids = joint_ids()

            # Build the working schedule: bolted defaults per joint, preserving any
            # prior edits and normalizing material/bolt to the valid dropdown values.
            prev = {}
            sched = st.session_state.connection_schedule
            if not sched.empty and "joint_id" in sched.columns:
                prev = {str(r["joint_id"]): r for _, r in sched.iterrows()}
            rows = []
            for j in ids:
                row = default_bolted_row(j)
                p = prev.get(str(j))
                if p is not None:
                    for k in ["sheet_t_in", "n_fasteners", "diameter_in", "edge_dist_in", "notes"]:
                        v = p.get(k)
                        if v not in (None, "") and pd.notna(v):
                            row[k] = v
                    m = map_material(p.get("material_grade"))
                    if m:
                        row["material_grade"] = m
                    b = map_bolt(p.get("fastener_type"))
                    if b:
                        row["fastener_type"] = b
                rows.append(row)
            working = ensure_connection_columns(pd.DataFrame(rows))

            # Show autopopulation status when Block 1 data has thickness values.
            if st.session_state.block1_raw is not None:
                filled = int((working["sheet_t_in"] > 0).sum())
                if filled > 0:
                    st.info(f"Thickness (sheet_t_in) autopopulated from Block 1 for {filled} of {len(working)} joints. Review before running.", icon="ℹ️")
                else:
                    st.caption("Thickness (sheet_t_in) could not be read from Block 1 — section wall_thickness may be zero or unclassified. Enter values manually.")

            with st.expander("How to fill this in", expanded=False):
                st.markdown(
                    "Each row is one bolted joint.\n\n"
                    "- **Connection material** — material of the governing (thinner / weaker) connected part. "
                    "Sets the ultimate strength Fu. Pick from the list.\n"
                    "- **Thickness (in)** — actual sheet/wall thickness of that governing part. Drives bearing and "
                    "tear-out. *Required* — a joint left at 0 is flagged for review, not silently failed.\n"
                    "- **# Fasteners** — number of bolts in the connection. Default 1.\n"
                    "- **Bolt diameter (in)** — nominal bolt diameter. Default 0.3125 (5/16 in).\n"
                    "- **Edge distance Le (in)** — center of the end bolt to the part edge along the load line. "
                    "Drives tear-out. Default 1.5*Bolt Diameter.\n"
                    "- **Bolt grade** — AISC strength group. Default ASTM A307, Grade A (most conservative). "
                    "Pick from the list.\n\n"
                    "Capacity is the factored group strength (φ = 0.75) per AISC Chapter J: the governing of bolt "
                    "shear vs. bearing. Screw and weld checks are no longer part of this step."
                )

            st.markdown("### Bolted connection schedule")
            display_cols = ["joint_id", "material_grade", "sheet_t_in", "n_fasteners",
                            "diameter_in", "edge_dist_in", "fastener_type", "notes"]
            edited = st.data_editor(
                working[display_cols],
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "joint_id": st.column_config.TextColumn("Joint ID", help="Must match a detected joint ID."),
                    "material_grade": st.column_config.SelectboxColumn("Connection material", options=MATERIALS, help="Governing connected part. Sets Fu."),
                    "sheet_t_in": st.column_config.NumberColumn("Thickness (in)", min_value=0.0, format="%.4f", help="Governing part thickness. Required."),
                    "n_fasteners": st.column_config.NumberColumn("# Fasteners", min_value=0, step=1),
                    "diameter_in": st.column_config.NumberColumn("Bolt dia. (in)", min_value=0.0, format="%.4f"),
                    "edge_dist_in": st.column_config.NumberColumn("Edge dist. Le (in)", min_value=0.0, format="%.3f"),
                    "fastener_type": st.column_config.SelectboxColumn("Bolt grade", options=BOLTS, help="AISC strength group."),
                    "notes": st.column_config.TextColumn("Notes"),
                },
            )

            # Persist: force the single bolted type and restore the full schema.
            edited = edited.copy()
            edited["connection_type"] = "bolted_bracket_joint"
            edited["template_name"] = "Bolted sheet-metal connection"
            st.session_state.connection_schedule = ensure_connection_columns(edited)

            if st.button("Reset to bolted defaults", use_container_width=True):
                st.session_state.connection_schedule = ensure_connection_columns(
                    pd.DataFrame([default_bolted_row(j) for j in ids]))
                st.rerun()
    # STEP 4
    elif step.startswith("4"):
        header("4. Loads and combinations", "Enter wind and seismic parameters to run the structural FE solve and compute joint forces.")

        st.markdown("### Structural FE analysis (Block 3 · PyNite)")
        if not BLOCK3_AVAILABLE:
            st.warning(f"PyNite engine unavailable: {BLOCK3_IMPORT_ERROR}. Install with `pip install PyNiteFEA` to enable the solve. You can still upload or edit a force table manually below.", icon="⚠️")
        elif st.session_state.block1_raw is None:
            st.info("Load Block 1 geometry in Step 2 first, then return here to run the solve.")
        else:
            from engine.loads.asce7_lrfd_loads import (
                PARAMETRIC_PRESETS, estimate_projected_area_ft2, parametric_to_point_loads,
            )
            b1 = st.session_state.block1_raw
            mass_unit = "lb"

            # ── Row 1: preset + fixity ────────────────────────────────────────
            r1a, r1b = st.columns([3, 1])
            with r1a:
                preset = st.selectbox("Qualification preset", list(PARAMETRIC_PRESETS.keys()),
                                      help="Equipment wind/seismic qualification level. Choosing a preset fills Wind and Seismic fields. Select 'Custom' to enter your own values.")
            with r1b:
                fixity = st.selectbox("Support fixity", ["fixed", "pinned"],
                                      help="The auto-detected supports are a collinear pair — 'pinned' alone is a mechanism in 3D. Use 'pinned' only with 3+ non-collinear supports.")

            pv = PARAMETRIC_PRESETS.get(preset) or {}
            sds = float(st.session_state.seismic_sds)

            # ── Row 2: Wind | Seismic inputs ─────────────────────────────────
            p1, p2 = st.columns(2)
            with p1:
                st.markdown("**Wind**")
                wind_psf = st.number_input("Design pressure q (psf)", min_value=0.0, step=5.0, value=float(pv.get("wind_psf", 30.0)))
                wind_dir = st.selectbox("Wind direction", ["X (+)", "Z (+)"], help="Lateral direction the wind acts on; the load combinations apply it in both ± directions automatically.")
                auto_area = estimate_projected_area_ft2(b1, wind_dir, "in")
                ref_area = st.number_input("Projected area A (ft²)", min_value=0.0, step=1.0, value=round(auto_area, 1), help=f"Silhouette area of the unit facing the wind. Auto-estimated from the Block 1 bounding box ≈ {auto_area:.1f} ft². Override with the real face area when known.")
            with p2:
                st.markdown("**Seismic**")
                sds = st.number_input("SDS (g)", min_value=0.0, step=0.05, value=float(pv.get("sds", sds)), help="Design short-period spectral acceleration from the site hazard. Drives vertical seismic Ev = 0.2·SDS·D in the combinations.")
                coeff = st.number_input("Horizontal coefficient (g)", min_value=0.0, step=0.05, value=float(pv.get("sds", sds)), help="Lateral seismic load coefficient. Defaults to SDS for a rigid mount at the top of the structure (x/h = 1.0). QE = coeff × operating weight.")
                seismic_dir = st.selectbox("Seismic direction", ["X (+)", "Z (+)"], help="Horizontal direction seismic QE acts on; combinations apply it reversibly.")

            # ── Row 3: secondary parameters ───────────────────────────────────
            r3a, r3b = st.columns(2)
            with r3a:
                rho = st.number_input("Redundancy factor ρ", min_value=1.0, max_value=1.3, step=0.1, value=1.3, help="ASCE 7-22 §12.3.4. Use ρ = 1.3 for most cooling tower frames. ρ = 1.0 only when a licensed engineer has confirmed sufficient redundancy.")
            with r3b:
                op_w = st.number_input("Operating weight Wp (lb)", min_value=0.0, step=10.0, value=0.0, help="Seismic weight of the equipment. Leave 0 to derive it from the model self-weight in Block 1.")

            wind_loads, seismic_loads, info = parametric_to_point_loads(
                b1, wind_psf=wind_psf, ref_area_ft2=ref_area, wind_dir=wind_dir,
                seismic_coeff_g=coeff, seismic_dir=seismic_dir, source_mass_unit=mass_unit,
                operating_weight_lb=(op_w or None))

            # ── Load summary cards ────────────────────────────────────────────
            st.markdown(metric_cards_html([
                ("Operating weight Wp", f"{info['operating_weight_lb']:,.0f} lb", "blue"),
                ("Total wind W", f"{info['wind_total_lb']:,.0f} lb", "teal"),
                ("Total seismic QE", f"{info['seismic_total_lb']:,.0f} lb", "yellow"),
            ]), unsafe_allow_html=True)
            st.caption(f"Loads distributed across {info['n_loaded_nodes']} joints proportionally to self-weight. LRFD combinations apply wind and seismic reversibly (± directions).")

            if st.button("Run structural FE solve", type="primary", use_container_width=True):
                res = run_lrfd_joint_analysis(
                    b1, S_DS=sds, rho=rho,
                    wind_loads=wind_loads, seismic_h_loads=seismic_loads,
                    support_fixity=fixity, source_mass_unit=mass_unit)
                if res.get("status") == "ok":
                    meta = res["meta"]
                    st.session_state.block3_result = res
                    st.session_state.combo_table = combos_table(res)
                    st.session_state.joint_forces = joint_forces_table(res)
                    st.session_state.balance_table = pd.DataFrame()
                    st.success(f"Solve complete. Supports {meta['support_nodes']} ({meta['support_fixity']}), {meta['n_elements']} elements, {len(meta['combos'])} combos. Joint forces and combinations updated.")
                    flagged = sorted({n['member'] for n in meta.get('build_notes', []) if n.get('needs_review')})
                    if flagged:
                        st.info("Sections using fallback / needs-review properties: " + ", ".join(flagged))
                else:
                    st.error(f"Solve failed: {res.get('error')}. {res.get('hint', '')}")

        st.divider()
        st.markdown("### Manual / demo force table (optional)")
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
        st.markdown("### Editable load combinations")
        st.session_state.combo_table = st.data_editor(st.session_state.combo_table, use_container_width=True, hide_index=True, num_rows="dynamic")
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

    # STEP 5
    elif step.startswith("5"):
        header("5. Validate before running", "This prevents crashes and tells the user what needs to be fixed in plain language.")
        checklist = draw_validation_summary()
        st.dataframe(checklist[["check", "status", "message", "recommended_action"]], use_container_width=True, hide_index=True)
        st.markdown("### Why validation matters")
        st.write("A design engineer should not have to debug Python errors. The app should identify missing geometry, incomplete connection rows, inconsistent units, or force rows that reference the wrong joint.")

    # STEP 6
    elif step.startswith("6"):
        header("6. Run and review results", "The first output is plain-language status. Technical tables are available underneath.")
        checklist = validation_checklist(st.session_state.joints, st.session_state.members, st.session_state.combo_table, st.session_state.connection_schedule, st.session_state.joint_forces, st.session_state.units)
        ready, message = is_ready_to_run(checklist)
        if not ready:
            st.error(message)
            st.dataframe(checklist[["check", "status", "message", "recommended_action"]], use_container_width=True, hide_index=True)
        run_anyway = st.checkbox("Run even with warnings/blockers", value=False, disabled=ready)
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

    # STEP 7
    elif step.startswith("7"):
        header("7. Export", "Create files that a design engineer can share with a reviewer or keep as a design record.")
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

