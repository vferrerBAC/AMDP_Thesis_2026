from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable

import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

from engine.connection_templates import (
    apply_template_to_rows,
    connection_keys,
    connection_schedule_template,
    default_connection_schedule,
    is_active_connection,
    template_names,
    templates_table,
)
from engine.connection_demand import (
    DEMAND_COLUMNS,
    connection_demands,
    demand_coverage,
)
from engine.joint_recommendation import (
    build_joint_verdicts,
    demand_from_analysis_results,
)
from engine.joining_cost import DEFAULT_PFD
from engine.joint_checks import (
    critical_joint_summary,
    default_joint_forces,
    ensure_connection_columns,
    ensure_force_columns,
    evaluate_connection_demands,
    joint_balance_table,
)
from engine.joint_io import (
    connected_members_summary,
    extract_joints_and_members,
    geometry_metrics,
)
from engine.load_combinations import default_load_combinations, load_combination_template
from engine.reporting import to_csv_bytes, to_excel_bytes, to_json_bytes, to_pdf_bytes
from engine.validation import is_ready_to_run, validation_checklist
from engine.bolted_capacity import (
    MATERIALS, BOLTS, default_bolted_row, map_material, map_bolt,
    governing_gauge, thickness_from_gauge,
)
from engine.cost_adapter import build_cost_inputs
from engine.cost_page import render_cost_run
from engine.manufacturability import (
    SUPPORTED_GAUGES,
    SUPPORTED_MATERIALS,
    build_manufacturability_inputs,
    evaluate_manufacturability_rows,
    summarize_results,
)

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
except Exception as _exc: # pragma: no cover
    BLOCK3_AVAILABLE = False
    BLOCK3_IMPORT_ERROR = str(_exc)

# Block 1 (live Inventor extraction) integration. Imported defensively: the
# module is platform-safe (win32 is guarded), so this import succeeds on any
# host. Only run()/extract_live_to_dict() require Windows + a live session.
try:
    from engine import block1 as block1_pipeline
    BLOCK1_AVAILABLE = True
    BLOCK1_IMPORT_ERROR = ""
except Exception as _exc: # pragma: no cover
    BLOCK1_AVAILABLE = False
    BLOCK1_IMPORT_ERROR = str(_exc)

st.set_page_config(
    page_title="BAC Integrated Joint Evaluation Tool (IJET)",
    page_icon=str(Path(__file__).parent / "IJET_logo.ico"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
/* ══ Quiet Field theme ══════════════════════════════════════════════════
   One accent (#2f5fd0), flat surfaces, cool-grey neutrals, no dark slab.
   Status is carried by color + text, not icons. */

/* ── Layout ── */
.block-container {padding-top: 1rem; max-width: 1180px;}

/* ── Sidebar · light ── */
[data-testid="stSidebar"] {background-color: #f6f8fb !important; border-right: 1px solid #e6ebf2;}
[data-testid="stSidebar"] > div:first-child {background-color: #f6f8fb !important;}
[data-testid="stSidebar"] h1 {font-size: 20px !important; font-weight: 700;
    letter-spacing: 0.02em; color: #16202f !important;}
[data-testid="stSidebar"] .stCaption p, [data-testid="stSidebar"] small {color: #8797ac !important;}
[data-testid="stSidebar"] hr {border-color: #e6ebf2 !important;}
[data-testid="stSidebar"] [data-testid="stExpander"] {
    background-color: #ffffff !important; border: 1px solid #e6ebf2 !important;}
[data-testid="stSidebar"] [data-testid="stRadio"] > div {gap: 2px;}
[data-testid="stSidebar"] [data-testid="stRadio"] label {
    font-size: 13.5px !important; padding: 7px 10px !important;
    border-radius: 7px !important; color: #47566b !important;}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {background: #eef1f6 !important;}

/* ── Section header · label over a hairline (replaces gradient banner) ── */
.sec-head {margin: 26px 0 16px;}
.sec-head .sec-label {font-size: 12px; font-weight: 700; letter-spacing: 0.11em;
    text-transform: uppercase; color: #16202f;}
.sec-head .sec-step {color: #8797ac; font-weight: 600;}
.sec-head .sec-cap {font-size: 13px; color: #6b7a90; margin-top: 4px;}
.sec-head .sec-rule {height: 1px; background: #e6ebf2; margin-top: 10px;}

/* ── Metric cards · flat with a thin top edge ── */
.mc-row {display: flex; gap: 12px; margin-bottom: 20px;}
.mc {
    flex: 1; background: #ffffff; border: 1px solid #e6ebf2;
    border-top: 3px solid #2f5fd0; border-radius: 10px;
    padding: 14px 16px; min-width: 0;}
.mc-label {
    font-size: 11px; font-weight: 700; color: #6b7a90;
    letter-spacing: 0.06em; text-transform: uppercase; margin-bottom: 8px;}
.mc-value {font-size: 26px; font-weight: 700; color: #16202f; line-height: 1;}
.mc-blue {border-top-color: #2f5fd0;}
.mc-green {border-top-color: #1f8a4c;}
.mc-teal {border-top-color: #0d9488;}
.mc-yellow {border-top-color: #b57608;}
.mc-red {border-top-color: #c5423a;}
.mc-gray {border-top-color: #b7c0cf;}

/* ── Tabs · quiet underline ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {gap: 22px;}
[data-testid="stTabs"] [data-baseweb="tab"] {padding: 6px 0;}
[data-testid="stTabs"] [aria-selected="true"] {color: #2f5fd0 !important;}

/* ── Source badge (sidebar footer) ── */
.src-badge {
    background: #ffffff; border: 1px solid #e6ebf2; border-radius: 8px;
    padding: 10px 14px; margin-top: 6px;}
.src-badge .src-label {
    font-size: 10px; color: #8797ac; text-transform: uppercase;
    letter-spacing: 0.05em; margin-bottom: 3px;}
.src-badge .src-value {font-size: 13px; font-weight: 600; color: #16202f;}

/* ── Status colours ── */
div[data-testid="stMetric"] {background: #ffffff; border: 1px solid #e6ebf2; padding: 12px 14px; border-radius: 10px;}
.small-card {background: #ffffff; padding: 14px 16px; border-radius: 10px; border: 1px solid #e6ebf2;}
.ok {color: #1f8a4c; font-weight: 700;}
.warn {color: #b57608; font-weight: 700;}
.fail {color: #c5423a; font-weight: 700;}
.gray {color: #667085; font-weight: 700;}
</style>
""",
    unsafe_allow_html=True,
)


# ── Performance: memoize pure engine computations ────────────────────────
# Streamlit re-runs this whole script — and every tab body — on each interaction.
# These engine calls are pure functions of their inputs, so cache them: unchanged
# geometry/tables return instantly instead of recomputing on every rerun. The
# imported names are rebound to their cached versions so every call site benefits.
_build_cost_inputs = build_cost_inputs
_build_joint_verdicts = build_joint_verdicts
_validation_checklist = validation_checklist
_geometry_metrics = geometry_metrics


@st.cache_data(show_spinner=False)
def build_cost_inputs(block1_data, aggregate_by_identifier: bool = True):
    return _build_cost_inputs(block1_data, aggregate_by_identifier)


@st.cache_data(show_spinner=False)
def build_joint_verdicts(block1_data, critical_summary=None, connection_schedule=None,
                         mfg_results=None, electrode="E70XX", pfd=DEFAULT_PFD):
    return _build_joint_verdicts(
        block1_data, critical_summary=critical_summary,
        connection_schedule=connection_schedule, mfg_results=mfg_results,
        electrode=electrode, pfd=pfd)


@st.cache_data(show_spinner=False)
def validation_checklist(joints, members, combos, connections, forces, units):
    return _validation_checklist(joints, members, combos, connections, forces, units)


@st.cache_data(show_spinner=False)
def geometry_metrics(joints, members):
    return _geometry_metrics(joints, members)


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
        "steel_code": "AISC 360 / AISI S100 / ASCE 7-22",
        "code_edition": "",
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
    st.session_state.setdefault("connection_schedule", default_connection_schedule([]))
    st.session_state.setdefault("joint_forces", default_joint_forces(["J001", "J002", "J003"], ["LRFD-01", "LRFD-02"]))
    st.session_state.setdefault("analysis_results", pd.DataFrame())
    st.session_state.setdefault("critical_summary", pd.DataFrame())
    st.session_state.setdefault("balance_table", pd.DataFrame())
    # Per-connection demand from the FE solve (engine/connection_demand.py) and the
    # flags explaining anything it could not resolve.
    st.session_state.setdefault("demand_table", pd.DataFrame(columns=DEMAND_COLUMNS))
    st.session_state.setdefault("demand_review", [])


init_state()


def joint_ids() -> list[str]:
    if st.session_state.joints.empty:
        return []
    return st.session_state.joints["joint_id"].astype(str).tolist()


def combo_ids() -> list[str]:
    if st.session_state.combo_table.empty or "combo_id" not in st.session_state.combo_table.columns:
        return []
    return st.session_state.combo_table["combo_id"].astype(str).tolist()


def editor_key(prefix: str, row_ids: Iterable[Any]) -> str:
    """A data_editor key that changes when the ROW SET changes.

    ``st.data_editor`` keeps its edit state keyed by row POSITION. With a constant
    key — or no key, where Streamlit derives one positionally — that state
    survives a rebuild of the underlying frame, so edits made against one set of
    rows replay onto whatever now sits at those indices. That silently overwrites
    real engineering inputs: it is what zeroed the gauge on scattered rows after
    the connection count changed from a face-contact extraction.

    Tying the key to the row identities drops the stale state exactly when it
    stops meaning anything. Edits are not lost by this: they round-trip through
    ``st.session_state.connection_schedule`` (restored via ``prev`` on the next
    run), not through widget state.
    """
    digest = hashlib.md5("\x00".join(str(i) for i in row_ids).encode()).hexdigest()
    return f"{prefix}_{digest[:12]}"


def _thickness_map_from_block1(block1_raw: dict, members_df: pd.DataFrame) -> dict:
    """Return {joint_id: min_wall_thickness_in} from Block 1 member cross-sections.

    Tries the real Block 1 schema first (joints[].member_names → member.cross_section
    .wall_thickness), then falls back to the simple schema where members carry
    start_joint / end_joint directly. Returns an empty dict when no thickness
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


def _gauge_map_from_block1(block1_raw: dict, members_df: pd.DataFrame) -> dict:
    """Return {joint_id: (gauge, thickness_in)} from the Inventor ``Gauge``
    iProperty read for each member (Block 1's cost extension: ``member.cost.Gauge``).

    Mirrors ``_thickness_map_from_block1``'s two schema paths, but the governing
    member is picked by SMALLEST resulting thickness (thinnest/weakest connected
    part), not by gauge number directly — gauge numbering runs opposite to
    thickness (8 ga = 0.153 in, 18 ga = 0.044 in). Thickness is translated from
    gauge via ``bolted_capacity.thickness_from_gauge``. Unrecognized/missing
    gauges are skipped; returns {} if no gauge iProperty was read anywhere.
    """
    try:
        result: dict = {}
        raw_members = block1_raw.get("members", [])

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
            gauges = []
            for mname in (j.get("member_names") or []):
                m = member_lookup.get(mname)
                if not isinstance(m, dict):
                    continue
                g = (m.get("cost") or {}).get("Gauge")
                if g is not None:
                    gauges.append(g)
            best = governing_gauge(gauges)
            if best is not None:
                result[jid] = best

        # Fallback: simple schema — members carry start_joint / end_joint directly.
        if not result and not members_df.empty:
            candidates: dict = {}
            for raw_m in raw_members:
                if not isinstance(raw_m, dict):
                    continue
                g = (raw_m.get("cost") or {}).get("Gauge")
                if g is None:
                    continue
                t = thickness_from_gauge(g)
                if t is None:
                    continue
                mid = str(raw_m.get("member_id", raw_m.get("occurrence_name", "")))
                matches = members_df[members_df["member_id"].astype(str) == mid]
                for _, mrow in matches.iterrows():
                    for jid in (str(mrow["start_joint"]), str(mrow["end_joint"])):
                        if jid and (jid not in candidates or t < candidates[jid][1]):
                            candidates[jid] = (g, t)
            result = candidates

        return result
    except Exception:
        return {}


def _member_lookup_from_block1(block1_raw: dict) -> dict:
    """{occurrence_path or occurrence_name: member dict}.

    Indexed by both names because a connection's ``member_a``/``member_b`` may
    carry either, depending on assembly nesting depth.
    """
    lookup: dict = {}
    for m in (block1_raw or {}).get("members", []) or []:
        if not isinstance(m, dict):
            continue
        path = m.get("occurrence_path", "")
        name = m.get("occurrence_name", "")
        if path:
            lookup[path] = m
        if name and name not in lookup:
            lookup[name] = m
    return lookup


def _conn_gauge_thickness_maps(block1_raw: dict) -> tuple[dict, dict]:
    """Per-CONNECTION {connection_id: (gauge, thickness)} and {connection_id: t}.

    Read straight off the two parts the patch actually joins (``member_a`` /
    ``member_b``), governed by the smallest resulting thickness. This is more
    specific than the joint-wide maps — two patches at one joint can bear on
    different parts — and it is the only source that works for connections
    created after upload: the Step 3 face pass renumbers and expands connections,
    so their ids did not exist when the joint-keyed pass ran at ingest.
    """
    gauge_by_cid: dict = {}
    t_by_cid: dict = {}
    try:
        lookup = _member_lookup_from_block1(block1_raw)
        if not lookup:
            return {}, {}
        for c in (block1_raw or {}).get("connections", []) or []:
            if not is_active_connection(c):
                continue
            cid = str(c.get("connection_id") or "").strip()
            if not cid:
                continue
            parts = [lookup.get(str(c.get(k) or "")) for k in ("member_a", "member_b")]
            parts = [m for m in parts if isinstance(m, dict)]
            if not parts:
                continue

            gauges = [g for g in ((m.get("cost") or {}).get("Gauge") for m in parts)
                      if g is not None]
            best = governing_gauge(gauges)
            if best is not None:
                gauge_by_cid[cid] = best

            thicknesses = []
            for m in parts:
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
                t_by_cid[cid] = min(thicknesses)
    except Exception:
        return gauge_by_cid, t_by_cid
    return gauge_by_cid, t_by_cid


def autofill_gauge_thickness(sched: pd.DataFrame, block1_raw: dict,
                             members_df: pd.DataFrame) -> pd.DataFrame:
    """Fill gauge / sheet_t_in on schedule rows that still carry no value.

    Resolution order per row: the connection's own two parts, then the joint-wide
    gauge map, then the joint-wide thickness map. A row is only touched when BOTH
    gauge and thickness are still 0 — once either holds a value (autofilled
    earlier or typed by the engineer) the row is left alone, so this can run on
    every rerun without overwriting edits.
    """
    if sched is None or sched.empty:
        return sched
    gauge_by_cid, t_by_cid = _conn_gauge_thickness_maps(block1_raw)
    gauge_map = _gauge_map_from_block1(block1_raw, members_df)
    t_map = _thickness_map_from_block1(block1_raw, members_df)
    if not (gauge_by_cid or t_by_cid or gauge_map or t_map):
        return sched

    out = sched.copy()
    for i, row in out.iterrows():
        try:
            if float(row.get("gauge") or 0) > 0 or float(row.get("sheet_t_in") or 0) > 0:
                continue
        except (TypeError, ValueError):
            continue
        cid = str(row.get("connection_id", ""))
        jid = str(row.get("joint_id", ""))
        if cid in gauge_by_cid:
            gauge, thickness = gauge_by_cid[cid]
            out.at[i, "gauge"] = gauge
            out.at[i, "sheet_t_in"] = thickness
        elif cid in t_by_cid:
            out.at[i, "sheet_t_in"] = t_by_cid[cid]
        elif jid in gauge_map:
            gauge, thickness = gauge_map[jid]
            out.at[i, "gauge"] = gauge
            out.at[i, "sheet_t_in"] = thickness
        elif jid in t_map:
            out.at[i, "sheet_t_in"] = t_map[jid]
    return out


def ingest_block1_dict(data: Dict[str, Any]) -> tuple[int, int]:
    """Single ingest path shared by upload, sample, and live-Inventor extraction.

    Keeps all three sources byte-identical downstream: parse joints/members,
    recover connectivity from the real Block-1 schema when the generic parser
    finds no members, then seed the connection schedule and force table.
    Returns (n_joints, n_members).
    """
    joints, members = extract_joints_and_members(data)

    # Flag centerline pairs whose members are farther apart than they could ever
    # touch, so the phantom near-pairs the generous centerline reach produces at
    # multi-member joints stay out of the schedule, counts, and recommendation.
    # Marked, not deleted: the heuristic assumes each solid stays inside its
    # nominal section box, which a gusset or bracket breaks, so the pair has to
    # remain inspectable in Step 3 and reachable by the face pass.
    try:
        conns, pruned_ids, prune_diag = block1_pipeline.prune_noncontact_connections(data)
        data["connections"] = conns
        data["connection_prune"] = prune_diag
    except Exception:
        pass

    st.session_state.block1_raw = data
    if members.empty and BLOCK3_AVAILABLE:
        members = members_dataframe_from_block1(data)
    st.session_state.joints = joints
    st.session_state.members = members
    ids = joint_ids()
    # Seed one row per CONNECTION, with bolted defaults (not the screw template)
    # so the bolted-connections step shows the right fastener/diameter/edge-distance
    # starting values. A joint with three contact patches gets three rows: they are
    # three separate connections, each with its own thickness, bolt count, and
    # material, and each has to be checked on its own.
    keys = connection_keys(data)
    if not keys:
        # No connections extracted yet (centerline pass not run, or legacy JSON).
        # Fall back to one row per joint, marked so the origin is not mistaken for
        # a real contact patch.
        keys = [(f"{j}:1", j) for j in ids]
    sched = ensure_connection_columns(
        pd.DataFrame([default_bolted_row(cid, jid) for cid, jid in keys]))
    if not connection_keys(data) and ids:
        sched["notes"] = "Placeholder: no extracted connection. Run Step 3."

    # Autopopulate gauge + thickness from Block 1. Gauge (Inventor iProperty,
    # translated to thickness via the standard gauge chart) takes priority over
    # the directly-measured wall thickness; where the connected members carry
    # different gauges, the smallest resulting thickness governs. Step 4 runs the
    # same pass again, so connections that only appear later (the face pass in
    # Step 3 renumbers and expands them) get filled too.
    sched = autofill_gauge_thickness(sched, data, members)

    st.session_state.connection_schedule = sched
    st.session_state.demand_table = pd.DataFrame(columns=DEMAND_COLUMNS)
    st.session_state.demand_review = []

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
    step = f'<span class="sec-step">Stage {num} · </span>' if num else ""
    st.markdown(
        f'<div class="sec-head"><div class="sec-label">{step}{name}</div>'
        f'<div class="sec-cap">{caption}</div><div class="sec-rule"></div></div>',
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


def connection_points(block1_raw: Any) -> pd.DataFrame:
    """Extract the connection contacts (yellow dots) from a Block 1 result dict.

    Reads the ``connections`` list emitted by the connection-aware Block 1
    (each = a member/face-pair contact with a 3D ``location``). Returns an empty
    frame for legacy/sample JSON that predates connections, so the preview
    simply shows no dots in that case.
    """
    cols = ["connection_id", "member_a", "member_b", "joint_id",
            "connection_type", "detection_method", "x", "y", "z"]
    if not isinstance(block1_raw, dict):
        return pd.DataFrame(columns=cols)
    rows = []
    for c in block1_raw.get("connections", []) or []:
        if not is_active_connection(c):
            continue
        loc = c.get("location") or []
        try:
            x, y, z = float(loc[0]), float(loc[1]), float(loc[2])
        except (TypeError, ValueError, IndexError):
            continue
        rows.append({
            "connection_id": c.get("connection_id", ""),
            "member_a": c.get("member_a", ""),
            "member_b": c.get("member_b", ""),
            "joint_id": c.get("joint_id", ""),
            "connection_type": c.get("connection_type", "unknown"),
            "detection_method": c.get("detection_method", ""),
            "x": x, "y": y, "z": z,
        })
    return pd.DataFrame(rows, columns=cols)


_WHY_TEXT = {
    "near_contact_not_face_confirmed":
        "NOT face-confirmed — the solid model found no touching faces, only "
        "parallel faces close enough to be plausible. Kept so a real load path is "
        "not erased. No contact patch, so it has no area/thickness data. "
        "Uncheck if these parts do not actually fasten.",
    "centerline_only_faces_unreadable":
        "Faces could not be read (no live Inventor session or no solid body). "
        "Centerline guess only — re-run face detection with the model open.",
}


def _connection_why(c: dict) -> str:
    """One line explaining why a connection is in the list, for the review table.

    The three classes the face pass emits are not otherwise visible, which is what
    makes an unconfirmed near-contact look like a mystery extra row rather than a
    deliberate 'you decide' flag.
    """
    if not isinstance(c, dict):
        return ""
    if c.get("pruned_noncontact") and not c.get("prune_override"):
        return "Unreachable — " + str(c.get("prune_reason") or "")
    if c.get("prune_override") and c.get("detection_method") != "face_contact":
        return ("Restored by hand over the geometric reach test. Run face detection "
                "to confirm it against real geometry.")
    reasons = [r.strip() for r in str(c.get("review_reason") or "").split(";") if r.strip()]
    for r in reasons:
        if r in _WHY_TEXT:
            return _WHY_TEXT[r]
    if c.get("detection_method") == "face_contact":
        return "Confirmed face contact — real touching faces measured on the solid model."
    if c.get("is_inferred"):
        return ("Centerline pair, near but not touching within tolerance. Run face "
                "detection to confirm or drop it.")
    if reasons:
        return "" + "; ".join(reasons)
    return "Centerline pair — run face detection to confirm."


def connection_patch_rings(block1_raw: Any) -> list:
    """Reconstruct the real 3D contact-patch outlines from face-contact connections.

    After Step 3 (face detection), each ``face_contact`` connection carries the
    true contact patch as a 2D outline (+ bolt holes) in its own ``patch_frame``
    (origin, x_axis, y_axis, all in inches). This maps that outline back to 3D so
    the preview can draw the actual bolt/weld footprint instead of only a dot.
    Centerline connections have no patch and are skipped.
    """
    rings: list = []
    if not isinstance(block1_raw, dict):
        return rings
    for c in block1_raw.get("connections", []) or []:
        if not isinstance(c, dict) or c.get("detection_method") != "face_contact":
            continue
        if not is_active_connection(c):
            continue
        frame = c.get("patch_frame") or {}
        origin, xa, ya = frame.get("origin"), frame.get("x_axis"), frame.get("y_axis")
        ext = c.get("patch_exterior_2d") or []
        if not (origin and xa and ya) or len(ext) < 3:
            continue

        def _to3d(u, v):
            return (origin[0] + u * xa[0] + v * ya[0],
                    origin[1] + u * xa[1] + v * ya[1],
                    origin[2] + u * xa[2] + v * ya[2])

        def _ring(pts2d):
            r = [_to3d(float(u), float(v)) for u, v in pts2d]
            if r:
                r.append(r[0]) # close the loop
            return r

        rings.append({
            "type": c.get("connection_type", "unknown"),
            "exterior": _ring(ext),
            "holes": [_ring(h) for h in (c.get("patch_holes_2d") or []) if len(h) >= 3],
            "label": f"{c.get('connection_id', '')} · {c.get('connection_type', '')}",
        })
    return rings


def ensure_connections_pruned() -> dict | None:
    """Flag geometrically-unreachable connections on the loaded model if it has
    not been done yet. Idempotent (keyed on the ``connection_prune`` marker), so
    it also marks a model loaded before this existed — no re-upload needed.
    Nothing is deleted: the flag hides a pair from the schedule while leaving it
    visible and restorable in Step 3. Returns the (updated) block1_raw dict."""
    raw = st.session_state.block1_raw
    if isinstance(raw, dict) and raw.get("connections") and "connection_prune" not in raw:
        try:
            conns, pruned_ids, diag = block1_pipeline.prune_noncontact_connections(raw)
            raw["connections"] = conns
            raw["connection_prune"] = diag
        except Exception:
            pass
    return raw


def geometry_figure(joints_df: pd.DataFrame, members_df: pd.DataFrame,
                    connections_df: pd.DataFrame | None = None,
                    patch_rings: list | None = None):
    """Plotly 3D scatter with joint labels, member lines, connection markers, and
    (after face detection) the real contact-patch outlines."""
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
    # axis. "up" is determined by convention, not by max span — the max-span
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

    # Connection markers, colored by type (bolted/welded/unknown) and split by
    # detection method: face-contact = solid diamond (a confirmed, measured
    # contact); centerline = small faded circle (an approximate, unconfirmed
    # near-pair). The visual split makes the phantom near-pairs obvious before
    # face detection, and shows them collapse to real contacts afterwards.
    type_color = {"bolted": "#facc15", "welded": "#f97316", "unknown": "#9ca3af"}

    def _conn_hover(sub):
        return [
            f"{r.connection_id} · {r.connection_type}"
            f"<br>{r.member_a} ↔ {r.member_b}"
            f"<br>joint {r.joint_id} · {r.detection_method}"
            for r in sub.itertuples(index=False)
        ]

    if connections_df is not None and not connections_df.empty:
        face = connections_df[connections_df["detection_method"] == "face_contact"]
        center = connections_df[connections_df["detection_method"] != "face_contact"]

        if not center.empty:
            fig.add_trace(go.Scatter3d(
                x=center["x"].tolist(), y=center["y"].tolist(), z=center["z"].tolist(),
                mode="markers",
                marker=dict(size=4, symbol="circle", opacity=0.45,
                            color=[type_color.get(str(t), "#9ca3af") for t in center["connection_type"]],
                            line=dict(color="#6b7280", width=0.5)),
                name="Connections (centerline · approx)",
                text=_conn_hover(center), hovertemplate="%{text}<extra></extra>",
            ))
        if not face.empty:
            fig.add_trace(go.Scatter3d(
                x=face["x"].tolist(), y=face["y"].tolist(), z=face["z"].tolist(),
                mode="markers",
                marker=dict(size=6, symbol="diamond",
                            color=[type_color.get(str(t), "#9ca3af") for t in face["connection_type"]],
                            line=dict(color="#713f12", width=1)),
                name="Connections (face contact)",
                text=_conn_hover(face), hovertemplate="%{text}<extra></extra>",
            ))

    # Real contact-patch outlines (face-contact only): the actual bolt/weld
    # footprint, with bolt holes drawn as dotted inner rings.
    if patch_rings:
        patch_color = {"bolted": "#eab308", "welded": "#ea580c", "unknown": "#6b7280"}
        legended: set = set()
        for ring in patch_rings:
            ext = ring.get("exterior") or []
            if len(ext) < 3:
                continue
            col = patch_color.get(ring.get("type"), "#6b7280")
            show = ring.get("type") not in legended
            legended.add(ring.get("type"))
            fig.add_trace(go.Scatter3d(
                x=[p[0] for p in ext], y=[p[1] for p in ext], z=[p[2] for p in ext],
                mode="lines", line=dict(color=col, width=5),
                name=f"Contact patch · {ring.get('type', 'unknown')}",
                legendgroup=f"patch-{ring.get('type')}", showlegend=show,
                text=ring.get("label", ""), hovertemplate="%{text}<extra></extra>",
            ))
            for h in ring.get("holes") or []:
                if len(h) < 3:
                    continue
                fig.add_trace(go.Scatter3d(
                    x=[p[0] for p in h], y=[p[1] for p in h], z=[p[2] for p in h],
                    mode="lines", line=dict(color=col, width=2, dash="dot"),
                    legendgroup=f"patch-{ring.get('type')}", showlegend=False,
                    hoverinfo="skip",
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
        "Passed": "Passed",
        "Warning": "Warning",
        "Blocked": "Blocked",
        "OK": "OK",
        "WarningResult": "Warning",
        "Not OK": "Not OK",
        "Incomplete": "Incomplete",
        "Ignored": "Ignored",
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
        st.success(message)
    else:
        st.error(message)
    return checklist


# Sidebar
st.sidebar.title("IJET")
st.sidebar.caption("Design-engineer-friendly joint screening")
step = st.sidebar.radio(
    "Workflow",
    [
        "1 · Project",
        "2 · Geometry",
        "3 · Extract connections",
        "4 · Bolted connections",
        "5 · Loads",
        "6 · Validate",
        "7 · Results",
        "8 · Export",
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

tab1, tab2, tab3, tab4 = st.tabs(["Structural Analysis", "Financial Analysis", "Manufacturability Analysis", "Summary"])


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
            st.text_input("Primary code/check family", value="AISC 360 / AISI S100 / ASCE 7-22", disabled=True, help="Fixed by project scope. Steel members per AISC 360 (hollow/HSS) and AISI S100 (open/cold-formed); loads and combinations per ASCE 7-22.")
            st.text_input("Notes", key="code_edition")
        st.warning("This tool is meant for preliminary structural screening only, designs should still be approved by a structural engineer.")

    # STEP 2
    elif step.startswith("2"):
        header("2. Get geometry", "Extract data directly from an active Inventor session, please only have one assembly (.iam) open to avoid runtime delays.")

        # ── Metric cards ──────────────────────────────────────────────────────
        _nj = len(st.session_state.joints)
        _nm = len(st.session_state.members)
        _raw = st.session_state.block1_raw or {}
        _meta = (st.session_state.block3_result or {}).get("meta", {})
        _n_flags = len([n for n in _meta.get("build_notes", []) if n.get("needs_review")])
        _n_conn = sum(1 for c in (_raw.get("connections") or [])
                      if is_active_connection(c))
        st.markdown(metric_cards_html([
            ("Nodes", _nj, "blue"),
            ("Elements", _nm, "green"),
            ("Connections", _n_conn, "yellow"),
            ("Flags", _n_flags, "yellow"),
        ]), unsafe_allow_html=True)

        status = block1_pipeline.live_inventor_status() if BLOCK1_AVAILABLE else {"available": False, "reason": BLOCK1_IMPORT_ERROR}
        st.markdown("### Live extraction from Inventor")
        if status["available"]:
            st.success(f"Inventor detected. {status['reason']}")
        else:
            st.info(f"Live extraction unavailable: {status['reason']} Use the JSON upload below instead.")
        if st.button("Extract from live Inventor", type="primary", width="stretch", disabled=not status["available"]):
            try:
                with st.spinner("Running IJET against the active assembly…"):
                    # Fast path: geometry + centerline connections only. The
                    # expensive face-contact detection is a separate on-demand
                    # step (3 · Extract connections) so this stays quick.
                    data = block1_pipeline.extract_live_to_dict(with_face_connections=False)
                nj, nm = ingest_block1_dict(data)
                u = data.get("units", {})
                st.success(f"Extracted {nj} joints and {nm} members from {Path(str(data.get('source_document',''))).name or 'the active assembly'}.")
                st.caption(f"Canonical units: length={u.get('length','in')}, mass={u.get('mass','lbm')}, force={u.get('force','lbf')}.")
                st.info("Connections are centerline approximations for now. Run **3 · Extract connections** to resolve real face contacts (bolt/weld areas).")
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
            ]), width="stretch", hide_index=True)

        st.divider()
        st.markdown("### 3D geometry model")
        if st.session_state.joints.empty:
            st.info("No geometry loaded yet.")
        elif PLOTLY_AVAILABLE:
            # Geometry-only model: frame members + joint markers, NO connection
            # dots. Connections have their own dedicated 3D model in Step 3, so
            # this structural view stays clean and uncluttered at the corners.
            ensure_connections_pruned()
            fig = geometry_figure(st.session_state.joints, st.session_state.members)
            st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
            _nconn = sum(1 for c in ((st.session_state.block1_raw or {}).get("connections") or [])
                         if is_active_connection(c))
            _prune = (st.session_state.block1_raw or {}).get("connection_prune") or {}
            if _nconn:
                _extra = (f" ({_prune['connections_pruned']} impossible near-pair(s) pruned)"
                          if _prune.get("connections_pruned") else "")
                st.caption(
                    f"Frame only. The {_nconn} connection(s){_extra} are drawn on their own "
                    "3D model in **3 · Extract connections**."
                )
        else:
            plot_df = st.session_state.joints.rename(columns={"x": "X", "y": "Y"})
            st.scatter_chart(plot_df, x="X", y="Y", size=80, color=None)
        st.markdown("### Detected joints and connected members")
        summary = connected_members_summary(st.session_state.joints, st.session_state.members)
        st.dataframe(summary, width="stretch", hide_index=True)
        with st.expander("Raw joint coordinate table"):
            st.dataframe(st.session_state.joints, width="stretch", hide_index=True)
        with st.expander("Raw member connectivity table"):
            st.dataframe(st.session_state.members, width="stretch", hide_index=True)

    # STEP 3 — Extract connections (on-demand face-contact detection)
    elif step.startswith("3"):
        header("3. Extract connections",
               "Face-contact detection runs here as its own step, separate from geometry, so "
               "Step 2 stays fast. Run it once the model is loaded to resolve real contact "
               "patches (bolt/weld areas). Requires a live Inventor session; the centerline "
               "connections from Step 2 remain the fallback.")

        if st.session_state.block1_raw is None:
            st.warning("Load geometry in Step 2 first, then return here to extract connections.")
        else:
            raw = ensure_connections_pruned()
            all_conns = raw.get("connections") or []
            # "conns" is the active set — the same definition the Step 4 schedule
            # and the demand table use, so the count on this card is the number of
            # rows you get downstream. Nothing else in this step may redefine it.
            conns = [c for c in all_conns if is_active_connection(c)]
            n_face = sum(1 for c in conns if c.get("detection_method") == "face_contact")
            n_center = len(conns) - n_face
            n_removed = sum(1 for c in all_conns if c.get("manually_removed"))
            n_unreach = sum(1 for c in all_conns
                            if c.get("pruned_noncontact") and not c.get("prune_override"))

            st.markdown(metric_cards_html([
                ("Connections", len(conns), "yellow"),
                ("Face-resolved", n_face, "green" if n_face else "gray"),
                ("Centerline only", n_center, "teal"),
                ("Unreachable", n_unreach, "gray"),
                ("Removed by hand", n_removed, "gray"),
            ]), unsafe_allow_html=True)

            # After a face pass, "centerline only" means the solid model could NOT
            # confirm the contact. Those are kept deliberately (never silently
            # dropped) but carry no patch, so they are exactly the rows that look
            # empty downstream. Call them out rather than leaving the engineer to
            # infer it from a Method column.
            unconfirmed = [c for c in conns
                           if c.get("detection_method") != "face_contact"]
            if n_face and unconfirmed:
                st.warning(
                    f"**{len(unconfirmed)} of {len(conns)} connection(s) are NOT "
                    f"face-confirmed** — {', '.join(str(c.get('connection_id')) for c in unconfirmed[:8])}"
                    f"{' …' if len(unconfirmed) > 8 else ''}. Face detection found no touching "
                    "faces for these member pairs, only faces close enough to be plausible, so "
                    "they are kept for you to judge instead of being deleted. They have no "
                    "contact patch — which is why they appear with no data in the Step 4 "
                    "schedule. Review them below and uncheck any that do not really fasten.",
                )

            # Unreachable pairs: listed with the measurement that condemned them,
            # and overridable. The reach test only knows each member's nominal
            # depth x width box, so a connection made through a gusset, bracket or
            # tab looks impossible to it — which is why this is an argued case the
            # engineer can reject, not a silent delete.
            flagged = [c for c in all_conns
                       if c.get("pruned_noncontact") and not c.get("prune_override")]
            if flagged:
                with st.expander(
                    f"{len(flagged)} pair(s) flagged as unreachable — excluded from the "
                    "schedule (click to review or restore)", expanded=False):
                    st.caption(
                        "Each member can only reach half its section diagonal from its own "
                        "centerline, so beyond the combined reach no orientation brings the two "
                        "into contact. That assumes the solid stays inside its nominal section — "
                        "**a gusset plate, bracket, tab or bent flange reaches further.** If one "
                        "of these is a real bolted connection, restore it: it will be included "
                        "in the schedule and offered to face detection."
                    )
                    st.dataframe(pd.DataFrame([{
                        "ID": c.get("connection_id", ""),
                        "Joint": c.get("joint_id", ""),
                        "Member A": c.get("member_a", ""),
                        "Member B": c.get("member_b", ""),
                        "Short by (in)": c.get("prune_shortfall_in", ""),
                        "Why": c.get("prune_reason", ""),
                    } for c in flagged]), width="stretch", hide_index=True)
                    _restore = st.multiselect(
                        "Restore as real connections",
                        [str(c.get("connection_id")) for c in flagged],
                        key="prune_restore",
                        help="Overrides the geometric test for these pairs. Face detection can "
                             "still overrule the decision — it measures real geometry.")
                    if st.button("Restore selected", width="stretch", disabled=not _restore):
                        for c in all_conns:
                            if str(c.get("connection_id")) in set(_restore):
                                c["prune_override"] = True
                                c.pop("pruned_noncontact", None)
                                c.pop("prune_reason", None)
                                c.pop("prune_shortfall_in", None)
                        st.rerun()
                    st.caption("Run face detection below to settle these against real geometry.")

            _overridden = [c for c in all_conns if c.get("prune_override")]
            if _overridden:
                st.caption(
                    f"{len(_overridden)} pair(s) restored by hand over the geometric test: "
                    f"{', '.join(str(c.get('connection_id')) for c in _overridden[:8])}"
                    f"{' …' if len(_overridden) > 8 else ''}."
                )

            # Result of the most recent face pass, persisted so it survives the
            # rerun that refreshes the metric cards above.
            _summary = st.session_state.pop("_conn_extract_summary", None)
            if _summary:
                if _summary.get("unmatched"):
                    st.warning(
                        f"Member-identification mismatch: {_summary['unmatched']} member(s) from "
                        f"Step 2 were not found in the live model during re-link "
                        f"({_summary['matched']}/{_summary['prior']} matched). The assembly likely "
                        "changed since geometry extraction — re-extract in Step 2 before trusting "
                        "these connections.",
                    )
                else:
                    st.success(
                        f"Resolved {_summary['n_face']} face-contact connection(s) in "
                        f"{_summary['elapsed']:.1f}s. Member identification reconciled "
                        f"({_summary['matched']}/{_summary['prior']} re-linked exactly; "
                        "cross-sections untouched).",
                    )

            status = block1_pipeline.live_inventor_status() if BLOCK1_AVAILABLE else {"available": False, "reason": BLOCK1_IMPORT_ERROR}
            if status["available"]:
                st.success(f"Inventor detected. {status['reason']}")
            else:
                st.info(
                    f"Live extraction unavailable: {status['reason']} Face-contact detection "
                    "needs the live model; the centerline connections from Step 2 remain available.",
                )

            if n_face == 0 and conns:
                st.caption(
                    "Connections are currently centerline approximations. The centerline pass "
                    "over-counts where several members meet (it counts every close member pair, "
                    "not just touching faces). Resolve real face contacts below."
                )

            if st.button("Extract connections (face contact) from live Inventor",
                         type="primary", width="stretch", disabled=not status["available"]):
                try:
                    with st.spinner("Detecting face contacts against the active assembly…"):
                        upd = block1_pipeline.enrich_connections_live(raw)
                    new_raw = dict(raw)
                    new_raw["connections"] = upd["connections"]
                    new_raw["connection_diagnostics"] = upd["connection_diagnostics"]
                    timings = dict(new_raw.get("timings") or {})
                    timings["connections_face_s"] = upd["connections_face_s"]
                    new_raw["timings"] = timings
                    st.session_state.block1_raw = new_raw
                    nf = sum(1 for c in upd["connections"] if c.get("detection_method") == "face_contact")
                    _cd = upd["connection_diagnostics"] or {}
                    st.session_state["_conn_extract_summary"] = {
                        "n_face": nf,
                        "elapsed": upd["connections_face_s"],
                        "prior": _cd.get("members_prior", 0),
                        "matched": _cd.get("members_matched", 0),
                        "unmatched": _cd.get("members_unmatched", 0),
                    }
                    st.rerun()
                except Exception as exc:
                    st.error(f"Connection extraction failed: {exc}")

            # ── Preview + table ────────────────────────────────────────────────
            conns_df = connection_points(st.session_state.block1_raw)
            patch_rings = connection_patch_rings(st.session_state.block1_raw)
            if not st.session_state.joints.empty and PLOTLY_AVAILABLE:
                st.markdown("### 3D connection model")
                # Once face detection has run, draw with the purpose-built connection
                # viewer (engine.connection_viz) — the same one that produces the
                # standalone connections.html and matches the JointLocatorV16 macro:
                # true contact patches (teal, bolt holes in red) with the V16 OBB
                # ring (amber dashed) drawn on top so the box-vs-true-patch gap is
                # visible. Before face detection there are no patches, so fall back
                # to the simple marker view.
                active_conns = [c for c in all_conns if is_active_connection(c)]
                has_face = any(c.get("detection_method") == "face_contact" for c in active_conns)
                if has_face:
                    cc1, cc2 = st.columns(2)
                    show_obb = cc1.checkbox(
                        "Show OBB ring (amber)", value=True,
                        help="The bounding-box contact the VB macro measures, drawn over the true "
                             "patch. The gap between them is the area the box method overstates.")
                    show_spheres = cc2.checkbox(
                        "Show joint cluster spheres", value=False,
                        help="The radius that decides which contacts merge into one joint.")
                    from engine.connection_viz import figure_from_result
                    active_data = {**st.session_state.block1_raw, "connections": active_conns}
                    fig = figure_from_result(
                        active_data, show_obb=show_obb, show_spheres=show_spheres,
                        show_patches=True, show_centroids=True, show_members=True)
                    st.plotly_chart(fig, width="stretch", config={"displayModeBar": True})
                    n_f = sum(1 for c in active_conns if c.get("detection_method") == "face_contact")
                    st.caption(
                        f"{n_f} true contact patch(es) (teal, bolt holes in red) with the "
                        "OBB ring (amber dashed) on top — same representation as "
                        "the VB macro / connections.html viewer."
                    )
                else:
                    st.caption("Centerline approximations shown as markers. Run face detection above "
                               "to draw the true contact patches (VB-macro representation).")
                    fig = geometry_figure(st.session_state.joints, st.session_state.members,
                                          conns_df, patch_rings)
                    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

            # ── Manual review / prune ─────────────────────────────────────────
            # Safe workaround for phantoms geometry can't rule out (e.g. a brace
            # shown against a member it doesn't actually bolt to): the engineer
            # unchecks them. Non-destructive — a "manually_removed" flag is set,
            # so unchecking is fully reversible and never deletes source data.
            st.markdown("### Detected connections — review & prune")
            st.caption(
                "Uncheck any connection that is not a real contact (you can see them in the model "
                "above and know the design), then **Apply**. This is remembered and reversible — "
                "re-check to bring one back. Use it for phantoms that geometry alone can't rule out "
                "(a brace paired with a member it doesn't bolt to)."
            )
            # Pairs flagged unreachable are handled in their own panel above, with
            # the measurement behind the call — listing them here as unchecked
            # boxes would lose that argument and imply the engineer made the call.
            review_conns = [c for c in all_conns
                            if not (c.get("pruned_noncontact") and not c.get("prune_override"))]
            editor_rows = [{
                "keep": not c.get("manually_removed", False),
                "connection_id": c.get("connection_id", ""),
                "joint_id": c.get("joint_id", ""),
                "member_a": c.get("member_a", ""),
                "member_b": c.get("member_b", ""),
                "connection_type": c.get("connection_type", ""),
                "detection_method": c.get("detection_method", ""),
                "why": _connection_why(c),
            } for c in review_conns]
            if editor_rows:
                # Colour-code by detection method: green = real face contact
                # (confirmed), yellow = centerline pair kept for the engineer to
                # judge (not face-confirmed). Light tints of the metric-card
                # green (#1f8a4c) / yellow (#b57608) accents.
                _conn_df = pd.DataFrame(editor_rows)

                def _shade_connection(row):
                    fill = "#e7f4ec" if row["detection_method"] == "face_contact" else "#fbf3e0"
                    return [f"background-color: {fill}; color: #16202f"] * len(row)

                edited_conns = st.data_editor(
                    _conn_df.style.apply(_shade_connection, axis=1),
                    width="stretch", hide_index=True,
                    # Identity-keyed, not a constant: a constant key let stale
                    # "keep" unchecks replay onto whatever row later occupied that
                    # position once the connection list changed length.
                    key=editor_key("conn_review",
                                   [r["connection_id"] for r in editor_rows]),
                    disabled=["connection_id", "joint_id", "member_a", "member_b",
                              "connection_type", "detection_method", "why"],
                    column_config={
                        "keep": st.column_config.CheckboxColumn(
                            "Keep", help="Uncheck to remove a connection that is not a real contact."),
                        "connection_id": st.column_config.TextColumn("ID"),
                        "joint_id": st.column_config.TextColumn("Joint"),
                        "member_a": st.column_config.TextColumn("Member A"),
                        "member_b": st.column_config.TextColumn("Member B"),
                        "connection_type": st.column_config.TextColumn("Type"),
                        "detection_method": st.column_config.TextColumn("Method"),
                        "why": st.column_config.TextColumn(
                            "Why it's here", width="large",
                            help="Confirmed contact vs. an unconfirmed pair kept for you to "
                                 "judge. Unconfirmed rows carry no patch geometry — they are "
                                 "the ones that show up blank in the Step 4 schedule."),
                    },
                )
                b1, b2 = st.columns([1, 1])
                if b1.button("Apply connection edits", type="primary", width="stretch"):
                    remove_ids = set(edited_conns.loc[~edited_conns["keep"], "connection_id"])
                    for c in all_conns:
                        c["manually_removed"] = c.get("connection_id") in remove_ids
                    st.rerun()
                if n_removed and b2.button("Restore all removed", width="stretch"):
                    for c in all_conns:
                        c.pop("manually_removed", None)
                    st.rerun()

            _diag = (st.session_state.block1_raw or {}).get("connection_diagnostics") or {}
            if _diag:
                with st.expander("Connection detection diagnostics", expanded=(n_face == 0)):
                    if _diag.get("fatal_error"):
                        st.error(f"Face detection crashed: {_diag['fatal_error']}")
                        if _diag.get("trace"):
                            st.code(_diag["trace"])
                    counter_keys = [
                        "connections_total", "with_occ", "occ_missing",
                        "surfacebodies_ok", "faces_examined", "planar_faces",
                        "non_planar_faces", "face_pairs_examined", "pairs_parallel",
                        "pairs_offset_ok", "contacts_found", "connections_out",
                    ]
                    present = [{"stage": k, "count": _diag.get(k, 0)}
                               for k in counter_keys if k in _diag]
                    if present:
                        st.dataframe(pd.DataFrame(present), width="stretch", hide_index=True)
                    errs = _diag.get("errors") or []
                    if errs:
                        st.markdown("**First COM errors encountered:**")
                        for e in errs:
                            st.markdown(f"- `{e}`")

            _face_s = ((st.session_state.block1_raw or {}).get("timings") or {}).get("connections_face_s")
            if _face_s is not None:
                st.caption(f"⏱Last face-contact pass: {_face_s:.1f}s.")

    # STEP 4 — Bolted connection design schedule
    elif step.startswith("4"):
        header("4. Bolted connections",
               "Every CONTACT PATCH is checked as its own bolted connection (AISC Chapter J). "
               "A joint with three patches is three rows here — each carries its own share of "
               "the member-end force and each can fail independently. "
               "Two inputs come from the connected parts; the rest follow the BAC standard bolt detail.")

        if st.session_state.joints.empty:
            st.warning("Load geometry in Step 2 first so the app knows which joints need connection data.")
        else:
            ids = joint_ids()
            keys = connection_keys(st.session_state.block1_raw or {})
            if not keys:
                keys = [(f"{j}:1", j) for j in ids]
                st.warning(
                    "No extracted connections in this model, so the schedule falls back to one "
                    "placeholder row per joint. Run **3 · Extract connections** to resolve the "
                    "real contact patches — until then a multi-patch joint cannot be described.",
                )

            # Build the working schedule: bolted defaults per connection, preserving
            # any prior edits and normalizing material/bolt to the valid dropdown values.
            prev = {}
            sched = st.session_state.connection_schedule
            if not sched.empty and "connection_id" in sched.columns:
                prev = {str(r["connection_id"]): r for _, r in sched.iterrows()}
            rows = []
            for cid, jid in keys:
                row = default_bolted_row(cid, jid)
                p = prev.get(str(cid))
                if p is not None:
                    for k in ["gauge", "sheet_t_in", "n_fasteners", "diameter_in", "edge_dist_in", "notes"]:
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

            # Re-run autopopulation here, not only at upload: Step 3's face pass
            # renumbers connections and expands one centerline pair into several
            # patches, so ids like C018 do not exist when the ingest-time pass
            # runs and would otherwise sit at gauge 0 / thickness 0 forever. Only
            # rows still blank are touched, so edits above survive.
            working = autofill_gauge_thickness(
                working, st.session_state.block1_raw or {}, st.session_state.members)

            # Show autopopulation status when Block 1 data has gauge/thickness values.
            if st.session_state.block1_raw is not None:
                gauge_filled = int((working["gauge"] > 0).sum())
                filled = int((working["sheet_t_in"] > 0).sum())
                if gauge_filled > 0:
                    st.info(f"Gauge autopopulated from Inventor iProperties for {gauge_filled} of {len(working)} connections "
                            f"(governing = smallest resulting thickness across connected members); Thickness (in) is "
                            f"translated from gauge via the standard gauge chart. Read per CONNECTION where the patch's "
                            f"two parts are known, otherwise per JOINT. Review before running.")
                elif filled > 0:
                    st.info(f"Thickness (sheet_t_in) autopopulated from Block 1 for {filled} of {len(working)} connections. Review before running.")
                else:
                    st.caption("Gauge/thickness could not be read from Block 1 — Gauge iProperty and section wall_thickness may be zero or unclassified. Enter values manually.")

                # Name the rows nothing could be read for. Silence here is what
                # makes a blank row look like a bug rather than missing source
                # data, so say which connection and why it could not resolve.
                blank = working.loc[(working["gauge"] <= 0) & (working["sheet_t_in"] <= 0),
                                    ["connection_id", "joint_id"]]
                if not blank.empty and (gauge_filled or filled):
                    _lookup = _member_lookup_from_block1(st.session_state.block1_raw or {})
                    _cmap = {str(c.get("connection_id")): c
                             for c in ((st.session_state.block1_raw or {}).get("connections") or [])
                             if isinstance(c, dict)}
                    _why = []
                    for _, br in blank.iterrows():
                        cid = str(br["connection_id"])
                        c = _cmap.get(cid) or {}
                        a, b = str(c.get("member_a") or ""), str(c.get("member_b") or "")
                        if not (a in _lookup or b in _lookup):
                            reason = "neither connected part matched a Block 1 member"
                        else:
                            reason = "connected parts carry no Gauge iProperty and no wall_thickness"
                        _why.append(f"**{cid}** ({br['joint_id'] or 'no joint'}) — {reason}")
                    st.warning(
                        "No gauge or thickness could be read for "
                        f"{len(blank)} connection(s); enter them by hand:\n\n- "
                        + "\n- ".join(_why[:10])
                        + ("\n- …" if len(_why) > 10 else ""),
                    )

            with st.expander("How to fill this in", expanded=False):
                st.markdown(
                    "Each row is one bolted CONNECTION — one contact patch between two parts. "
                    "A joint where three members meet through three patches has three rows, and the "
                    "joint is only as good as its worst row.\n\n"
                    "The row set is fixed here: it comes from the connections extracted in Step 3. "
                    "To add or remove a connection, do it there.\n\n"
                    "- **Connection material** — material of the governing (thinner / weaker) connected part. "
                    "Sets the ultimate strength Fu. Pick from the list.\n"
                    "- **Gauge** — sheet-metal gauge of the governing connected part, read from the Inventor "
                    "``Gauge`` iProperty. For joints whose members are different gauges, the smallest resulting "
                    "thickness governs (gauge numbering runs opposite to thickness: 8 ga = 0.153 in, 18 ga = "
                    "0.044 in). Leave at 0 to enter Thickness manually instead.\n"
                    "- **Thickness (in)** — actual sheet/wall thickness of the governing part. Drives bearing and "
                    "tear-out. When Gauge > 0 this is recalculated from the standard gauge chart and overrides any "
                    "manual entry. *Required* — a joint left at 0 is flagged for review, not silently failed.\n"
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
            display_cols = ["connection_id", "joint_id", "material_grade", "gauge", "sheet_t_in",
                            "n_fasteners", "diameter_in", "edge_dist_in", "fastener_type", "notes"]
            edited = st.data_editor(
                working[display_cols],
                width="stretch",
                hide_index=True,
                # Fixed, not dynamic: the row set is derived from the extracted
                # connections, so a row added here was silently dropped on the next
                # rerun. Add connections in Step 3, not in this grid.
                num_rows="fixed",
                key=editor_key("bolted_sched", working["connection_id"]),
                disabled=["connection_id", "joint_id"],
                column_config={
                    "connection_id": st.column_config.TextColumn("Connection ID", help="One contact patch. Read-only — the row set comes from Step 3."),
                    "joint_id": st.column_config.TextColumn("Joint ID", help="The joint this patch belongs to. Read-only. Results roll up to it."),
                    "material_grade": st.column_config.SelectboxColumn("Connection material", options=MATERIALS, help="Governing connected part. Sets Fu."),
                    "gauge": st.column_config.NumberColumn("Gauge", min_value=0, step=1,
                                                            help="Sheet-metal gauge (Inventor iProperty). 0 = unknown; "
                                                                 "enter Thickness manually instead. When set, Thickness "
                                                                 "is recalculated from the standard gauge chart."),
                    "sheet_t_in": st.column_config.NumberColumn("Thickness (in)", min_value=0.0, format="%.4f",
                                                                 help="Governing part thickness. Required. Overridden "
                                                                      "by Gauge when Gauge > 0."),
                    "n_fasteners": st.column_config.NumberColumn("# Fasteners", min_value=0, step=1),
                    "diameter_in": st.column_config.NumberColumn("Bolt dia. (in)", min_value=0.0, format="%.4f"),
                    "edge_dist_in": st.column_config.NumberColumn("Edge dist. Le (in)", min_value=0.0, format="%.3f"),
                    "fastener_type": st.column_config.SelectboxColumn("Bolt grade", options=BOLTS, help="AISC strength group."),
                    "notes": st.column_config.TextColumn("Notes"),
                },
            )

            # Gauge (when known) is the source of truth for thickness: translate via the
            # standard gauge chart and overwrite any manually-typed Thickness. Unrecognized
            # gauges are left alone and flagged in Notes rather than silently guessed.
            edited = edited.copy()
            unrecognized_gauges = []
            for i, r in edited.iterrows():
                g = r.get("gauge", 0)
                try:
                    g_int = int(round(float(g)))
                except (TypeError, ValueError):
                    g_int = 0
                if g_int > 0:
                    t = thickness_from_gauge(g_int)
                    if t is not None:
                        edited.at[i, "gauge"] = g_int
                        edited.at[i, "sheet_t_in"] = t
                    else:
                        unrecognized_gauges.append(str(r.get("connection_id", i)))
            if unrecognized_gauges:
                st.warning(f"Gauge not in the standard chart for connection(s): {', '.join(unrecognized_gauges)}. "
                           f"Thickness left as entered — review.")

            # Persist: force the single bolted type and restore the full schema.
            edited["connection_type"] = "bolted_bracket_joint"
            edited["template_name"] = "Bolted sheet-metal connection"
            st.session_state.connection_schedule = ensure_connection_columns(edited)

            if st.button("Reset to bolted defaults", width="stretch"):
                st.session_state.connection_schedule = ensure_connection_columns(
                    pd.DataFrame([default_bolted_row(cid, jid) for cid, jid in keys]))
                st.rerun()
    # STEP 5
    elif step.startswith("5"):
        header("5. Loads and combinations", "Enter wind and seismic parameters to run the structural FE solve and compute joint forces.")

        st.markdown("### Structural FE analysis (Block 3 · PyNite)")
        if not BLOCK3_AVAILABLE:
            st.warning(f"PyNite engine unavailable: {BLOCK3_IMPORT_ERROR}. Install with `pip install PyNiteFEA` to enable the solve. You can still upload or edit a force table manually below.")
        elif st.session_state.block1_raw is None:
            st.info("Load Block 1 geometry in Step 2 first, then return here to run the solve.")
        else:
            from engine.loads.asce7_lrfd_loads import (
                PARAMETRIC_PRESETS, tributary_wind_area, parametric_to_point_loads,
            )
            b1 = st.session_state.block1_raw
            mass_unit = "lb"

            # ── Row 1: preset ─────────────────────────────────────────────────
            preset = st.selectbox("Qualification preset", list(PARAMETRIC_PRESETS.keys()),
                                  help="Equipment wind/seismic qualification level. Choosing a preset fills Wind and Seismic fields. Select 'Custom' to enter your own values.")
            # Supports are auto-detected as a collinear pair — 'pinned' alone is a
            # mechanism in 3D — so the solve is always run fully fixed.
            fixity = "fixed"

            pv = PARAMETRIC_PRESETS.get(preset) or {}
            sds = float(st.session_state.seismic_sds)

            # ── Row 2: Wind | Seismic inputs ─────────────────────────────────
            p1, p2 = st.columns(2)
            with p1:
                st.markdown("**Wind**")
                wind_psf = st.number_input("Design pressure q (psf)", min_value=0.0, step=5.0, value=float(pv.get("wind_psf", 30.0)))
                # Wind axis is fixed to X (+) by structural-engineering directive:
                # assemblies are always modelled with the broad face normal to X,
                # so wind striking that face is the standard, conservative case. The
                # ± combinations are still applied automatically downstream.
                wind_dir = "X (+)"
                # Each member collects wind on its own face width only (open frame).
                # A wind-collecting panel across the members is handled by typing the
                # true windward face area into "Projected area A" below.
                _adj_gap = None
                area_info = tributary_wind_area(b1, wind_dir, "in", adjacency_gap_in=_adj_gap)
                auto_area = area_info["total_area_ft2"]
                _wax = area_info.get("wind_axis")
                if area_info["n_collectors"] > 0:
                    area_help = (f"Total projected wind area A ≈ {auto_area:.1f} ft² = sum over "
                                 f"{area_info['n_collectors']} vertical members of (height × own face "
                                 f"width). Wind on {_wax}-axis (⟂ frame), width "
                                 f"measured on {area_info['span_axis']}-axis. Override when known.")
                else:
                    area_help = ("No vertical collecting members found. Enter the windward face area manually.")
                ref_area = st.number_input("Projected area A (ft²)", min_value=0.0, step=1.0, value=round(auto_area, 1), help=area_help)
                # keep the resolved axis so the combinations/solve use the same direction
                wind_dir = f"{_wax} (+)" if _wax else "X (+)"
                if area_info["n_collectors"] > 0:
                    _tag = "auto" if area_info.get("auto") else "forced"
                    st.caption(f"A = {auto_area:.1f} ft² · height × own face width · wind ⟂ frame on "
                               f"**{_wax}**-axis ({_tag}) · {area_info['n_collectors']} members.")
                    if auto_area <= 0.01:
                        st.warning("Projected area is ~0 — the wind axis is in the frame plane (edge-on). "
                                   "Enter the true windward face area manually above.")
                else:
                    st.caption("Projected area could not be formed — enter the true windward face area.")
            with p2:
                st.markdown("**Seismic**")
                sds = st.number_input("SDS (g)", min_value=0.0, step=0.05, value=float(pv.get("sds", sds)), help="Design short-period spectral acceleration from the site hazard. Drives vertical seismic Ev = 0.2·SDS·D in the combinations.")
                coeff = st.number_input("Horizontal coefficient (g)", min_value=0.0, step=0.05, value=float(pv.get("sds", sds)), help="Lateral seismic load coefficient. Defaults to SDS for a rigid mount at the top of the structure (x/h = 1.0). QE = coeff × operating weight.")
                seismic_dir = st.selectbox("Seismic direction", ["X (+)"], help="Horizontal direction seismic QE acts on; combinations apply it reversibly. Fixed to X per structural engineering direction.")

            # ── Row 3: secondary parameters ───────────────────────────────────
            # Operating (unit) weight is always derived from the Block 1 model
            # self-weight — autopopulated — so there is no manual override input.
            rho = st.number_input("Redundancy factor ρ", min_value=1.0, max_value=1.3, step=0.1, value=1.3, help="ASCE 7-22 §12.3.4. Use ρ = 1.3 for most cooling tower frames. ρ = 1.0 only when a licensed engineer has confirmed sufficient redundancy.")

            wind_loads, seismic_loads, info = parametric_to_point_loads(
                b1, wind_psf=wind_psf, ref_area_ft2=ref_area, wind_dir=wind_dir,
                seismic_coeff_g=coeff, seismic_dir=seismic_dir, source_mass_unit=mass_unit,
                operating_weight_lb=None)

            # ── Load summary cards ────────────────────────────────────────────
            st.markdown(metric_cards_html([
                ("Unit weight", f"{info['operating_weight_lb']:,.0f} lb", "blue"),
                ("Total wind W", f"{info['wind_total_lb']:,.0f} lb", "teal"),
                ("Total seismic QE", f"{info['seismic_total_lb']:,.0f} lb", "yellow"),
            ]), unsafe_allow_html=True)
            st.caption(f"Loads distributed across {info['n_loaded_nodes']} joints proportionally to self-weight. LRFD combinations apply wind and seismic reversibly (± directions).")

            if st.button("Run structural FE solve", type="primary", width="stretch"):
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

                    # The solve is per JOINT (one FE node each). Distribute each
                    # member-end action across the contact patches that physically
                    # carry it, so the checks in Step 7 can be per CONNECTION.
                    dt, _detail, dreview = connection_demands(b1, res)
                    st.session_state.demand_table = dt
                    st.session_state.demand_review = dreview

                    st.success(f"Solve complete. Supports {meta['support_nodes']} ({meta['support_fixity']}), {meta['n_elements']} elements, {len(meta['combos'])} combos. Joint forces and combinations updated.")
                    cov = demand_coverage(dt)
                    if cov["total"]:
                        msg = (f"Demand resolved onto {cov['checked']} of {cov['total']} "
                               f"connections.")
                        if cov["unchecked"]:
                            st.warning(msg + f" {cov['unchecked']} could not be resolved and "
                                             f"will report as **Unchecked** — they are not "
                                             f"given a substitute demand.")
                        else:
                            st.info(msg)
                    flagged = sorted({n['member'] for n in meta.get('build_notes', []) if n.get('needs_review')})
                    if flagged:
                        st.info("Sections using fallback / needs-review properties: " + ", ".join(flagged))
                    if dreview:
                        with st.expander(f"Demand distribution notes ({len(dreview)})"):
                            for f in dreview:
                                st.write("- " + f)
                else:
                    st.error(f"Solve failed: {res.get('error')}. {res.get('hint', '')}")

        st.divider()
        st.markdown("### Editable load combinations")
        st.session_state.combo_table = st.data_editor(st.session_state.combo_table, width="stretch", hide_index=True, num_rows="dynamic")
        st.markdown("### Editable joint force demands")
        st.session_state.joint_forces = ensure_force_columns(st.data_editor(
            ensure_force_columns(st.session_state.joint_forces),
            width="stretch",
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
        st.dataframe(checklist[["check", "status", "message", "recommended_action"]], width="stretch", hide_index=True)

    # STEP 7
    elif step.startswith("7"):
        header("7. Run and review results",
               "Checks run per CONNECTION — each contact patch against its own share of the "
               "member-end force. The joint rollup above the detail is reporting only: a joint "
               "is only as good as its worst patch.")
        checklist = validation_checklist(st.session_state.joints, st.session_state.members, st.session_state.combo_table, st.session_state.connection_schedule, st.session_state.joint_forces, st.session_state.units)
        ready, message = is_ready_to_run(checklist)
        if not ready:
            st.error(message)
            st.dataframe(checklist[["check", "status", "message", "recommended_action"]], width="stretch", hide_index=True)

        demand_table = st.session_state.demand_table
        if demand_table is None or demand_table.empty:
            st.warning(
                "No per-connection demand yet. Run the **structural FE solve in Step 5** — it is "
                "what distributes each member-end action across the contact patches that carry "
                "it. Without it every connection reports Unchecked, by design: a demand that did "
                "not come from equilibrium is not a result.",
            )

        run_anyway = st.checkbox("Run even with warnings/blockers", value=False, disabled=ready)
        if st.button("Run connection screening", type="primary", width="stretch",
                     disabled=(not ready and not run_anyway)):
            st.session_state.analysis_results = evaluate_connection_demands(
                st.session_state.demand_table, st.session_state.connection_schedule)
            st.session_state.critical_summary = critical_joint_summary(st.session_state.analysis_results)
            st.session_state.balance_table = joint_balance_table(st.session_state.joint_forces)
            st.success("Connection screening complete.")
        results = st.session_state.analysis_results
        if results.empty:
            st.info("No results yet. Run the connection screening after validation passes.")
        else:
            counts = results["status"].value_counts().to_dict()
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("OK", counts.get("OK", 0))
            c2.metric("Warnings", counts.get("Warning", 0))
            c3.metric("Not OK", counts.get("Not OK", 0))
            c4.metric("Incomplete", counts.get("Incomplete", 0))
            c5.metric("Unchecked", counts.get("Unchecked", 0),
                      help="No demand could be resolved onto this patch. Never given a "
                           "substitute number — see the reason column.")
            c6.metric("Ignored", counts.get("Ignored", 0))
            if counts.get("Not OK", 0) > 0:
                st.error("Overall result: review required. At least one connection's demand exceeds the screening capacity.")
            elif counts.get("Unchecked", 0) > 0:
                st.warning(f"Overall result: incomplete coverage. {counts['Unchecked']} connection(s) "
                           f"could not be checked at all — the model is not fully screened.")
            elif counts.get("Warning", 0) > 0:
                st.warning("Overall result: close to limit. Review yellow connections before release.")
            elif counts.get("Incomplete", 0) > 0:
                st.warning("Overall result: incomplete. Some connections need more input data.")
            else:
                st.success("Overall result: every connection is green in the current screening model.")

            st.markdown("### Joint rollup")
            st.caption("Worst connection per joint. `n_connections` is how many contact patches "
                       "the joint is actually made of; `governing_connection` is the one that set "
                       "its status.")
            st.dataframe(st.session_state.critical_summary, width="stretch", hide_index=True)

            st.markdown("### Per-connection results")
            fc1, fc2 = st.columns([3, 2])
            all_status = sorted(results["status"].dropna().unique().tolist())
            status_filter = fc1.multiselect("Filter by status", all_status, default=all_status)
            joint_opts = sorted(results["joint_id"].dropna().astype(str).unique().tolist())
            joint_filter = fc2.multiselect("Filter by joint", joint_opts, default=[])
            view = results[results["status"].isin(status_filter)] if status_filter else results
            if joint_filter:
                view = view[view["joint_id"].astype(str).isin(joint_filter)]
            friendly_cols = ["connection_id", "joint_id", "combo_id", "status", "percent_used",
                             "tension_lbf", "shear_lbf", "screening_capacity_lbf",
                             "plain_language_issue", "suggested_fix"]
            st.dataframe(view[friendly_cols], width="stretch", hide_index=True)
            with st.expander("Engineering details: demand split, capacity basis, detection method"):
                st.dataframe(view, width="stretch", hide_index=True)
            with st.expander("Demand distribution: how each member-end action was split"):
                st.caption("From the FE solve. `governing_member` is which member's patch group "
                           "produced the worse of the two estimates for that patch.")
                st.dataframe(st.session_state.demand_table, width="stretch", hide_index=True)
                for f in st.session_state.demand_review or []:
                    st.write("- " + f)
            with st.expander("SAP-like joint balance / reaction-style table"):
                st.dataframe(st.session_state.balance_table, width="stretch", hide_index=True)

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
            "Connection Demand": st.session_state.demand_table,
            "Results": st.session_state.analysis_results,
            "Joint Rollup": st.session_state.critical_summary,
            "Joint Balance": st.session_state.balance_table,
        }
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("Download Excel workbook", to_excel_bytes(sheets), "bac_joint_screening_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
        with c2:
            st.download_button("Download project config JSON", to_json_bytes(config), "project_joint_config.json", "application/json", width="stretch")
        with c3:
            try:
                pdf_bytes = to_pdf_bytes(config, st.session_state.critical_summary, st.session_state.analysis_results, checklist)
                st.download_button("Download PDF summary", pdf_bytes, "bac_joint_screening_report.pdf", "application/pdf", width="stretch")
            except Exception as exc:
                st.warning(f"PDF export unavailable: {exc}")
        st.markdown("### Templates")
        t1, t2, t3 = st.columns(3)
        t1.download_button("Connection schedule CSV", to_csv_bytes(connection_schedule_template()), "connection_schedule_template.csv", "text/csv", width="stretch")
        t2.download_button("Joint force CSV", to_csv_bytes(default_joint_forces(["J001", "J002"], ["LRFD-01", "LRFD-02"])), "joint_force_template.csv", "text/csv", width="stretch")
        t3.download_button("Load combinations CSV", to_csv_bytes(load_combination_template()), "load_combination_template.csv", "text/csv", width="stretch")
        st.info("Exported reports include a clear limitation statement because the current capacities are placeholders.")

from engine.cost_page import render_cost_run
def draw_financial_analysis_tab() -> None:
    st.header("Financial Analysis")
    st.caption("Pick the fabrication region, units, and per-part cutting/forming methods, then run the cost engine.")

    if st.session_state.block1_raw is None:
        st.info("Load Block 1 geometry in Step 2 (Structural Analysis tab) first.")
        return

    cost_parts, cost_joints, review_flags = build_cost_inputs(st.session_state.block1_raw)
    flags_df = pd.DataFrame(review_flags)

    # Missing-data feedback: which parts/joints Block 1 could not fully extract.
    # A joint missing its length is silently dropped from the cost run, so this
    # is the place that surfaces an incomplete costing.
    if not flags_df.empty:
        st.warning(
            f"{len(flags_df)} field(s) are missing from Block 1 and are required for costing. "
            "See the details below.",
        )
        with st.expander("Missing data / review flags", expanded=True):
            st.dataframe(flags_df, width="stretch", hide_index=True)

    # ── Cost engine run (Milestone 2) ──────────────────────────────
    render_cost_run(cost_parts, cost_joints)


_MFG_STATUS_BADGE = {
    "TUBE LASER OK": "Tube Laser OK",
    "AUTOMATED PANEL BENDER OK": "Automated Panel Bender OK",
    "MANUAL PANEL BENDER OK": "Manual Panel Bender OK",
    "NOT MANUFACTURABLE": "Not Manufacturable",
    "REVIEW REQUIRED": "Review Required",
}


def draw_manufacturability_tab() -> None:
    st.header("Manufacturability Analysis")
    st.caption(
        "IJET machine-capability screening derived from Block 1 geometry. Each part is checked "
        "against the Tube Laser, Automated Panel Bender, and Manual Panel Bender envelopes; the "
        "cheapest eligible process is selected automatically."
    )

    if st.session_state.block1_raw is None:
        st.info("Load Block 1 geometry in Step 2 (Structural Analysis tab) first.")
        return

    if "mfg_rows" not in st.session_state or st.session_state.get("mfg_rows_source") != id(st.session_state.block1_raw):
        rows, _ = build_manufacturability_inputs(st.session_state.block1_raw)
        st.session_state["mfg_rows"] = rows
        st.session_state["mfg_rows_source"] = id(st.session_state.block1_raw)

    rows_df = pd.DataFrame(st.session_state["mfg_rows"])

    st.markdown("### Part inputs")
    st.caption(
        "Edit x (width), y (length), gauge, and material where missing or incorrect, then run the "
        "assessment below. `*_source` columns show where each value came from in Block 1."
    )
    edited_df = st.data_editor(
        rows_df,
        hide_index=True,
        width="stretch",
        disabled=["part_identifier", "dimension_source", "gauge_source", "material_source"],
        column_config={
            "part_identifier": st.column_config.TextColumn("Part"),
            "quantity": st.column_config.NumberColumn("Qty", min_value=1, step=1),
            "x_width_in": st.column_config.NumberColumn("x width (in)", min_value=0.0, format="%.3f"),
            "y_length_in": st.column_config.NumberColumn("y length (in)", min_value=0.0, format="%.3f"),
            "gauge": st.column_config.SelectboxColumn("Gauge", options=list(SUPPORTED_GAUGES)),
            "material": st.column_config.SelectboxColumn("Material", options=list(SUPPORTED_MATERIALS)),
            "dimension_source": st.column_config.TextColumn("Dimension source"),
            "gauge_source": st.column_config.TextColumn("Gauge source"),
            "material_source": st.column_config.TextColumn("Material source"),
        },
        key="mfg_input_editor",
    )
    st.session_state["mfg_rows"] = edited_df.to_dict("records")

    if st.button("Run manufacturability assessment", type="primary", width="stretch"):
        results, review_flags = evaluate_manufacturability_rows(st.session_state["mfg_rows"])
        st.session_state["mfg_results"] = results
        st.session_state["mfg_flags"] = review_flags

    results = st.session_state.get("mfg_results")
    if not results:
        return

    review_flags = st.session_state.get("mfg_flags") or []
    summary = summarize_results(results)

    st.markdown("### Results")
    st.markdown(metric_cards_html([
        ("Part types", summary["part_types"], "blue"),
        ("Total qty", summary["total_quantity"], "blue"),
        ("Tube Laser OK (qty)", summary["tube_laser_ok_quantity"], "green"),
        ("Auto Panel Bender OK (qty)", summary["automated_panel_bender_ok_quantity"], "teal"),
        ("Manual Panel Bender OK (qty)", summary["manual_panel_bender_ok_quantity"], "yellow"),
        ("Not manufacturable (qty)", summary["not_manufacturable_quantity"], "red" if summary["not_manufacturable_quantity"] else "green"),
        ("Review required (qty)", summary["review_required_quantity"], "yellow" if summary["review_required_quantity"] else "green"),
    ]), unsafe_allow_html=True)

    if review_flags:
        with st.expander(f"{len(review_flags)} row(s) need review", expanded=True):
            st.dataframe(pd.DataFrame(review_flags), width="stretch", hide_index=True)

    results_df = pd.DataFrame(results).copy()
    results_df["failure_reasons"] = results_df["failure_reasons"].apply(
        lambda d: "; ".join(f"{k}: {v}" for k, v in d.items()) if d else ""
    )
    results_df["status"] = results_df["status"].map(lambda s: _MFG_STATUS_BADGE.get(s, s))
    results_df = results_df.drop(columns=["manufacturable_any_process"], errors="ignore")
    st.dataframe(results_df, width="stretch", hide_index=True)

    export_results = [{k: v for k, v in r.items() if k != "manufacturable_any_process"}
                      for r in results]
    st.download_button(
        "Download manufacturability results JSON",
        data=json.dumps({"results": export_results, "review_flags": review_flags}, indent=2, default=str).encode("utf-8"),
        file_name="manufacturability_results.json",
        mime="application/json",
        width="stretch",
    )


with tab1:
    draw_structural_analysis_tab()

with tab2:
    draw_financial_analysis_tab()

with tab3:
    draw_manufacturability_tab()


def _part_to_joints(block1_raw: dict) -> dict:
    """Map each part identifier to the set of joint_ids it participates in.

    Handles both Block 1 schemas: the real Inventor schema (joint.member_names →
    member occurrence → part_number) and the simple/legacy schema (members carry
    start_joint / end_joint, identified by member_id). The part identifier uses
    the same precedence as manufacturability's _part_identifier so the keys line
    up with the manufacturability and cost tables.
    """
    result: dict = {}
    members = (block1_raw or {}).get("members") or []
    by_occ: dict = {}
    for m in members:
        if not isinstance(m, dict):
            continue
        for k in ("occurrence_path", "occurrence_name"):
            if m.get(k):
                by_occ[m[k]] = m

    def _pid(m):
        for k in ("part_number", "part_identifier", "occurrence_path",
                  "occurrence_name", "member_id"):
            if m.get(k) not in (None, ""):
                return str(m[k])
        return None

    linked = False
    for j in (block1_raw or {}).get("joints") or []:
        if not isinstance(j, dict):
            continue
        jid = str(j.get("joint_id", ""))
        for mname in (j.get("member_names") or []):
            m = by_occ.get(mname)
            if m and jid:
                p = _pid(m)
                if p:
                    result.setdefault(p, set()).add(jid)
                    linked = True

    if not linked:  # simple schema: members carry their end joints directly
        for m in members:
            if not isinstance(m, dict):
                continue
            p = _pid(m)
            if not p:
                continue
            for jk in ("start_joint", "end_joint"):
                jv = m.get(jk)
                if jv not in (None, ""):
                    result.setdefault(p, set()).add(str(jv))
    return result


def draw_joint_recommendation_tab() -> None:
    """Bolted joint summary.

    Stage 1 (Structural Analysis) establishes the DEMAND at every joint. This
    tab consumes those results -- it never re-solves -- and reports, per part:
    is the bolted connection at its joints structurally sound, can the BLM
    LT8.20 tube laser cut it, and what does it cost. IJET screens BOLTED
    connections only; welding is not part of the analysis.
    """
    st.header("Summary")
    st.caption(
        "Per part: is the bolted connection at its joints structurally sound, "
        "can the BLM LT8.20 tube laser cut it, and what does it cost."
    )

    # --- Stage 1 gate -----------------------------------------------------
    block1_raw = st.session_state.get("block1_raw")
    if block1_raw is None:
        st.info("Load a model in **Structural Analysis** to begin.")
        return

    analysis_results = st.session_state.get("analysis_results")
    if analysis_results is None or getattr(analysis_results, "empty", True):
        st.warning(
            "**Run the structural analysis first.** A joint cannot be recommended without "
            "knowing what it has to carry. Go to **Structural Analysis → Run analysis**, "
            "then come back."
        )
        return

    mfg_results = st.session_state.get("mfg_results")
    if not mfg_results:
        st.info(
            "Manufacturability has not been evaluated yet, so tube-laser results will be "
            "blank. Run **Manufacturability Analysis** for the complete picture.",
        )

    # --- Assumptions the user can change ---------------------------------
    # IJET screens bolted connections only; the weld-electrode input was removed.
    # The recommendation engine still takes an electrode argument for its internal
    # (no-longer-displayed) weld figures, so a default is passed through.
    electrode = "E70XX"
    with st.expander("Assumptions", expanded=False):
        pfd = st.number_input(
            "Personal Fatigue and Delay factor", min_value=1.0, max_value=2.0,
            value=float(DEFAULT_PFD), step=0.05,
            help=(
                "Labor allowance from the cost workbook (Joints List column T), a "
                "regional lookup in the sheet. It scales the labor half of fastener "
                "cost. Confirm the BAC value."
            ),
        )

    # --- Build the verdicts ----------------------------------------------
    demand = demand_from_analysis_results(analysis_results)
    verdicts = build_joint_verdicts(
        block1_raw,
        critical_summary=demand,
        connection_schedule=st.session_state.get("connection_schedule"),
        mfg_results=mfg_results,
        electrode=electrode,
        pfd=pfd,
    )
    if not verdicts:
        st.warning("No joints found in the Block 1 extraction.")
        return

    st.session_state["joint_verdicts"] = [v.to_dict() for v in verdicts]

    def _opt(v, key, field, default=None):
        o = v.options.get(key)
        return getattr(o, field, default) if o else default

    def _bolted_sound(v):
        """Does the BOLTED connection carry the joint's demand? True/False, or
        None when the bolted capacity or demand could not be established. IJET
        screens bolted connections only, so soundness never depends on a weld."""
        o = v.options.get("bolted")
        if not o or o.capacity_lbf is None:
            return None
        if o.percent_used is not None:
            return o.percent_used <= 100.0
        if v.demand_lbf:
            return o.capacity_lbf >= v.demand_lbf
        return True

    sound_by_joint = {v.joint_id: _bolted_sound(v) for v in verdicts}

    # --- Problem alert (bolted-only) -------------------------------------
    failing = sorted(jid for jid, s in sound_by_joint.items() if s is False)
    if failing:
        st.error(
            f"**{len(failing)} joint(s) are not structurally sound (bolted):** "
            f"{', '.join(failing)}. The parts at these joints show **No** below."
        )

    # --- Pass / fail by part ---------------------------------------------
    st.subheader("Pass / fail by part")

    # Three plain yes/no answers per part:
    #   Structurally sound? — do ALL joints this part belongs to carry their load
    #   Tube-laser OK?      — can the BLM tube laser cut this part
    #   Part cost           — fully burdened total from the cost engine
    # "?" means that input has not been produced yet (run the relevant tab).
    part_joints = _part_to_joints(block1_raw)
    tube_by_part = {str(r.get("part_identifier")): r.get("tube_laser_ok")
                    for r in (mfg_results or [])}
    cost_by_part = {str(r.get("Part Identifier")): r.get("Extended Fully Burdened Total Cost")
                    for r in ((st.session_state.get("cost_result") or {}).get("parts") or [])}

    yn = {True: "Yes", False: "No", None: "?"}
    part_ids = list(tube_by_part) or sorted(part_joints)
    part_rows = []
    for p in part_ids:
        vals = [sound_by_joint[j] for j in part_joints.get(p, set()) if j in sound_by_joint]
        if any(x is False for x in vals):
            sound = "No"                       # any joint that fails fails the part
        elif not vals or any(x is None for x in vals):
            sound = "?"                        # unlinked or a capacity couldn't be computed
        else:
            sound = "Yes"
        part_rows.append({
            "part": p,
            "sound": sound,
            "tube": yn[tube_by_part.get(p)] if p in tube_by_part else "?",
            "cost": cost_by_part.get(p),
        })

    if part_rows:
        st.dataframe(
            pd.DataFrame(part_rows), width="stretch", hide_index=True,
            column_config={
                "part": st.column_config.TextColumn("Part"),
                "sound": st.column_config.TextColumn("Structurally sound?"),
                "tube": st.column_config.TextColumn("Tube-laser OK?"),
                "cost": st.column_config.NumberColumn("Part cost", format="$%.2f"),
            },
        )
        st.caption(
            "**Structurally sound?** — every joint this part belongs to carries its load. "
            "**Tube-laser OK?** — from Manufacturability Analysis. "
            "**Part cost** — fully burdened total from the cost engine (Financial Analysis). "
            "“?” means that step hasn't been run yet."
        )
    else:
        st.info("No parts to summarize yet.")

    # Full engineering comparison — kept for the detail-minded and the CSV export,
    # but tucked away so the pass/fail table above stays clean.
    table = pd.DataFrame([{
        "joint_id": v.joint_id,
        "type": v.geom_descriptor or v.joint_type,
        "verdict": v.verdict,
        "demand_lbf": v.demand_lbf,
        "combo": v.governing_combo,
        "bolted_lbf": _opt(v, "bolted", "capacity_lbf"),
        "bolted_$": _opt(v, "bolted", "cost_usd"),
        "tube_laser": {True: "Yes", False: "No", None: "?"}[v.tube_laser_ok],
    } for v in verdicts])

    with st.expander("Full detail — bolted capacity & cost by joint"):
        st.dataframe(
            table, width="stretch", hide_index=True,
            column_config={
                "demand_lbf": st.column_config.NumberColumn("Demand (lbf)", format="%.0f"),
                "bolted_lbf": st.column_config.NumberColumn("Bolted capacity (lbf)", format="%.0f"),
                "bolted_$": st.column_config.NumberColumn("Bolted cost ($)", format="$%.2f"),
            },
        )

    # --- Per-joint detail --------------------------------------------------
    st.subheader("Joint detail")
    ids = [v.joint_id for v in verdicts]
    chosen = st.selectbox("Joint", ids, key="recommendation_joint_select")
    v = next(x for x in verdicts if x.joint_id == chosen)

    st.markdown(f"### {v.joint_id} — {v.verdict}")

    d = st.columns(3)
    d[0].metric("Geometry", v.geom_descriptor or v.joint_type or "—")
    d[1].metric("Demand", f"{v.demand_lbf:,.0f} lbf" if v.demand_lbf else "—",
                help=f"Governing combination: {v.governing_combo}" if v.governing_combo else None)
    d[2].metric("Tube laser", {True: "OK", False: "No", None: "?"}[v.tube_laser_ok])

    if v.branch_name or v.chord_name:
        st.caption(f"Branch: `{v.branch_name}` ·  Chord: `{v.chord_name}`")

    # Bolted connection, stated plainly.
    bolted = v.options.get("bolted")
    if bolted is None or bolted.capacity_lbf is None:
        st.error("Bolted capacity could not be computed for this joint.")
    elif bolted.percent_used is not None and bolted.percent_used > 100:
        st.error(f"**Bolted connection is overstressed** — {bolted.percent_used:.0f}% of capacity used.")
    else:
        _pu = f" — {bolted.percent_used:.0f}% of capacity used" if bolted.percent_used is not None else ""
        st.success(f"**Bolted connection carries the demand**{_pu}.")

    # Bolted option detail.
    if bolted is not None and bolted.capacity_lbf is not None:
        cc = st.columns(2)
        cc[0].metric("Bolted capacity", f"{bolted.capacity_lbf:,.0f} lbf",
                     delta=(f"{bolted.percent_used:.0f}% used" if bolted.percent_used is not None else None),
                     delta_color="inverse")
        if bolted.cost_usd is not None:
            cc[1].metric("Bolted cost", f"${bolted.cost_usd:,.2f}")
        st.caption(f"Governing: {bolted.governing_mode or '—'}")
        with st.expander("Basis"):
            st.caption(f"**Capacity** — {bolted.capacity_basis or '—'}")
            st.caption(f"**Cost** — {bolted.cost_basis or '—'}")

    # Actions.
    if v.actions:
        st.markdown("**Recommended actions**")
        for a in v.actions:
            st.markdown(f"- {a}")

    if v.member_processes:
        st.markdown("**Manufacturing process per member**")
        st.dataframe(
            pd.DataFrame([{"member": k, "process": p} for k, p in v.member_processes.items()]),
            width="stretch", hide_index=True,
        )

    if v.review_reasons:
        with st.expander(f"Review flags ({len(v.review_reasons)})", expanded=v.verdict == "REVIEW"):
            for r in v.review_reasons:
                st.markdown(f"- {r}")

    # --- Export ------------------------------------------------------------
    st.subheader("Export")
    e1, e2 = st.columns(2)
    e1.download_button(
        "Download recommendations (CSV)",
        data=table.to_csv(index=False).encode("utf-8"),
        file_name=f"ijet_joint_recommendations_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        width="stretch",
    )
    e2.download_button(
        "Download full audit record (JSON)",
        data=json.dumps(st.session_state["joint_verdicts"], indent=2, default=str).encode("utf-8"),
        file_name=f"ijet_joint_verdicts_{datetime.now():%Y%m%d_%H%M}.json",
        mime="application/json",
        width="stretch",
        help="Every capacity, cost, basis string, and review flag behind the table above.",
    )


with tab4:
    draw_joint_recommendation_tab()

