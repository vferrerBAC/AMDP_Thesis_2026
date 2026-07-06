"""
ASCE 7-22 LRFD loads + Block-3 FE bridge for the Streamlit app.
================================================================
Turns the raw Block-1 geometry dict into the two tables the existing engine
already consumes:

  * joint_forces  -> columns  joint_id, combo_id, Fx_lbf..Mz_lbf_in
                     (fed to joint_checks.evaluate_joint_demands)
  * combo_table   -> columns  combo_id, combo_name, expression, plain_language_notes
                     (replaces load_combinations.default_load_combinations so
                      joint_forces.combo_id always matches the combo table)

Heavy lifting (PyNite frame solve) lives in engine.block3_solver; this file only
adapts its output to the app's schema and derives a member-connectivity table
for display from the real Block-1 schema (joints[].member_names), which the
app's generic parser does not pick up.

Units throughout: lb, in, lb*in (imperial).
"""
from __future__ import annotations

import math
from typing import Any, Dict, List, Optional

import pandas as pd

from engine.block3_solver import analyze_joints, DEFAULT_CONFIG

FORCE_COLUMNS = ["Fx_lbf", "Fy_lbf", "Fz_lbf", "Mx_lbf_in", "My_lbf_in", "Mz_lbf_in"]

_COMBO_META = {
    "1.4D": ("Dead only", "1.4D", "Gravity-only strength case."),
    "W1+":  ("Wind (1.2D)",              "1.2D + 1.0W",            "Wind, full gravity, +direction."),
    "W1-":  ("Wind (1.2D), rev.",        "1.2D - 1.0W",            "Wind, full gravity, reversed."),
    "W2+":  ("Uplift wind (0.9D)",       "0.9D + 1.0W",            "Wind, reduced stabilizing dead, +direction."),
    "W2-":  ("Uplift wind (0.9D), rev.", "0.9D - 1.0W",            "Wind, reduced stabilizing dead, reversed."),
    "E1+":  ("Seismic (1.2D)",           "(1.2+0.2Sds)D + rho*QE", "Seismic, full gravity + vertical Ev, +direction."),
    "E1-":  ("Seismic (1.2D), rev.",     "(1.2+0.2Sds)D - rho*QE", "Seismic, full gravity + vertical Ev, reversed."),
    "E2+":  ("Uplift seismic (0.9D)",    "(0.9-0.2Sds)D + rho*QE", "Seismic, reduced dead - Ev, +direction."),
    "E2-":  ("Uplift seismic (0.9D), rev.", "(0.9-0.2Sds)D - rho*QE", "Seismic, reduced dead - Ev, reversed."),
}


def members_dataframe_from_block1(block1_data: Dict[str, Any]) -> pd.DataFrame:
    """Build the app-schema member table from joints[].member_names.

    The app's generic parser only links members carrying start_joint/end_joint
    ids. The real Block-1 schema lists member_names on each joint, so we invert
    that to recover start_joint / end_joint for display.
    """
    cols = ["member_id", "start_joint", "end_joint", "section_type", "material"]
    joints = block1_data.get("joints", [])
    members = {m.get("occurrence_name"): m for m in block1_data.get("members", [])}

    incidence: Dict[str, List[str]] = {}
    for j in joints:
        for mname in j.get("member_names", []):
            incidence.setdefault(mname, []).append(j.get("joint_id"))
    jloc = {j.get("joint_id"): j.get("location") for j in joints}

    def _order(mname: str, jids: List[str]) -> List[str]:
        m = members.get(mname)
        if not m or len(jids) < 2:
            return jids
        s, e = m.get("start_point"), m.get("end_point")
        ax = [e[i] - s[i] for i in range(3)]
        L2 = sum(c * c for c in ax) or 1.0
        return sorted(jids, key=lambda jid: sum((jloc[jid][i] - s[i]) * ax[i]
                                                 for i in range(3)) / L2)

    rows = []
    for mname, jids in incidence.items():
        m = members.get(mname, {})
        cs = m.get("cross_section", {}) if isinstance(m, dict) else {}
        ordered = _order(mname, jids)
        rows.append({
            "member_id": mname,
            "start_joint": ordered[0] if ordered else "",
            "end_joint": ordered[-1] if len(ordered) > 1 else "",
            "section_type": str(cs.get("section_type", "unknown")),
            "material": str(cs.get("family", m.get("material", "unknown"))),
        })
    return pd.DataFrame(rows, columns=cols)


def run_lrfd_joint_analysis(
    block1_data: Dict[str, Any],
    *,
    S_DS: float = 1.0,
    rho: float = 1.3,
    wind_loads: Optional[List[Dict[str, Any]]] = None,
    seismic_h_loads: Optional[List[Dict[str, Any]]] = None,
    support_nodes: Optional[List[str]] = None,
    support_fixity: str = "fixed",
    source_length_unit: str = "in",
    source_mass_unit: str = "lb",
    include_1_4D: bool = True,
) -> Dict[str, Any]:
    """Wrapper over engine.block3_solver.analyze_joints with app defaults.

    wind_loads / seismic_h_loads: lists of {"node","FX","FY","FZ"} point loads (lb).
    Returns the solver dict (status 'ok' or 'error').
    """
    config = {
        **DEFAULT_CONFIG,
        "S_DS": float(S_DS), "rho": float(rho),
        "source_length_unit": source_length_unit,
        "source_mass_unit": source_mass_unit,
        "support_fixity": support_fixity,
        "include_1_4D": include_1_4D,
    }
    if support_nodes:
        config["support_nodes"] = support_nodes
    loads = {"wind": wind_loads or [], "seismic_h": seismic_h_loads or []}
    return analyze_joints(block1_data, config=config, loads=loads)


def _resultant(row: Dict[str, float]) -> float:
    shear = math.sqrt(row["N"] ** 2 + row["Vy"] ** 2 + row["Vz"] ** 2)
    moment = math.sqrt(row["T"] ** 2 + row["My"] ** 2 + row["Mz"] ** 2)
    return shear + moment


def joint_forces_table(result: Dict[str, Any]) -> pd.DataFrame:
    """Reduce per-member end forces to one governing demand per (joint, combo).

    The app schedules one connection per joint, so for each (joint, combo) we
    report the member end force with the largest resultant - the governing
    demand on that joint's connection. Member-local components map to the app's
    columns (the demand metric uses only invariant shear/moment magnitudes, so
    the local->global naming does not affect the result):
        N -> Fx_lbf, Vy -> Fy_lbf, Vz -> Fz_lbf,
        T -> Mx_lbf_in, My -> My_lbf_in, Mz -> Mz_lbf_in
    """
    out: List[Dict[str, Any]] = []
    if result.get("status") != "ok":
        return pd.DataFrame(columns=["joint_id", "combo_id"] + FORCE_COLUMNS)
    for combo, joints in result["joint_member_forces"].items():
        for jid, rows in joints.items():
            gov = max(rows, key=_resultant) if rows else \
                {"N": 0, "Vy": 0, "Vz": 0, "T": 0, "My": 0, "Mz": 0}
            out.append({
                "joint_id": jid, "combo_id": combo,
                "Fx_lbf": gov["N"], "Fy_lbf": gov["Vy"], "Fz_lbf": gov["Vz"],
                "Mx_lbf_in": gov["T"], "My_lbf_in": gov["My"], "Mz_lbf_in": gov["Mz"],
            })
    return pd.DataFrame(out, columns=["joint_id", "combo_id"] + FORCE_COLUMNS)


def combos_table(result: Dict[str, Any]) -> pd.DataFrame:
    """Combo table matching the combo_ids present in the solved forces."""
    combos = result.get("meta", {}).get("combos", []) if result.get("status") == "ok" else []
    rows = []
    for cid in combos:
        name, expr, note = _COMBO_META.get(cid, (cid, cid, ""))
        rows.append({"combo_id": cid, "combo_name": name,
                     "expression": expr, "plain_language_notes": note})
    return pd.DataFrame(rows, columns=["combo_id", "combo_name", "expression", "plain_language_notes"])


# --------------------------------------------------------------------------- #
#  Parametric loads:  psf / g  ->  nodal point loads
# --------------------------------------------------------------------------- #
# Equipment qualification presets (wind pressure, SDS).  At a rigid mount and
# x/h = 1.0 the horizontal seismic coefficient is taken equal to SDS (the full
# top-of-structure rigid case); swap in the ASCE 7-22 Ch.13 Fp(ap,Rp,Ip,z/h)
# equation here if a code-rigorous component force is required.
PARAMETRIC_PRESETS = {
    "30 psf / 0.3g  (SDS 0.3, rigid, x/h=1.0)": {"wind_psf": 30.0, "sds": 0.3},
    "70 psf / 0.7g  (SDS 0.7, rigid, x/h=1.0)": {"wind_psf": 70.0, "sds": 0.7},
    "Custom": None,
}
_LEN_TO_FT = {"in": 1.0 / 12.0, "ft": 1.0, "cm": 1.0 / 30.48, "mm": 1.0 / 304.8}
_COMP = {"X (+)": "FX", "Z (+)": "FZ", "Y (+)": "FY"}


def nodal_dead_weights(block1_data: Dict[str, Any],
                       source_mass_unit: str = "lb") -> Dict[str, float]:
    """Lbf of self-weight lumped to each joint - mirrors the solver's lumping
    (half of each member's weight to its two extreme end joints)."""
    from engine.block3_solver import _MASS_TO_LBF
    kmass = _MASS_TO_LBF[source_mass_unit]
    joints = block1_data.get("joints", [])
    members = {m.get("occurrence_name"): m for m in block1_data.get("members", [])}
    jloc = {j.get("joint_id"): j.get("location") for j in joints}
    incidence: Dict[str, List[str]] = {}
    for j in joints:
        for mn in j.get("member_names", []):
            incidence.setdefault(mn, []).append(j.get("joint_id"))
    w = {j.get("joint_id"): 0.0 for j in joints}
    for mn, jids in incidence.items():
        m = members.get(mn)
        if not m or len(jids) < 2:
            continue
        wt = float(m.get("dry_mass", 0.0)) * kmass
        s, e = m.get("start_point"), m.get("end_point")
        ax = [e[i] - s[i] for i in range(3)]
        L2 = sum(c * c for c in ax) or 1.0
        ends = sorted(jids, key=lambda jid: sum((jloc[jid][i] - s[i]) * ax[i]
                                                 for i in range(3)) / L2)
        w[ends[0]] += wt / 2.0
        w[ends[-1]] += wt / 2.0
    return w


def estimate_projected_area_ft2(block1_data: Dict[str, Any], wind_dir: str = "X (+)",
                                source_length_unit: str = "in") -> float:
    """Bounding-box silhouette area (ft^2) on the plane normal to the wind.
    Y is vertical; conservative envelope estimate for a clad unit - override
    with the real projected face area when known."""
    js = [j.get("location") for j in block1_data.get("joints", []) if j.get("location")]
    for m in block1_data.get("members", []):       # include member ends for true envelope
        for k in ("start_point", "end_point"):
            if m.get(k):
                js.append(m[k])
    if not js:
        return 0.0
    k = _LEN_TO_FT.get(source_length_unit, 1.0 / 12.0)
    dx = (max(p[0] for p in js) - min(p[0] for p in js)) * k
    dy = (max(p[1] for p in js) - min(p[1] for p in js)) * k
    dz = (max(p[2] for p in js) - min(p[2] for p in js)) * k
    if wind_dir.startswith("X"):
        return abs(dy * dz)
    if wind_dir.startswith("Z"):
        return abs(dy * dx)
    return abs(dx * dz)


def parametric_to_point_loads(
    block1_data: Dict[str, Any], *,
    wind_psf: float, ref_area_ft2: float, wind_dir: str,
    seismic_coeff_g: float, seismic_dir: str,
    source_mass_unit: str = "lb",
    operating_weight_lb: Optional[float] = None,
) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    """Convert (wind pressure, seismic coefficient) into nodal point loads.

    Wind:    total = wind_psf * ref_area_ft2,  distributed to joints by weight
             fraction (proxy for tributary area), applied along wind_dir.
    Seismic: per-joint horizontal QE = seismic_coeff_g * (joint self-weight)
             i.e. total = coeff * operating weight, mass-proportional (F = m*a),
             applied along seismic_dir.  Returned QE is the reference vector;
             the load combinations apply +/- and the rho factor.
    """
    w = nodal_dead_weights(block1_data, source_mass_unit)
    Wp_dead = sum(w.values())
    Wp = float(operating_weight_lb) if operating_weight_lb else (Wp_dead or 1.0)
    scale = (Wp / Wp_dead) if Wp_dead else 1.0       # scale weights to Wp

    wind_total = float(wind_psf) * float(ref_area_ft2)
    seismic_total = float(seismic_coeff_g) * Wp
    wcomp = _COMP.get(wind_dir, "FX")
    scomp = _COMP.get(seismic_dir, "FX")

    wind_loads, seismic_loads = [], []
    for node, wn in w.items():
        wn_s = wn * scale
        frac = wn_s / Wp if Wp else 0.0
        wind_loads.append({"node": node, wcomp: wind_total * frac})
        seismic_loads.append({"node": node, scomp: float(seismic_coeff_g) * wn_s})

    info = {"operating_weight_lb": round(Wp, 1), "dead_weight_lb": round(Wp_dead, 1),
            "wind_total_lb": round(wind_total, 1), "seismic_total_lb": round(seismic_total, 1),
            "n_loaded_nodes": len(w)}
    return wind_loads, seismic_loads, info
