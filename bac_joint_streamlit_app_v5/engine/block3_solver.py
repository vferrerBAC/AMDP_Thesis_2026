"""
Block 3 - Joint-focused LRFD linear static solver  (IMPERIAL: lb, in, psi)
==========================================================================
Consumes Block 1 geometry-recognition JSON, builds a 3D centerline frame in
PyNite, applies ASCE 7-22 LRFD load combinations, solves, and returns the
per-joint results (support reactions + every incident member's end forces)
that Block 4 connection checks consume.

Web-app integration:  call analyze_joints(block1_dict, config=?, loads=?) -> dict.
No file I/O, no globals, no notebook state.  JSON in -> JSON-serialisable out.

Working/output units:  force = lb, length = in, moment = lb*in, stress = psi.
"""
from __future__ import annotations
import math
from typing import Any

# PyNite is imported lazily inside the solve so that importing this module
# (and the Streamlit app) does not hard-require it being installed.

# ----------------------------------------------------------------------------
# Units.  We solve in imperial (lb, in, psi).  Block-1 source units are stated
# separately so any authoring unit can be ingested.  source_mass_unit is the
# one still pending the Block-1 COM self-check (lbmass vs kg) -> one flag flip.
# ----------------------------------------------------------------------------
_LEN_TO_IN = {"in": 1.0, "ft": 12.0,
              "cm": 0.39370078740157, "mm": 0.039370078740157}
_MASS_TO_LBF = {"lb": 1.0, "kg": 2.2046226218488}     # source mass -> lbf weight

DEFAULT_CONFIG: dict[str, Any] = {
    "source_length_unit": "in",   # Block-1 geometry/section authoring unit
    "source_mass_unit":   "lb",   # Block-1 dry_mass unit  ("lb" or "kg")
    "E": 29_000_000.0,            # psi (all structural steel ~ grade-independent)
    "nu": 0.30,
    "vertical_axis": "Y",         # gravity acts -Y (legs span in Y in this model)
    "support_nodes": None,        # None -> auto: joints with is_support_candidate
    "support_fixity": "fixed",    # "fixed" (all 6) or "pinned" (DX,DY,DZ).
    #   NOTE: the 2 Block-1 support candidates are a collinear pinned pair,
    #   a rigid-body mechanism in 3D (free rotation about the line joining
    #   them).  "fixed" removes it.  Use "pinned" only with >=3 non-collinear
    #   supports.
    # ---- ASCE 7-22 seismic scalars (HOOKs - supply real site values) --------
    "S_DS": 1.0,                  # design spectral accel, short period
    "rho": 1.3,                   # redundancy factor
    "include_1_4D": True,         # add gravity-only combo (often governs)
}


def _unit_factors(cfg: dict) -> dict[str, float]:
    klen = _LEN_TO_IN[cfg["source_length_unit"]]
    return {"len": klen, "area": klen ** 2, "inertia": klen ** 4,
            "mass_to_lbf": _MASS_TO_LBF[cfg["source_mass_unit"]]}


# ============================== section props ===============================
def _thinwall_channel(d: float, b: float, t: float) -> dict[str, float]:
    """Block-1-consistent thin-wall channel properties (source length units)."""
    A = t * (d + 2 * b)
    Iz = t * d ** 3 / 12.0 + b * t * d ** 2 / 2.0          # strong (web depth)
    xbar = b ** 2 / (d + 2 * b)
    Iy = d * t * xbar ** 2 + 2 * (t * b ** 3 / 12.0
                                  + b * t * (b / 2.0 - xbar) ** 2)  # weak
    J = (d + 2 * b) * t ** 3 / 3.0                          # open-section torsion
    return {"A": A, "Iy": Iy, "Iz": Iz, "J": J}


def _resolve_props(cs: dict) -> tuple[dict[str, float], str]:
    """Return (props, source_tag). Uses Block-1 props if present, else fallback."""
    if all(cs.get(k) is not None for k in ("A", "Iy", "Iz", "J")):
        return ({k: float(cs[k]) for k in ("A", "Iy", "Iz", "J")},
                cs.get("props_method") or "block1")
    d, b, t = cs.get("depth"), cs.get("width"), cs.get("wall_thickness")
    if None in (d, b, t):
        raise ValueError("section has neither properties nor depth/width/wall")
    return _thinwall_channel(float(d), float(b), float(t)), "fallback_thinwall_channel"


# ============================== connectivity ================================
def _member_incidence(joints: list[dict]) -> dict[str, list[str]]:
    inc: dict[str, list[str]] = {}
    for j in joints:
        for mname in j["member_names"]:
            inc.setdefault(mname, []).append(j["joint_id"])
    return inc


def _order_joints_along(member: dict, jids: list[str],
                        jloc: dict[str, list[float]]) -> list[str]:
    """Sort the joints a member touches by their projection along its axis."""
    s = member["start_point"]; e = member["end_point"]
    ax = [e[i] - s[i] for i in range(3)]
    L2 = sum(c * c for c in ax) or 1.0
    def t(jid):
        p = jloc[jid]
        return sum((p[i] - s[i]) * ax[i] for i in range(3)) / L2
    return sorted(jids, key=t)


def _dist_in(a: list[float], b: list[float], klen: float) -> float:
    return math.sqrt(sum(((a[i] - b[i]) * klen) ** 2 for i in range(3)))


# ================================ build =====================================
def _build_model(b1: dict, cfg: dict, uf: dict):
    from Pynite import FEModel3D
    klen = uf["len"]
    model = FEModel3D()
    G = cfg["E"] / (2.0 * (1.0 + cfg["nu"]))
    model.add_material("steel", cfg["E"], G, cfg["nu"], 0.284)  # rho psi-unused

    joints = b1["joints"]
    members = {m["occurrence_name"]: m for m in b1["members"]}
    jloc = {j["joint_id"]: j["location"] for j in joints}

    for j in joints:                                   # node at each joint centroid
        p = j["location"]
        model.add_node(j["joint_id"], p[0] * klen, p[1] * klen, p[2] * klen)

    inc = _member_incidence(joints)
    notes: list[dict] = []
    skipped: list[str] = []
    elem_count = 0

    for mname, jids in inc.items():
        m = members.get(mname)
        if m is None:
            notes.append({"member": mname, "issue": "in joints but not in members[] - skipped"})
            continue
        if len(jids) < 2:
            skipped.append(mname)
            notes.append({"member": mname, "issue": "only one incident joint - not modeled"})
            continue

        props, tag = _resolve_props(m["cross_section"])
        if tag.startswith("fallback") or m["cross_section"].get("needs_review"):
            notes.append({"member": mname, "section": m["cross_section"].get("section_type"),
                          "props_source": tag, "needs_review": True})
        sname = f"sec_{mname}"
        model.add_section(sname, props["A"] * uf["area"],
                          props["Iy"] * uf["inertia"],
                          props["Iz"] * uf["inertia"],
                          props["J"]  * uf["inertia"])

        ordered = _order_joints_along(m, jids, jloc)
        for a, b in zip(ordered[:-1], ordered[1:]):     # consecutive segments
            model.add_member(f"{mname}__{a}_{b}", a, b, "steel", sname, rotation=0.0)
            elem_count += 1

    sup = cfg["support_nodes"] or [j["joint_id"] for j in joints
                                   if j.get("is_support_candidate")]
    fixed = cfg["support_fixity"] == "fixed"
    for n in sup:
        model.def_support(n, True, True, True, fixed, fixed, fixed)

    meta = {"support_nodes": sup, "support_fixity": cfg["support_fixity"],
            "n_elements": elem_count, "skipped_members": skipped,
            "build_notes": notes}
    return model, meta, joints


# ================================ loads =====================================
def _apply_loads(model, b1: dict, cfg: dict, uf: dict, loads: dict) -> None:
    up = cfg["vertical_axis"].upper()                  # gravity acts -<up>
    joints = b1["joints"]
    members = {m["occurrence_name"]: m for m in b1["members"]}
    inc = _member_incidence(joints)
    jloc = {j["joint_id"]: j["location"] for j in joints}

    # --- D : lump half of each member's self-weight (lbf) to its 2 end joints
    nodal_D: dict[str, float] = {j["joint_id"]: 0.0 for j in joints}
    for mname, jids in inc.items():
        m = members.get(mname)
        if m is None or len(jids) < 2:
            continue
        w = float(m.get("dry_mass", 0.0)) * uf["mass_to_lbf"]      # lbf, downward
        ends = _order_joints_along(m, jids, jloc)
        nodal_D[ends[0]] += w / 2.0
        nodal_D[ends[-1]] += w / 2.0
    for nid, w in nodal_D.items():
        if w:
            model.add_node_load(nid, f"F{up}", -w, case="D")

    # --- W : user point loads (lbf, reference dir; reversed in combos)
    for ld in loads.get("wind", []):
        for comp in ("FX", "FY", "FZ"):
            if ld.get(comp):
                model.add_node_load(ld["node"], comp, float(ld[comp]), case="W")

    # --- QE : user horizontal seismic point loads (lbf; rho applied in combos)
    for ld in loads.get("seismic_h", []):
        for comp in ("FX", "FY", "FZ"):
            if ld.get(comp):
                model.add_node_load(ld["node"], comp, float(ld[comp]), case="E")


# ============================== combinations ================================
def _add_combos(model, cfg: dict) -> list[str]:
    sds, rho = cfg["S_DS"], cfg["rho"]
    combos: dict[str, dict] = {}
    if cfg["include_1_4D"]:
        combos["1.4D"] = {"D": 1.4}
    combos["W1+"] = {"D": 1.2, "W": 1.0};  combos["W1-"] = {"D": 1.2, "W": -1.0}
    combos["W2+"] = {"D": 0.9, "W": 1.0};  combos["W2-"] = {"D": 0.9, "W": -1.0}
    combos["E1+"] = {"D": 1.2 + 0.2 * sds, "E":  rho}
    combos["E1-"] = {"D": 1.2 + 0.2 * sds, "E": -rho}
    combos["E2+"] = {"D": 0.9 - 0.2 * sds, "E":  rho}
    combos["E2-"] = {"D": 0.9 - 0.2 * sds, "E": -rho}
    for name, f in combos.items():
        model.add_load_combo(name, f)
    return list(combos.keys())


# =============================== extract ====================================
def _round(x): return round(float(x), 3)

def _extract(model, joints: list[dict], combos: list[str], uf: dict) -> dict:
    klen = uf["len"]
    jset = {j["joint_id"] for j in joints}
    sup = {n for n in model.nodes if model.nodes[n].support_DX}

    touch: dict[str, list[tuple[str, str, float]]] = {jid: [] for jid in jset}
    for ename, el in model.members.items():
        i, j = el.i_node, el.j_node
        Lin = _dist_in([i.X / klen, i.Y / klen, i.Z / klen],
                       [j.X / klen, j.Y / klen, j.Z / klen], klen)
        if i.name in jset:
            touch[i.name].append((ename, "i", 0.0))
        if j.name in jset:
            touch[j.name].append((ename, "j", Lin))

    reactions: dict[str, dict] = {}
    jmf: dict[str, dict] = {}
    for c in combos:
        reactions[c] = {n: {"FX": _round(model.nodes[n].RxnFX[c]),
                            "FY": _round(model.nodes[n].RxnFY[c]),
                            "FZ": _round(model.nodes[n].RxnFZ[c]),
                            "MX": _round(model.nodes[n].RxnMX[c]),
                            "MY": _round(model.nodes[n].RxnMY[c]),
                            "MZ": _round(model.nodes[n].RxnMZ[c])} for n in sup}
        jmf[c] = {}
        for jid, items in touch.items():
            rows = []
            for ename, end, x in items:
                el = model.members[ename]
                rows.append({"member": ename.split("__")[0], "element": ename,
                             "end": end,
                             "N":  _round(el.axial(x, c)),
                             "Vy": _round(el.shear("Fy", x, c)),
                             "Vz": _round(el.shear("Fz", x, c)),
                             "T":  _round(el.torque(x, c)),
                             "My": _round(el.moment("My", x, c)),
                             "Mz": _round(el.moment("Mz", x, c))})
            jmf[c][jid] = rows
    return {"reactions": reactions, "joint_member_forces": jmf}


# ================================ public ====================================
def analyze_joints(block1_data: dict, config: dict | None = None,
                   loads: dict | None = None) -> dict:
    """
    block1_data : parsed Block-1 JSON (dict with 'members' and 'joints').
    config      : overrides for DEFAULT_CONFIG (E, S_DS, rho, source units, ...).
    loads       : {"wind":[{"node","FX","FY","FZ"}...],
                   "seismic_h":[{"node","FX","FY","FZ"}...]}  point loads in lb.
    Returns a JSON-serialisable dict keyed by load combination (imperial units).
    """
    cfg = {**DEFAULT_CONFIG, **(config or {})}
    loads = loads or {"wind": [], "seismic_h": []}
    uf = _unit_factors(cfg)

    try:
        import Pynite  # noqa: F401
    except Exception as exc:
        return {"status": "error", "error": f"PyNite not installed: {exc}",
                "hint": "pip install PyNiteFEA"}

    model, meta, joints = _build_model(block1_data, cfg, uf)
    _apply_loads(model, block1_data, cfg, uf, loads)
    combos = _add_combos(model, cfg)

    try:
        model.analyze_linear(check_statics=True)
    except Exception as exc:
        return {"status": "error", "error": f"{type(exc).__name__}: {exc}",
                "hint": "model unstable - check support_nodes / connectivity",
                "meta": meta}

    results = _extract(model, joints, combos, uf)
    return {
        "status": "ok",
        "meta": {
            "units": "force=lb, length=in, moment=lb*in, stress=psi",
            "source_length_unit": cfg["source_length_unit"],
            "source_mass_unit": cfg["source_mass_unit"],
            "idealization": "rigid centerline frame", "design_basis": "ASCE 7-22 LRFD",
            "E_psi": cfg["E"], "S_DS": cfg["S_DS"], "rho": cfg["rho"],
            "combos": combos, **meta,
        },
        **results,
    }
