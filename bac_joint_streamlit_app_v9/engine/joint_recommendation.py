"""Joint recommendation — the Milestone 4 integration layer.

This is the module that answers the question the whole tool exists to answer:

    "Here is a joint. Should it be bolted or welded, can the tube laser make it,
     and what does each option cost?"

It does NOT re-run any analysis. Stage 1 (load the full frame, solve the FEA)
has already established the DEMAND at every joint; that is the only way to know
what a joint has to carry. This layer reads those existing results and joins
them to capacity, manufacturability, and cost:

    Block 1   Joint.member_names   -> which members meet here
              Joint.member_roles   -> which is branch, which is chord
              Joint.geom_descriptor-> T / Y / K / X
    Block 3   critical_summary     -> demand (lbf) and governing load combo
    Capacity  bolted_capacity.py   -> bolted capacity for this joint
              weld_capacity.py     -> welded capacity for this joint
    Mfg       manufacturability.py -> can the BLM LT8.20 cut these members
    Cost      joining_cost.py      -> bolted cost and welded cost for this joint

and emits one ``JointVerdict`` per joint.

THE JOIN CHAIN (the thing that makes this work):

    joint.member_names  ->  occurrence_path  ->  Member record
                        ->  part_identifier  ->  manufacturability result

``member_names`` holds ``occurrence_path`` (Block 1 ``_name_from_contact``
prefers the unique path), while manufacturability keys on ``part_number`` where
one exists (``manufacturability._part_identifier``). ``_resolve_part_key``
below walks that chain in the same precedence order so the two line up.

CAPACITY-PER-DOLLAR: the same geometry drives both sides. Weld length
(capacity template col P = branch width + branch height) is fed straight into
welding cost (Joints List col AA), and the bolt count from the connection
schedule drives both bolted capacity and fastener cost. So

    capacity_per_dollar = capacity_lbf / cost_usd

compares two options that were costed on the same joint, not two unrelated
numbers.

Everything uncertain routes to ``needs_review`` with a reason rather than
silently defaulting -- consistent with the rest of the pipeline.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Mapping, Optional, Sequence

from engine.bolted_capacity import bolted_capacity
from engine.weld_capacity import weld_capacity
from engine.joining_cost import (
    fastener_cost,
    welding_cost,
    fastener_type_for,
    DEFAULT_PFD,
)

# Verdicts, worst-first (used for sorting the summary table).
VERDICT_RANK = {
    "NOT MANUFACTURABLE": 4,
    "NO ADEQUATE OPTION": 3,
    "REVIEW": 2,
    "OK": 1,
}


# ---------------------------------------------------------------------------
# Data contracts
# ---------------------------------------------------------------------------


@dataclass
class ConnectionOption:
    """One way of realising a joint (bolted or welded), fully costed."""
    connection_type: str                    # "bolted" | "welded"
    capacity_lbf: Optional[float] = None
    governing_mode: str = ""
    cost_usd: Optional[float] = None
    capacity_per_dollar: Optional[float] = None   # lbf per USD
    utilization: Optional[float] = None           # demand / capacity
    percent_used: Optional[float] = None
    adequate: Optional[bool] = None               # capacity >= demand
    needs_review: bool = False
    capacity_basis: str = ""
    cost_basis: str = ""
    review_reasons: List[str] = field(default_factory=list)

    @property
    def available(self) -> bool:
        return self.capacity_lbf is not None and self.capacity_lbf > 0


@dataclass
class JointVerdict:
    """The single handoff between this layer and the Streamlit UI.

    Serializable with ``asdict`` so it can be logged as an audit record next to
    the existing cost/manufacturability audit workbooks.
    """
    joint_id: str
    geom_descriptor: str = ""               # "T" / "Y" / "K" / "X" from Block 1
    joint_type: str = ""
    member_names: List[str] = field(default_factory=list)
    branch_name: str = ""
    chord_name: str = ""

    # --- demand, from Stage 1 (already solved) ---
    demand_lbf: Optional[float] = None
    governing_combo: str = ""
    structural_status: str = ""             # worst_status from critical_summary

    # --- the connections this joint is actually made of ---
    # A joint is a CLUSTER of contact patches; each is checked on its own and the
    # worst one governs the joint. These name which one did, so the joint-level
    # verdict stays traceable to the connection it came from.
    n_connections: int = 0
    n_unchecked_connections: int = 0
    governing_connection_id: str = ""

    # --- the comparison ---
    options: Dict[str, ConnectionOption] = field(default_factory=dict)
    recommended: Optional[str] = None       # "bolted" | "welded" | None
    recommendation_rationale: str = ""

    # --- manufacturability (BLM LT8.20) ---
    member_processes: Dict[str, str] = field(default_factory=dict)
    tube_laser_ok: Optional[bool] = None    # True only if EVERY member is TL-cuttable
    manufacturability_status: str = ""

    # --- overall ---
    verdict: str = "REVIEW"
    actions: List[str] = field(default_factory=list)
    needs_review: bool = False
    review_reasons: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


# ---------------------------------------------------------------------------
# Join-chain resolution
# ---------------------------------------------------------------------------


_PART_KEY_PRECEDENCE = ("part_number", "part_identifier", "occurrence_path",
                        "occurrence_name", "member_id")


def _resolve_part_key(member: Mapping[str, Any]) -> str:
    """Same precedence as ``manufacturability._part_identifier`` so a Member
    resolves to the key its manufacturability result is filed under."""
    for key in _PART_KEY_PRECEDENCE:
        value = member.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def index_members(members: Sequence[Mapping[str, Any]]) -> Dict[str, Mapping[str, Any]]:
    """Index Block 1 members by EVERY name a joint might refer to them by.

    ``Joint.member_names`` holds ``occurrence_path``, but be forgiving: index by
    occurrence_name and part_number too, so a joint that was written with a
    different key still resolves instead of silently dropping the member.
    """
    index: Dict[str, Mapping[str, Any]] = {}
    for m in members:
        for key in ("occurrence_path", "occurrence_name", "part_number", "member_id"):
            value = m.get(key)
            if value not in (None, "") and str(value) not in index:
                index[str(value)] = m
    return index


def _branch_and_chord(joint: Mapping[str, Any]) -> tuple[str, str]:
    """Read branch/chord straight out of Block 1's ``member_roles``.

    Block 1 assigns roles as chord/branch for hollow sections, and
    primary/crossing or primary/leg in other configurations. Treat the
    chord-like role as the chord and the other as the branch.
    """
    roles = joint.get("member_roles") or {}
    branch = chord = ""
    for name, role in roles.items():
        r = str(role or "").strip().lower()
        if r in ("chord", "primary", "through") and not chord:
            chord = str(name)
        elif r in ("branch", "crossing", "leg", "secondary", "ending") and not branch:
            branch = str(name)

    names = [str(n) for n in (joint.get("member_names") or [])]
    # Fall back to positional order only when roles are absent entirely.
    if not chord and not branch and len(names) >= 2:
        chord, branch = names[0], names[1]
    elif not branch:
        branch = next((n for n in names if n != chord), "")
    elif not chord:
        chord = next((n for n in names if n != branch), "")
    return branch, chord


# ---------------------------------------------------------------------------
# Building the two connection options
# ---------------------------------------------------------------------------


def _section(member: Mapping[str, Any]) -> Mapping[str, Any]:
    return member.get("cross_section") or {}


def build_welded_conn(branch: Mapping[str, Any], chord: Mapping[str, Any],
                      electrode: str = "E70XX") -> Dict[str, Any]:
    """Assemble a welded-capacity row from two Block 1 members.

    The welded template is a tube-to-tube model, so it wants the BRANCH section
    (width, height, wall) and the CHORD wall. Block 1 already carries all of it.
    """
    bs, cs = _section(branch), _section(chord)
    return {
        "branch_material": branch.get("material") or branch.get("NCx_Material"),
        "branch_gauge": bs.get("gauge"),
        "branch_t_in": bs.get("wall_thickness"),
        "branch_width_in": bs.get("width"),
        "branch_height_in": bs.get("depth"),
        "chord_material": chord.get("material") or chord.get("NCx_Material"),
        "chord_gauge": cs.get("gauge"),
        "chord_t_in": cs.get("wall_thickness"),
        "fastener_type": electrode,
    }


def evaluate_bolted_option(conn: Mapping[str, Any], demand_lbf: Optional[float],
                           pfd: float = DEFAULT_PFD) -> ConnectionOption:
    """Bolted capacity (AISC Ch. J) + fastener cost (Joints List cols U-X)."""
    opt = ConnectionOption(connection_type="bolted")

    cap = bolted_capacity(dict(conn))
    opt.capacity_lbf = cap.shear_lbf
    opt.governing_mode = cap.governing_mode
    opt.capacity_basis = cap.basis
    if cap.needs_review:
        opt.needs_review = True
        opt.review_reasons.append(f"bolted capacity: {cap.basis}")

    cost = fastener_cost(
        conn.get("n_fasteners"),
        fastener_type_for(conn.get("connection_type") or "bolted_bracket_joint"),
        pfd=pfd,
    )
    opt.cost_usd = cost.cost_usd
    opt.cost_basis = cost.basis
    if cost.needs_review:
        opt.needs_review = True
        opt.review_reasons.append(f"bolted cost: {cost.basis}")

    _finalize_option(opt, demand_lbf)
    return opt


def evaluate_welded_option(conn: Mapping[str, Any], demand_lbf: Optional[float],
                           pfd: float = DEFAULT_PFD) -> ConnectionOption:
    """Welded capacity (template cols O-V) + welding cost (Joints List cols Z-AC).

    The weld length that sets capacity (col P) is the SAME number fed to cost
    (col AA). That is what makes the two options comparable.
    """
    opt = ConnectionOption(connection_type="welded")

    cap = weld_capacity(dict(conn))
    opt.capacity_lbf = cap.shear_lbf
    opt.governing_mode = cap.governing_mode
    opt.capacity_basis = cap.basis
    if cap.needs_review:
        opt.needs_review = True
        opt.review_reasons.append(f"welded capacity: {cap.basis}")

    cost = welding_cost(cap.weld_length_in, include_setup=True, pfd=pfd)
    opt.cost_usd = cost.cost_usd
    opt.cost_basis = cost.basis
    if cost.needs_review:
        opt.needs_review = True
        opt.review_reasons.append(f"welded cost: {cost.basis}")

    _finalize_option(opt, demand_lbf)
    return opt


def _finalize_option(opt: ConnectionOption, demand_lbf: Optional[float]) -> None:
    """Derive utilization, adequacy, and capacity-per-dollar."""
    if opt.capacity_lbf and opt.cost_usd and opt.cost_usd > 0:
        opt.capacity_per_dollar = opt.capacity_lbf / opt.cost_usd

    if demand_lbf is None or opt.capacity_lbf is None:
        return
    if opt.capacity_lbf <= 0:
        opt.adequate = False
        return

    opt.utilization = demand_lbf / opt.capacity_lbf
    opt.percent_used = opt.utilization * 100.0
    opt.adequate = opt.capacity_lbf >= demand_lbf


# ---------------------------------------------------------------------------
# The recommendation rule
# ---------------------------------------------------------------------------


def recommend(options: Mapping[str, ConnectionOption],
              demand_lbf: Optional[float]) -> tuple[Optional[str], str]:
    """Pick the better connection type.

    The rule, in order:
      1. An option that cannot carry the demand is never recommended, however
         cheap it is. Adequacy is a gate, not a term in a score.
      2. Among adequate options, prefer the higher CAPACITY PER DOLLAR -- the
         most structure bought per dollar spent.
      3. If neither is adequate, recommend nothing and say so.
      4. If demand is unknown, fall back to capacity-per-dollar but flag that
         the choice is not demand-checked.
    """
    usable = {k: o for k, o in options.items() if o.available}
    if not usable:
        return None, "Neither a bolted nor a welded capacity could be computed for this joint."

    if demand_lbf is None:
        best = max(usable, key=lambda k: usable[k].capacity_per_dollar or 0.0)
        return best, (f"No demand available for this joint, so this is a cost-efficiency "
                      f"comparison only: {best} gives the most capacity per dollar. "
                      f"Not checked against demand.")

    adequate = {k: o for k, o in usable.items() if o.adequate}

    if not adequate:
        best = max(usable, key=lambda k: usable[k].capacity_lbf or 0.0)
        o = usable[best]
        return None, (f"Neither option carries the {demand_lbf:,.0f} lbf demand. "
                      f"The strongest is {best} at {o.capacity_lbf:,.0f} lbf "
                      f"({o.percent_used:.0f}% used). Increase the connection or "
                      f"resize the members.")

    if len(adequate) == 1:
        only = next(iter(adequate))
        o = adequate[only]
        other = next((k for k in usable if k != only), None)
        tail = ""
        if other:
            oo = usable[other]
            tail = (f" The {other} option only reaches {oo.capacity_lbf:,.0f} lbf "
                    f"against a {demand_lbf:,.0f} lbf demand.")
        return only, (f"{only.capitalize()} is the only option that carries the demand "
                      f"({o.capacity_lbf:,.0f} lbf capacity vs {demand_lbf:,.0f} lbf, "
                      f"{o.percent_used:.0f}% used).{tail}")

    # Both work: buy the most capacity per dollar.
    best = max(adequate, key=lambda k: adequate[k].capacity_per_dollar or 0.0)
    other = next(k for k in adequate if k != best)
    b, o = adequate[best], adequate[other]

    if b.capacity_per_dollar and o.capacity_per_dollar:
        ratio = b.capacity_per_dollar / o.capacity_per_dollar
        return best, (
            f"Both options carry the {demand_lbf:,.0f} lbf demand, so cost decides. "
            f"{best.capitalize()}: {b.capacity_lbf:,.0f} lbf for ${b.cost_usd:,.2f} "
            f"({b.capacity_per_dollar:,.0f} lbf/$). "
            f"{other.capitalize()}: {o.capacity_lbf:,.0f} lbf for ${o.cost_usd:,.2f} "
            f"({o.capacity_per_dollar:,.0f} lbf/$). "
            f"{best.capitalize()} buys {ratio:.1f}x more capacity per dollar."
        )

    return best, (f"Both options carry the demand; {best} is recommended, but cost data "
                  f"was incomplete so the comparison is on capacity alone.")


# ---------------------------------------------------------------------------
# Per-joint evaluation
# ---------------------------------------------------------------------------


def evaluate_joint(joint: Mapping[str, Any],
                   member_index: Mapping[str, Mapping[str, Any]],
                   demand: Optional[Mapping[str, Any]] = None,
                   bolted_conns: Optional[Sequence[Mapping[str, Any]]] = None,
                   mfg_by_part: Optional[Mapping[str, Mapping[str, Any]]] = None,
                   electrode: str = "E70XX",
                   pfd: float = DEFAULT_PFD) -> JointVerdict:
    """Produce one JointVerdict. Pure function -- no Streamlit, no Excel, no I/O."""
    jid = str(joint.get("joint_id", ""))
    names = [str(n) for n in (joint.get("member_names") or [])]
    branch_name, chord_name = _branch_and_chord(joint)

    v = JointVerdict(
        joint_id=jid,
        geom_descriptor=str(joint.get("geom_descriptor") or ""),
        joint_type=str(joint.get("joint_type") or ""),
        member_names=names,
        branch_name=branch_name,
        chord_name=chord_name,
    )

    if joint.get("needs_review"):
        v.needs_review = True
        reason = joint.get("review_reason") or "flagged during joint detection"
        v.review_reasons.append(f"Block 1: {reason}")

    # --- demand, from Stage 1 --------------------------------------------
    if demand:
        raw = demand.get("demand_lbf", demand.get("max_demand_lbf"))
        try:
            v.demand_lbf = float(raw) if raw not in (None, "") else None
        except (TypeError, ValueError):
            v.demand_lbf = None
        v.governing_combo = str(demand.get("governing_combo") or "")
        v.structural_status = str(demand.get("worst_status") or "")
    if v.demand_lbf is None:
        v.needs_review = True
        v.review_reasons.append(
            "No demand for this joint. Run the structural analysis (Stage 1) first -- "
            "a joint cannot be recommended without knowing what it has to carry."
        )

    # --- resolve members --------------------------------------------------
    branch = member_index.get(branch_name)
    chord = member_index.get(chord_name)
    missing = [n for n in names if n not in member_index]
    if missing:
        v.needs_review = True
        v.review_reasons.append(f"members not found in Block 1 output: {', '.join(missing)}")

    # --- option A: welded --------------------------------------------------
    if branch and chord:
        welded_conn = build_welded_conn(branch, chord, electrode=electrode)
        v.options["welded"] = evaluate_welded_option(welded_conn, v.demand_lbf, pfd=pfd)
    else:
        v.needs_review = True
        v.review_reasons.append(
            "could not identify a branch/chord pair, so no welded option was evaluated "
            "(the welded template is a tube-to-tube model)"
        )

    # --- option B: bolted --------------------------------------------------
    # One row per CONNECTION, not per joint. Every patch at this joint is a
    # separate bolted connection with its own thickness, bolt count and material,
    # so each is evaluated and the WORST governs the joint. (This used to index
    # the schedule by joint_id into a dict, which silently kept whichever row
    # happened to be last and threw the rest away.)
    rows = list(bolted_conns or [])
    v.n_connections = len(rows)
    if rows:
        scored = [(str(r.get("connection_id") or ""),
                   evaluate_bolted_option(r, v.demand_lbf, pfd=pfd))
                  for r in rows]
        # Worst = highest utilization; where utilization is unknown, lowest capacity.
        def _worse(item):
            _, o = item
            return (o.utilization if o.utilization is not None else float("inf"),
                    -(o.capacity_lbf or 0.0))
        cid, worst = max(scored, key=_worse)
        v.options["bolted"] = worst
        v.governing_connection_id = cid
        if len(rows) > 1:
            v.review_reasons.append(
                f"joint has {len(rows)} contact patches; the bolted verdict is the "
                f"governing one ({cid or 'unnamed'}). Check the per-connection table "
                f"for the others."
            )
    else:
        v.needs_review = True
        v.review_reasons.append(
            "no connection-schedule rows for this joint, so no bolted option was evaluated"
        )

    # --- manufacturability on the BLM LT8.20 ------------------------------
    if mfg_by_part:
        flags: List[Optional[bool]] = []
        for name in names:
            member = member_index.get(name)
            if not member:
                continue
            result = mfg_by_part.get(_resolve_part_key(member))
            if not result:
                continue
            process = result.get("selected_process")
            v.member_processes[name] = str(process) if process else "NONE"
            flags.append(result.get("tube_laser_ok"))

        if flags and all(f is True for f in flags):
            v.tube_laser_ok = True
            v.manufacturability_status = "TUBE LASER OK"
        elif any(f is None for f in flags) or not flags:
            v.tube_laser_ok = None
            v.manufacturability_status = "REVIEW REQUIRED"
            v.needs_review = True
            v.review_reasons.append("manufacturability could not be assessed for every member")
        else:
            v.tube_laser_ok = False
            v.manufacturability_status = "NOT TUBE LASER"
            blocked = [n for n, p in v.member_processes.items() if p == "NONE"]
            if blocked:
                v.manufacturability_status = "NOT MANUFACTURABLE"

    # --- recommendation ---------------------------------------------------
    v.recommended, v.recommendation_rationale = recommend(v.options, v.demand_lbf)

    for opt in v.options.values():
        if opt.needs_review:
            v.needs_review = True
            v.review_reasons.extend(opt.review_reasons)

    # --- overall verdict --------------------------------------------------
    if v.manufacturability_status == "NOT MANUFACTURABLE":
        v.verdict = "NOT MANUFACTURABLE"
    elif v.options and v.recommended is None and v.demand_lbf is not None:
        v.verdict = "NO ADEQUATE OPTION"
    elif v.needs_review:
        v.verdict = "REVIEW"
    else:
        v.verdict = "OK"

    v.actions = _actions_for(v)
    return v


def _actions_for(v: JointVerdict) -> List[str]:
    """Plain-language next steps -- what the manager actually reads."""
    actions: List[str] = []

    if v.recommended:
        opt = v.options[v.recommended]
        detail = f"{opt.capacity_lbf:,.0f} lbf capacity"
        if opt.cost_usd is not None:
            detail += f" at ${opt.cost_usd:,.2f}"
        if opt.percent_used is not None:
            detail += f", {opt.percent_used:.0f}% used"
        actions.append(f"Detail this {v.geom_descriptor or 'joint'} as {opt.connection_type.upper()} "
                       f"({detail}).")

    if v.verdict == "NO ADEQUATE OPTION":
        actions.append("Neither connection type carries the demand. Increase the bolt group "
                       "or weld size, upgrade the section, or reduce the demand on this joint.")

    if v.tube_laser_ok is True:
        actions.append("All members cut on the BLM LT8.20 tube laser as-is.")
    elif v.tube_laser_ok is False:
        offenders = [n for n, p in v.member_processes.items() if p != "Tube Laser"]
        if offenders:
            actions.append("Not fully tube-laser cuttable. Members needing another process: "
                           + ", ".join(offenders) + ".")
    if v.verdict == "NOT MANUFACTURABLE":
        blocked = [n for n, p in v.member_processes.items() if p == "NONE"]
        actions.append("No available process can make: " + ", ".join(blocked)
                       + ". Resize these members.")

    if v.needs_review:
        actions.append("Flagged for review -- see reasons before relying on this result.")

    return actions


# ---------------------------------------------------------------------------
# Batch entry point (what the Streamlit tab calls)
# ---------------------------------------------------------------------------


def build_joint_verdicts(block1_data: Mapping[str, Any],
                         critical_summary: Any = None,
                         connection_schedule: Any = None,
                         mfg_results: Optional[Sequence[Mapping[str, Any]]] = None,
                         electrode: str = "E70XX",
                         pfd: float = DEFAULT_PFD) -> List[JointVerdict]:
    """Evaluate every joint in a Block 1 extraction.

    ``critical_summary`` and ``connection_schedule`` accept a pandas DataFrame or
    a list of dicts; ``mfg_results`` is the list emitted by
    ``manufacturability.evaluate_manufacturability_rows``. All are optional --
    a missing input degrades that joint to ``needs_review`` with a stated reason
    rather than failing.
    """
    members = block1_data.get("members", block1_data.get("parts", [])) or []
    joints = block1_data.get("joints", []) or []
    member_index = index_members(members)

    demand_by_joint = _records_by(critical_summary, "joint_id")
    conns_by_joint = _records_grouped_by(connection_schedule, "joint_id")
    mfg_by_part = _records_by(mfg_results, "part_identifier")

    verdicts = [
        evaluate_joint(
            joint,
            member_index,
            demand=demand_by_joint.get(str(joint.get("joint_id", ""))),
            bolted_conns=conns_by_joint.get(str(joint.get("joint_id", ""))),
            mfg_by_part=mfg_by_part,
            electrode=electrode,
            pfd=pfd,
        )
        for joint in joints
    ]

    verdicts.sort(key=lambda v: (-VERDICT_RANK.get(v.verdict, 0), v.joint_id))
    return verdicts


def demand_from_analysis_results(analysis_results: Any) -> List[Dict[str, Any]]:
    """Extract the governing demand per joint from Block 3's per-combo results.

    IMPORTANT: take the MAXIMUM ``force_demand_lbf`` across load combinations,
    not the combo that ``critical_joint_summary`` reports as governing.

    Those are different things. ``critical_joint_summary`` ranks combos by
    STATUS, and status is utilisation -- demand over *the scheduled connection's*
    capacity. That makes its "governing combo" a function of the bolted design
    that happens to be in the connection schedule. Demand itself is a property of
    the structure and does not depend on how the joint is connected, so a fair
    bolted-vs-welded comparison must be made against the worst demand the joint
    actually sees.

    Returns records shaped for ``build_joint_verdicts(critical_summary=...)``.
    """
    if analysis_results is None:
        return []
    records = analysis_results
    if hasattr(analysis_results, "to_dict"):
        if getattr(analysis_results, "empty", False):
            return []
        records = analysis_results.to_dict("records")

    worst: Dict[str, Dict[str, Any]] = {}
    for row in records or []:
        jid = row.get("joint_id")
        if jid in (None, ""):
            continue
        try:
            demand = float(row.get("force_demand_lbf"))
        except (TypeError, ValueError):
            continue
        jid = str(jid)
        current = worst.get(jid)
        if current is None or demand > current["demand_lbf"]:
            worst[jid] = {
                "joint_id": jid,
                "demand_lbf": demand,
                "governing_combo": str(row.get("combo_id") or ""),
                "worst_status": str(row.get("status") or ""),
            }
    return list(worst.values())


def _records_by(source: Any, key: str) -> Dict[str, Mapping[str, Any]]:
    """Index a DataFrame or list-of-dicts by a string key. Tolerates None."""
    if source is None:
        return {}
    records = source
    if hasattr(source, "to_dict"):              # pandas DataFrame
        if getattr(source, "empty", False):
            return {}
        records = source.to_dict("records")
    out: Dict[str, Mapping[str, Any]] = {}
    for rec in records or []:
        value = rec.get(key)
        if value not in (None, ""):
            out.setdefault(str(value), rec)
    return out


def _records_grouped_by(source: Any, key: str) -> Dict[str, List[Mapping[str, Any]]]:
    """Like ``_records_by`` but keeps EVERY record under a key instead of the
    first. Used for the connection schedule, where several connections legitimately
    share a joint_id and dropping the duplicates loses real load paths."""
    if source is None:
        return {}
    records = source
    if hasattr(source, "to_dict"):
        if getattr(source, "empty", False):
            return {}
        records = source.to_dict("records")
    out: Dict[str, List[Mapping[str, Any]]] = {}
    for rec in records or []:
        value = rec.get(key)
        if value not in (None, ""):
            out.setdefault(str(value), []).append(rec)
    return out


def summarize_verdicts(verdicts: Sequence[JointVerdict]) -> Dict[str, Any]:
    """Headline numbers for the tab's metric cards."""
    summary = {
        "joints": len(verdicts),
        "ok": 0,
        "review": 0,
        "no_adequate_option": 0,
        "not_manufacturable": 0,
        "recommend_bolted": 0,
        "recommend_welded": 0,
        "tube_laser_ok": 0,
        "total_cost_recommended": 0.0,
    }
    for v in verdicts:
        summary["ok"] += v.verdict == "OK"
        summary["review"] += v.verdict == "REVIEW"
        summary["no_adequate_option"] += v.verdict == "NO ADEQUATE OPTION"
        summary["not_manufacturable"] += v.verdict == "NOT MANUFACTURABLE"
        summary["recommend_bolted"] += v.recommended == "bolted"
        summary["recommend_welded"] += v.recommended == "welded"
        summary["tube_laser_ok"] += v.tube_laser_ok is True
        if v.recommended:
            cost = v.options[v.recommended].cost_usd
            if cost:
                summary["total_cost_recommended"] += cost
    return summary
