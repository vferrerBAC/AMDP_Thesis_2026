"""
engine/connection_demand.py
===========================
The bridge from JOINT-level FE results to CONNECTION-level demand.

Why this module exists
----------------------
Block 3 solves a centerline frame: one node per joint, member-end actions at each
node. That is the correct idealization for global analysis — a contact patch has
no stiffness we can defensibly assign, so it cannot be a node.

But nothing can be CHECKED at a joint. Every quantity a capacity equation needs
lives on the ``Connection``: ``contact_area_in2``, ``weld_length_in`` (the patch
perimeter), ``hole_count``, ``patch_frame``, ``contact_normal``. A ``Joint`` has a
location, a member list, and some angles. So:

    analysis by joint   ->   checks by connection

and this module is the ``->``. It takes the member-end actions Block 3 already
produced and hands them to ``patch_groups`` to split across the patches that
physically carry them, per load combination, then envelopes.

What it deliberately does NOT do
--------------------------------
It never invents a demand. A connection that cannot receive a patch demand — no
joint_id, no contact plane, no member-end action from the solver — comes back
with ``status="Unchecked"`` and a reason, and is counted separately. Loading such
a connection with the full member-end force would be arithmetically conservative
but practically useless: at a four-patch joint it oversizes by ~4x and the tool
stops discriminating a good connection from a bad one, which is the only thing it
is for.

The one case that LOOKS like a fallback but is not: a member whose group has
exactly one patch. There the full member-end action IS the exact demand, and
``patch_groups.distribute`` handles it directly (flagged ``single_patch``).

Units: lbf, lbf-in, inches. Global axes throughout.
"""

from __future__ import annotations

import math
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd

from engine.connection_templates import is_active_connection
from engine.patch_groups import (
    PatchDemand,
    build_patch_groups,
    distribute_all_combos,
    unresolvable_connections,
)

DEMAND_COLUMNS = [
    "connection_id", "joint_id", "member_a", "member_b", "governing_member",
    "combo_id", "tension_lbf", "bearing_lbf", "shear_lbf", "force_demand_lbf",
    "contact_area_in2", "weld_length_in", "hole_count", "detection_method",
    "status", "review_reason",
]


def _g(c: Any, key: str, default: Any = None) -> Any:
    return c.get(key, default) if isinstance(c, dict) else getattr(c, key, default)


# --------------------------------------------------------------------------- #
# Member-name resolution
# --------------------------------------------------------------------------- #
# Block 3 keys its member-end actions by whatever string sits in
# ``Joint.member_names`` (``_member_incidence`` reads that list, and the element
# name is f"{mname}__{a}_{b}"). ``Connection.member_a/member_b`` hold
# occurrence_path "or name". Usually identical; when Block 1 wrote one and the
# joints the other, an unresolved name would silently mean "no FE action" and the
# connection would vanish into Unchecked. So alias them explicitly.

_MEMBER_ALIAS_KEYS = ("occurrence_path", "occurrence_name", "part_number", "member_id")


def member_alias_map(block1_data: Mapping[str, Any]) -> Dict[str, str]:
    """{any name a member is known by -> the name Block 3 files its actions under}.

    The canonical name is the one appearing in ``Joint.member_names``, because
    that is what the solver built its element names from.
    """
    canonical = set()
    for j in block1_data.get("joints") or []:
        for n in _g(j, "member_names") or []:
            canonical.add(str(n))

    alias: Dict[str, str] = {c: c for c in canonical}
    for m in block1_data.get("members") or []:
        names = [str(_g(m, k)) for k in _MEMBER_ALIAS_KEYS
                 if _g(m, k) not in (None, "")]
        target = next((n for n in names if n in canonical), None)
        if target is None:
            continue
        for n in names:
            alias.setdefault(n, target)
    return alias


def _resolve(name: Any, alias: Mapping[str, str]) -> str:
    s = str(name or "")
    return alias.get(s, s)


# --------------------------------------------------------------------------- #
# Block 3 result -> member-end actions
# --------------------------------------------------------------------------- #


def actions_by_combo(
    block3_result: Mapping[str, Any],
) -> Dict[str, Dict[Tuple[str, str], Tuple[List[float], List[float]]]]:
    """Reshape ``joint_member_actions`` into what ``patch_groups`` wants.

    {combo: {(joint_id, member): (F_xyz, M_xyz)}}

    A member can touch a joint at both ends of two different elements (a joint
    mid-span of a member that Block 3 split into segments). Those are two
    separate actions on the SAME (joint, member) group, so they are summed: the
    group carries the net action delivered there, which is what equilibrium at
    the node says.
    """
    out: Dict[str, Dict[Tuple[str, str], Tuple[List[float], List[float]]]] = {}
    if block3_result.get("status") != "ok":
        return out
    for combo, joints in (block3_result.get("joint_member_actions") or {}).items():
        acc: Dict[Tuple[str, str], Tuple[List[float], List[float]]] = {}
        for jid, rows in joints.items():
            for r in rows:
                key = (str(jid), str(r.get("member", "")))
                F = [float(v) for v in r.get("F", [0.0, 0.0, 0.0])]
                M = [float(v) for v in r.get("M", [0.0, 0.0, 0.0])]
                if key in acc:
                    pF, pM = acc[key]
                    acc[key] = ([pF[i] + F[i] for i in range(3)],
                                [pM[i] + M[i] for i in range(3)])
                else:
                    acc[key] = (F, M)
        out[str(combo)] = acc
    return out


def _remap_connections(connections: Sequence[Any],
                       alias: Mapping[str, str]) -> List[Dict[str, Any]]:
    """Copy connections with member names normalized to the solver's keys."""
    out = []
    for c in connections:
        d = dict(c) if isinstance(c, dict) else {
            k: getattr(c, k) for k in vars(c)} if hasattr(c, "__dict__") else dict(c)
        d["member_a"] = _resolve(d.get("member_a"), alias)
        d["member_b"] = _resolve(d.get("member_b"), alias)
        out.append(d)
    return out


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #


def connection_demands(
    block1_data: Mapping[str, Any],
    block3_result: Mapping[str, Any],
) -> Tuple[pd.DataFrame, Dict[str, List[PatchDemand]], List[str]]:
    """Per-connection demand, enveloped over every load combination.

    Returns (demand_table, per-combo detail by connection_id, review_flags).

    Every connection in ``block1_data["connections"]`` appears in the table
    exactly once. Those that could not be resolved carry ``status="Unchecked"``
    and a ``review_reason``; they are never given a substitute demand.
    """
    # Inactive connections are not contacts at all, so they get no demand row --
    # an "Unchecked" row for a connection the engineer already removed reads as a
    # gap in the check rather than a removal.
    connections = [c for c in (block1_data.get("connections") or [])
                   if is_active_connection(c)]
    joints_by_id = {str(_g(j, "joint_id")): (j if isinstance(j, dict) else vars(j))
                    for j in (block1_data.get("joints") or [])}

    alias = member_alias_map(block1_data)
    remapped = _remap_connections(connections, alias)

    unresolvable = unresolvable_connections(remapped)
    per_combo_actions = actions_by_combo(block3_result)

    review: List[str] = []
    governing: Dict[str, PatchDemand] = {}
    detail: Dict[str, List[PatchDemand]] = {}

    if not per_combo_actions:
        review.append(
            "no_fe_actions: Block 3 produced no joint_member_actions "
            f"(solver status: {block3_result.get('status', 'not run')}). Every "
            "connection is Unchecked until the FE solve runs."
        )
    else:
        governing, detail, flags = distribute_all_combos(
            remapped, joints_by_id, per_combo_actions)
        review.extend(flags)

    # Which groups the solver had no action for — so an individual connection can
    # say WHY it is unchecked rather than just that it is.
    groups = build_patch_groups(remapped)
    covered = set()
    if per_combo_actions:
        any_combo = next(iter(per_combo_actions.values()))
        for key in groups:
            if key in any_combo:
                covered.update(groups[key].connection_ids)

    rows: List[Dict[str, Any]] = []
    for c in remapped:
        cid = str(_g(c, "connection_id") or "")
        d = governing.get(cid)
        if d is not None:
            normal = float(d.normal_lbf)
            row_status, reason = "Checked", "; ".join(d.needs_review)
            rows.append({
                "connection_id": cid,
                "joint_id": str(_g(c, "joint_id") or ""),
                "member_a": _g(c, "member_a") or "",
                "member_b": _g(c, "member_b") or "",
                "governing_member": d.member,
                "combo_id": d.combo_id,
                "tension_lbf": round(max(normal, 0.0), 2),
                "bearing_lbf": round(max(-normal, 0.0), 2),
                "shear_lbf": round(abs(d.shear_lbf), 2),
                "force_demand_lbf": round(math.hypot(normal, d.shear_lbf), 2),
                "contact_area_in2": _g(c, "contact_area_in2"),
                "weld_length_in": _g(c, "weld_length_in"),
                "hole_count": _g(c, "hole_count", 0),
                "detection_method": _g(c, "detection_method") or "",
                "status": row_status,
                "review_reason": reason,
            })
            continue

        if cid in unresolvable:
            reason = unresolvable[cid]
        elif not per_combo_actions:
            reason = "FE solve has not produced member-end actions yet"
        elif cid not in covered:
            reason = ("solver reported no member-end action for this patch group "
                      "(member not modeled, or only one incident joint)")
        else:
            reason = "patch group resolved but produced no demand"
        rows.append({
            "connection_id": cid,
            "joint_id": str(_g(c, "joint_id") or ""),
            "member_a": _g(c, "member_a") or "",
            "member_b": _g(c, "member_b") or "",
            "governing_member": "",
            "combo_id": "",
            "tension_lbf": None,
            "bearing_lbf": None,
            "shear_lbf": None,
            "force_demand_lbf": None,
            "contact_area_in2": _g(c, "contact_area_in2"),
            "weld_length_in": _g(c, "weld_length_in"),
            "hole_count": _g(c, "hole_count", 0),
            "detection_method": _g(c, "detection_method") or "",
            "status": "Unchecked",
            "review_reason": reason,
        })

    return pd.DataFrame(rows, columns=DEMAND_COLUMNS), detail, review


def demand_coverage(demand_table: pd.DataFrame) -> Dict[str, int]:
    """Headline counts for the UI: how much of the model actually got checked."""
    if demand_table.empty:
        return {"total": 0, "checked": 0, "unchecked": 0}
    checked = int((demand_table["status"] == "Checked").sum())
    return {"total": int(len(demand_table)),
            "checked": checked,
            "unchecked": int(len(demand_table)) - checked}
