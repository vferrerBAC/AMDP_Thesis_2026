"""
engine/patch_groups.py
======================
Decision D2: distributing FE member-end demand across the contact patches that
physically carry it.

The problem
-----------
PyNite returns a force and a moment for each MEMBER END at each joint node. It
has no idea that the member is attached through metal. Physically that member
end is held by N contact patches (connections), and each patch must be checked
against the share of the demand it actually carries.

Loading every patch with the full member-end force is not conservative, it is
useless: a four-patch connection would be sized for 4x the real load, every
check would fail, and the tool would stop discriminating between a good
connection and a bad one.

The method
----------
Standard elastic connection-group analysis, with contact AREA as the stiffness
weight (the same role unit area plays in an elastic bolt group or weld group).

    Group  = every patch belonging to member M at joint J.
             Not "every patch between M and N" — if M touches two other members
             at that node, its end force distributes across all of them, or
             equilibrium is violated.

    1. Group centroid, area-weighted:
           C = sum(A_i * c_i) / sum(A_i)

    2. Reduce the FE member-end action from the work point P to C. The offset
       P -> C is real eccentricity: the patches sit off the centerline node.
           F_C = F
           M_C = M + (P - C) x F

    3. Direct component, area-proportional:
           f_i_direct = (A_i / sum(A)) * F_C

    4. Rotational component. Assume the member end is rigid and rotates by a
       vector theta about C. Patch displacement u_i = theta x r_i, patch force
       f_i = k * A_i * u_i. Summing moments:
           M_C = sum(r_i x f_i) = k * J * theta
       with the area-inertia tensor
           J = sum(A_i * (|r_i|^2 * I - r_i r_i^T))
       so
           theta_hat = pinv(J) @ M_C          (absorbs k)
           f_i_moment = A_i * (theta_hat x r_i)

    5. f_i = f_i_direct + f_i_moment, then split into the patch's own frame:
           normal component  -> tension / bearing
           in-plane component -> shear

This reduces correctly: one patch takes everything; two patches straddling a
web resist the member-end moment as a force couple rather than sharing it by
area, which is the behaviour that area-splitting a moment would get flatly
wrong.

Degenerate cases (single patch, collinear patch centroids) leave J singular in
one or more directions. pinv handles it numerically, but the moment about a
degenerate axis genuinely cannot be resisted by the patch group, so it is
reported as unresisted rather than silently dropped.

Every connection sits in TWO groups (one per member) and therefore receives two
demand estimates. They agree only if both distributions were exact, which they
are not, so the envelope is taken.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Optional, Sequence

import numpy as np

SINGULAR_TOL = 1e-6


@dataclass
class PatchGroup:
    """All patches carrying one member's end force at one joint.

    ``normals`` are ORIENTED OUT OF ``member``: each one points away from this
    group's member and toward the member it bears against. Block 1 stores a
    single ``contact_normal`` per connection — member_a's outward face normal
    (block1 ``_face_contact``: ``"normal": oa.normal``) — so the same stored
    vector points out of member_a but INTO member_b. ``build_patch_groups``
    flips it for the member_b group. Without that flip ``f . n`` would mean
    tension in one group and bearing in the other, and the two would be
    enveloped against each other as if they were the same quantity.
    """

    joint_id: str
    member: str
    connection_ids: list[str]
    areas: np.ndarray          # (n,) in^2
    centroids: np.ndarray      # (n,3) in
    normals: np.ndarray        # (n,3) unit, oriented out of `member`
    centroid: np.ndarray = field(default=None)   # (3,) area-weighted, in
    total_area: float = 0.0

    def __post_init__(self):
        self.total_area = float(self.areas.sum())
        if self.total_area <= 0:
            self.centroid = self.centroids.mean(axis=0)
        else:
            self.centroid = (
                self.areas[:, None] * self.centroids
            ).sum(axis=0) / self.total_area


@dataclass
class PatchDemand:
    """Demand on ONE patch, resolved into that patch's own frame."""

    connection_id: str
    member: str
    joint_id: str
    force: tuple[float, float, float]     # lbf, global
    normal_lbf: float                     # + = tension pulling faces apart
    shear_lbf: float                      # in-plane resultant
    moment_share: float                   # fraction of total group force
    combo_id: str = ""                    # load combination this demand came from
    needs_review: list[str] = field(default_factory=list)


def build_patch_groups(
    connections: Sequence,
    *,
    min_area_in2: float = 1e-6,
) -> dict[tuple[str, str], PatchGroup]:
    """
    Group connections by (joint_id, member). Each connection lands in exactly two
    groups — one for member_a, one for member_b. Connections with joint_id None
    are skipped; they have no FE node to draw demand from.
    """
    def _g(c, k, d=None):
        return c.get(k, d) if isinstance(c, dict) else getattr(c, k, d)

    buckets: dict[tuple[str, str], list] = {}
    for c in connections:
        jid = _g(c, "joint_id")
        area = _g(c, "contact_area_in2")
        normal = _g(c, "contact_normal")
        if not jid or not area or area < min_area_in2:
            continue
        if normal is None:
            # A centerline connection has no contact plane, so it has no patch to
            # carry force through. It is not silently given one.
            continue
        # sign = +1 for member_a (stored normal already points out of it), -1 for
        # member_b (the same normal points into member_b). See PatchGroup.
        for m, sign in ((_g(c, "member_a"), 1.0), (_g(c, "member_b"), -1.0)):
            buckets.setdefault((jid, m), []).append((c, sign))

    groups: dict[tuple[str, str], PatchGroup] = {}
    for (jid, m), items in buckets.items():
        groups[(jid, m)] = PatchGroup(
            joint_id=jid,
            member=m,
            connection_ids=[_g(c, "connection_id") for c, _ in items],
            areas=np.array([_g(c, "contact_area_in2") for c, _ in items], float),
            centroids=np.array([_g(c, "location") for c, _ in items], float),
            normals=np.array([np.asarray(_g(c, "contact_normal"), float) * s
                              for c, s in items], float),
        )
    return groups


def area_inertia_tensor(areas: np.ndarray, r: np.ndarray) -> np.ndarray:
    """J = sum(A_i * (|r_i|^2 I - r_i r_i^T)) — the group's area-inertia tensor."""
    J = np.zeros((3, 3))
    for A, ri in zip(areas, r):
        J += A * (float(ri @ ri) * np.eye(3) - np.outer(ri, ri))
    return J


def distribute(
    group: PatchGroup,
    force_lbf: Sequence[float],
    moment_lbf_in: Sequence[float],
    node_location_in: Sequence[float],
    combo_id: str = "",
) -> list[PatchDemand]:
    """
    Distribute one member-end action across the group's patches.

    force_lbf         : the action the MEMBER DELIVERS into the joint node, global
                        XYZ. This is ``block3_solver._global_end_action``'s "F",
                        already negated from PyNite's nodes-on-member convention.
                        Passing the un-negated vector flips every tension result
                        to bearing and back.
    moment_lbf_in     : the matching member-end moment, global XYZ
    node_location_in  : the joint work point (Joint.location)
    """
    F = np.asarray(force_lbf, float)
    M = np.asarray(moment_lbf_in, float)
    P = np.asarray(node_location_in, float)

    A = group.areas
    C = group.centroid
    r = group.centroids - C
    n = group.normals
    total_A = group.total_area

    flags: list[str] = []

    if total_A <= 0:
        return [
            PatchDemand(cid, group.member, group.joint_id, (0.0, 0.0, 0.0),
                        0.0, 0.0, 0.0, combo_id, ["zero_group_area"])
            for cid in group.connection_ids
        ]

    # Step 2: move the action from the work point to the group centroid.
    M_C = M + np.cross(P - C, F)
    ecc = float(np.linalg.norm(P - C))
    if ecc > 1e-6:
        flags.append(f"eccentricity_{ecc:.2f}in: patch group offset from work point")

    # Step 3: direct, area-proportional.
    f_direct = (A[:, None] / total_A) * F

    # Step 4: rotational.
    if len(A) == 1:
        # One patch resists everything, including all of M_C — but a single
        # planar patch cannot develop a couple about its own normal without
        # relying on the bolt group / weld path inside it. That is the capacity
        # module's problem, not ours; hand the moment down untouched.
        f_moment = np.zeros((1, 3))
        flags.append("single_patch: full member-end moment passed to the patch")
    else:
        J = area_inertia_tensor(A, r)
        w, _ = np.linalg.eigh(J)
        scale = max(abs(w).max(), 1e-12)
        if (abs(w) / scale < SINGULAR_TOL).any():
            # Collinear or coplanar-degenerate patch layout: at least one axis of
            # rotation is unresisted. pinv gives the least-squares answer; the
            # component of M_C in the null space is genuinely NOT carried.
            J_inv = np.linalg.pinv(J, rcond=SINGULAR_TOL)
            resisted = J @ (J_inv @ M_C)
            unresisted = float(np.linalg.norm(M_C - resisted))
            if unresisted > 1e-6 * max(np.linalg.norm(M_C), 1.0):
                flags.append(
                    f"degenerate_patch_layout: {unresisted:.0f} lbf-in of member-end "
                    f"moment has no patch couple to resist it"
                )
        else:
            J_inv = np.linalg.inv(J)

        theta = J_inv @ M_C
        f_moment = A[:, None] * np.cross(np.broadcast_to(theta, r.shape), r)

    f = f_direct + f_moment
    mags = np.linalg.norm(f, axis=1)
    total_mag = float(mags.sum()) or 1.0

    out: list[PatchDemand] = []
    for i, cid in enumerate(group.connection_ids):
        along = float(f[i] @ n[i])                     # + = member pushes into the
                                                       #     mating part = bearing
        fn = -along                                    # so + normal_lbf = tension
        fs = float(np.linalg.norm(f[i] - along * n[i]))  # in-plane resultant
        out.append(
            PatchDemand(
                connection_id=cid,
                member=group.member,
                joint_id=group.joint_id,
                force=tuple(f[i]),
                normal_lbf=fn,
                shear_lbf=fs,
                moment_share=float(mags[i] / total_mag),
                combo_id=combo_id,
                needs_review=list(flags)
                + (
                    ["multi_patch_distribution: demand split across "
                     f"{len(A)} patches by contact area"]
                    if len(A) > 1
                    else []
                ),
            )
        )
    return out


def severity(d: PatchDemand) -> float:
    """Ranking scalar for enveloping. Deliberately crude — it only has to order
    candidate demands on the SAME patch, not stand in for a capacity check."""
    return abs(d.normal_lbf) + abs(d.shear_lbf)


def envelope(demands: Iterable[PatchDemand]) -> dict[str, PatchDemand]:
    """
    Each connection gets a demand from member_a's group and member_b's group, for
    every load combination. Equilibrium says the two member groups should match;
    the two distributions are independent approximations, so they will not
    exactly. Keep the worst across both members AND all combos — the surviving
    demand carries the ``combo_id`` that governed.
    """
    best: dict[str, PatchDemand] = {}
    for d in demands:
        cur = best.get(d.connection_id)
        if cur is None or severity(d) > severity(cur):
            best[d.connection_id] = d
    return best


def distribute_all(
    connections: Sequence,
    joints_by_id: dict[str, dict],
    member_end_actions: dict[tuple[str, str], tuple[Sequence[float], Sequence[float]]],
    combo_id: str = "",
) -> tuple[dict[str, PatchDemand], list[str]]:
    """
    Full sweep for ONE load combination.

    member_end_actions : {(joint_id, member_occurrence_path): (F_xyz, M_xyz)}
                         the action each member DELIVERS into its joint, lbf and
                         lbf-in, global axes. See ``distribute`` on the sign.

    Returns ({connection_id: worst-case PatchDemand}, review_flags).
    """
    demands, review = _sweep(connections, joints_by_id, member_end_actions, combo_id)
    return envelope(demands), review


def distribute_all_combos(
    connections: Sequence,
    joints_by_id: dict[str, dict],
    actions_by_combo: dict[str, dict[tuple[str, str], tuple[Sequence[float], Sequence[float]]]],
) -> tuple[dict[str, PatchDemand], dict[str, list[PatchDemand]], list[str]]:
    """
    Sweep every load combination and envelope across all of them.

    actions_by_combo : {combo_id: {(joint_id, member): (F_xyz, M_xyz)}}

    Returns (governing demand per connection, every demand per connection,
    review_flags). The per-combo detail is kept so the UI can show WHY a
    connection governs, not just that it does.
    """
    review: list[str] = []
    per_conn: dict[str, list[PatchDemand]] = {}
    for combo, actions in actions_by_combo.items():
        demands, flags = _sweep(connections, joints_by_id, actions, combo)
        for d in demands:
            per_conn.setdefault(d.connection_id, []).append(d)
        # The same structural gap (a group with no FE action) repeats in every
        # combo. Report it once, not once per combo.
        for f in flags:
            if f not in review:
                review.append(f)
    governing = {cid: max(ds, key=severity) for cid, ds in per_conn.items() if ds}
    return governing, per_conn, review


def _sweep(
    connections: Sequence,
    joints_by_id: dict[str, dict],
    member_end_actions: dict[tuple[str, str], tuple[Sequence[float], Sequence[float]]],
    combo_id: str,
) -> tuple[list[PatchDemand], list[str]]:
    review: list[str] = []
    groups = build_patch_groups(connections)
    all_d: list[PatchDemand] = []

    for key, g in groups.items():
        act = member_end_actions.get(key)
        if act is None:
            review.append(
                f"no_fe_action[{key[0]}/{key[1]}]: patch group has no member-end "
                f"force from the solver; its connections are unchecked"
            )
            continue
        F, M = act
        j = joints_by_id.get(g.joint_id)
        if j is None:
            review.append(f"unknown_joint[{g.joint_id}]: patch group references a "
                          f"joint that is not in the model")
            continue
        loc = j["location"] if isinstance(j, dict) else j.location
        all_d.extend(distribute(g, F, M, loc, combo_id))

    return all_d, review


def unresolvable_connections(connections: Sequence) -> dict[str, str]:
    """Connections that can never receive a patch demand, and why.

    These are the rows that must land in the "Unchecked" bucket rather than be
    given an invented number: a connection with no joint has no FE node to draw
    from, and a centerline connection has no contact plane to resolve a force
    into. Both are real states of the Block 1 output, not errors.
    """
    def _g(c, k, d=None):
        return c.get(k, d) if isinstance(c, dict) else getattr(c, k, d)

    out: dict[str, str] = {}
    for c in connections:
        cid = str(_g(c, "connection_id") or "")
        if not cid:
            continue
        if not _g(c, "joint_id"):
            out[cid] = "no joint_id: not attached to an FE node"
        elif _g(c, "contact_normal") is None:
            out[cid] = (f"{_g(c, 'detection_method') or 'centerline'} detection: no "
                        f"contact plane, so member-end force cannot be resolved "
                        f"into normal/shear on a patch")
        else:
            area = _g(c, "contact_area_in2")
            if not area or float(area) <= 0:
                out[cid] = "zero contact area: no patch to carry load"
    return out
