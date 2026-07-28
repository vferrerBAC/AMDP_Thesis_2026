"""Tests for the joint recommendation integration layer (Milestone 4).

Covers the three things that can silently go wrong in an integration layer:
  1. the JOIN CHAIN (joint -> members -> parts) resolving to the wrong records,
  2. the RECOMMENDATION RULE preferring a cheap option that cannot carry load,
  3. MISSING INPUTS producing a confident answer instead of a review flag.

Run:  pytest test_joint_recommendation.py   (or  python test_joint_recommendation.py)
"""

import math

from joint_recommendation import (
    build_joint_verdicts,
    evaluate_joint,
    index_members,
    recommend,
    summarize_verdicts,
    ConnectionOption,
    _branch_and_chord,
    _resolve_part_key,
)


# --- Fixtures ---------------------------------------------------------------

def _members():
    return [
        {"occurrence_path": "CHORD:1", "part_number": "HSS4X4X12",
         "material": "Carbon Steel Tube",
         "cross_section": {"width": 4.0, "depth": 4.0, "wall_thickness": 0.127, "gauge": 10}},
        {"occurrence_path": "BRANCH:1", "part_number": "HSS4X4X10",
         "material": "Carbon Steel Tube",
         "cross_section": {"width": 4.0, "depth": 4.0, "wall_thickness": 0.095, "gauge": 12}},
    ]


def _joint(**over):
    j = {
        "joint_id": "J001",
        "geom_descriptor": "T",
        "joint_type": "T",
        "member_names": ["CHORD:1", "BRANCH:1"],
        "member_roles": {"CHORD:1": "chord", "BRANCH:1": "branch"},
    }
    j.update(over)
    return j


def _block1(**over):
    b = {"members": _members(), "joints": [_joint()]}
    b.update(over)
    return b


def _critical(demand=9000.0):
    return [{"joint_id": "J001", "demand_lbf": demand,
             "governing_combo": "1.2D+1.6L", "worst_status": "OK"}]


def _schedule(n=4, **over):
    row = {"joint_id": "J001", "connection_type": "bolted_bracket_joint",
           "material_grade": "Carbon Steel Tube", "sheet_t_in": 0.095,
           "n_fasteners": n, "diameter_in": 0.625, "edge_dist_in": 1.25,
           "fastener_type": "ASTM F3125, Grade A325, Type 1"}
    row.update(over)
    return [row]


def _mfg(tl=True):
    return [
        {"part_identifier": "HSS4X4X12",
         "selected_process": "Tube Laser" if tl else "Manual Panel Bender",
         "tube_laser_ok": tl},
        {"part_identifier": "HSS4X4X10",
         "selected_process": "Tube Laser" if tl else "Manual Panel Bender",
         "tube_laser_ok": tl},
    ]


# --- 1. The join chain ------------------------------------------------------

def test_member_index_resolves_every_alias():
    idx = index_members(_members())
    # a joint may name a member by path OR part number; both must resolve
    assert idx["CHORD:1"]["part_number"] == "HSS4X4X12"
    assert idx["HSS4X4X12"]["occurrence_path"] == "CHORD:1"


def test_part_key_precedence_matches_manufacturability():
    """manufacturability._part_identifier prefers part_number over occurrence_path."""
    assert _resolve_part_key({"part_number": "P1", "occurrence_path": "O1"}) == "P1"
    assert _resolve_part_key({"occurrence_path": "O1"}) == "O1"
    assert _resolve_part_key({}) == ""


def test_branch_and_chord_read_from_block1_roles():
    branch, chord = _branch_and_chord(_joint())
    assert branch == "BRANCH:1"
    assert chord == "CHORD:1"


def test_branch_and_chord_handles_primary_crossing_roles():
    """Block 1 uses primary/crossing for non-hollow configurations."""
    j = _joint(member_roles={"CHORD:1": "primary", "BRANCH:1": "crossing"})
    branch, chord = _branch_and_chord(j)
    assert branch == "BRANCH:1" and chord == "CHORD:1"


def test_manufacturability_joins_through_part_number():
    """Joint names members by occurrence_path; mfg files them by part_number.
    If the chain breaks, member_processes comes back empty."""
    v = build_joint_verdicts(_block1(), _critical(), _schedule(), _mfg())[0]
    assert v.member_processes == {"CHORD:1": "Tube Laser", "BRANCH:1": "Tube Laser"}
    assert v.tube_laser_ok is True


# --- 2. The recommendation rule --------------------------------------------

def _opt(cap, cost, demand):
    o = ConnectionOption("x", capacity_lbf=cap, cost_usd=cost)
    o.capacity_per_dollar = cap / cost
    o.utilization = demand / cap
    o.percent_used = o.utilization * 100
    o.adequate = cap >= demand
    return o


def test_inadequate_option_never_recommended_however_cheap():
    """The whole point of the gate: a $1 connection that fails is not a bargain."""
    opts = {
        "bolted": _opt(cap=1_000, cost=1.0, demand=9_000),    # 1000 lbf/$ but FAILS
        "welded": _opt(cap=20_000, cost=50.0, demand=9_000),  # 400 lbf/$ and passes
    }
    best, why = recommend(opts, demand_lbf=9_000)
    assert best == "welded"
    assert "only option" in why


def test_among_adequate_options_best_capacity_per_dollar_wins():
    opts = {
        "bolted": _opt(cap=10_000, cost=5.0, demand=9_000),   # 2000 lbf/$
        "welded": _opt(cap=20_000, cost=20.0, demand=9_000),  # 1000 lbf/$
    }
    best, why = recommend(opts, demand_lbf=9_000)
    assert best == "bolted"
    assert "cost decides" in why
    assert "2.0x" in why


def test_no_adequate_option_recommends_nothing():
    opts = {
        "bolted": _opt(cap=1_000, cost=5.0, demand=9_000),
        "welded": _opt(cap=2_000, cost=8.0, demand=9_000),
    }
    best, why = recommend(opts, demand_lbf=9_000)
    assert best is None
    assert "Neither option carries" in why


def test_unknown_demand_flags_that_choice_is_not_demand_checked():
    opts = {"bolted": _opt(10_000, 5.0, 1), "welded": _opt(20_000, 20.0, 1)}
    best, why = recommend(opts, demand_lbf=None)
    assert best == "bolted"
    assert "Not checked against demand" in why


def test_no_computable_option_recommends_nothing():
    best, why = recommend({}, demand_lbf=9_000)
    assert best is None


# --- 3. Capacity-per-dollar consistency ------------------------------------

def test_weld_length_drives_both_capacity_and_cost():
    """The link that makes capacity-per-dollar meaningful: the weld length used
    for cost must be the same one the capacity template computed (col P = F+G)."""
    from weld_capacity import weld_capacity
    from joining_cost import welding_cost

    v = build_joint_verdicts(_block1(), _critical(), _schedule(), _mfg())[0]
    welded = v.options["welded"]

    cap = weld_capacity({
        "branch_material": "Carbon Steel Tube", "branch_t_in": 0.095,
        "branch_width_in": 4.0, "branch_height_in": 4.0,
        "chord_material": "Carbon Steel Tube", "chord_t_in": 0.127,
        "fastener_type": "E70XX",
    })
    assert cap.weld_length_in == 8.0                      # col P = 4 + 4
    expected_cost = welding_cost(8.0).cost_usd
    assert math.isclose(welded.cost_usd, expected_cost, rel_tol=1e-9)
    assert math.isclose(welded.capacity_per_dollar,
                        cap.shear_lbf / expected_cost, rel_tol=1e-9)


def test_capacity_matches_golden_j001():
    """Cross-check: this T-joint is golden case J001 from the welded template."""
    v = build_joint_verdicts(_block1(), _critical(), _schedule(), _mfg())[0]
    assert math.isclose(v.options["welded"].capacity_lbf, 22270.5, rel_tol=1e-6)


def test_bolt_count_drives_both_capacity_and_cost():
    v4 = build_joint_verdicts(_block1(), _critical(), _schedule(n=4), _mfg())[0]
    v8 = build_joint_verdicts(_block1(), _critical(), _schedule(n=8), _mfg())[0]
    # Capacity is a group total -> doubles. Cost rises too (but not linearly:
    # material scales with n, labor scales with n).
    assert math.isclose(v8.options["bolted"].capacity_lbf,
                        2 * v4.options["bolted"].capacity_lbf, rel_tol=1e-9)
    assert v8.options["bolted"].cost_usd > v4.options["bolted"].cost_usd


# --- 4. Missing inputs degrade to review, never to a confident answer -------

def test_missing_demand_routes_to_review():
    v = build_joint_verdicts(_block1(), None, _schedule(), _mfg())[0]
    assert v.needs_review is True
    assert v.demand_lbf is None
    assert any("Stage 1" in r for r in v.review_reasons)


def test_missing_connection_schedule_gives_welded_only():
    v = build_joint_verdicts(_block1(), _critical(), None, _mfg())[0]
    assert "welded" in v.options
    assert "bolted" not in v.options
    assert v.needs_review is True


def test_missing_member_roles_still_resolves_a_pair():
    j = _joint(member_roles={})
    v = build_joint_verdicts({"members": _members(), "joints": [j]},
                             _critical(), _schedule(), _mfg())[0]
    assert "welded" in v.options       # fell back to positional order
    assert v.branch_name and v.chord_name
    assert v.branch_name != v.chord_name


def test_unresolvable_member_flags_review():
    j = _joint(member_names=["CHORD:1", "GHOST:9"],
               member_roles={"CHORD:1": "chord", "GHOST:9": "branch"})
    v = build_joint_verdicts({"members": _members(), "joints": [j]},
                             _critical(), _schedule(), _mfg())[0]
    assert v.needs_review is True
    assert any("not found" in r for r in v.review_reasons)


def test_block1_review_flag_propagates():
    j = _joint(needs_review=True, review_reason="section ambiguous")
    v = build_joint_verdicts({"members": _members(), "joints": [j]},
                             _critical(), _schedule(), _mfg())[0]
    assert v.needs_review is True
    assert any("section ambiguous" in r for r in v.review_reasons)


# --- 5. Verdicts and manufacturability -------------------------------------

def test_not_tube_laser_is_reported_but_not_fatal():
    v = build_joint_verdicts(_block1(), _critical(), _schedule(), _mfg(tl=False))[0]
    assert v.tube_laser_ok is False
    assert v.verdict != "NOT MANUFACTURABLE"     # MPB still works
    assert any("Not fully tube-laser" in a for a in v.actions)


def test_no_process_available_is_not_manufacturable():
    mfg = [{"part_identifier": "HSS4X4X12", "selected_process": None, "tube_laser_ok": False},
           {"part_identifier": "HSS4X4X10", "selected_process": None, "tube_laser_ok": False}]
    v = build_joint_verdicts(_block1(), _critical(), _schedule(), mfg)[0]
    assert v.verdict == "NOT MANUFACTURABLE"


def test_no_adequate_option_verdict():
    """Demand far beyond both options."""
    v = build_joint_verdicts(_block1(), _critical(demand=5_000_000),
                             _schedule(), _mfg())[0]
    assert v.verdict == "NO ADEQUATE OPTION"
    assert v.recommended is None


def test_verdicts_sort_worst_first():
    b = _block1()
    b["joints"] = [_joint(joint_id="J_OK"), _joint(joint_id="J_BAD")]
    crit = [{"joint_id": "J_OK", "demand_lbf": 9000.0, "governing_combo": "c"},
            {"joint_id": "J_BAD", "demand_lbf": 5_000_000.0, "governing_combo": "c"}]
    sched = _schedule() + [dict(_schedule()[0], joint_id="J_BAD")]
    sched[0]["joint_id"] = "J_OK"
    vs = build_joint_verdicts(b, crit, sched, _mfg())
    assert vs[0].joint_id == "J_BAD"      # worst first


def test_verdict_is_json_serializable():
    v = build_joint_verdicts(_block1(), _critical(), _schedule(), _mfg())[0]
    import json
    json.dumps(v.to_dict())               # must not raise


def test_summarize_verdicts():
    vs = build_joint_verdicts(_block1(), _critical(), _schedule(), _mfg())
    s = summarize_verdicts(vs)
    assert s["joints"] == 1
    assert s["tube_laser_ok"] == 1
    assert s["recommend_bolted"] + s["recommend_welded"] == 1


def test_empty_model_does_not_crash():
    assert build_joint_verdicts({"members": [], "joints": []}) == []


if __name__ == "__main__":
    import sys
    failed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"PASS  {name}")
            except AssertionError as exc:
                failed += 1
                print(f"FAIL  {name}: {exc}")
    print(f"\n{'ALL PASSED' if not failed else f'{failed} FAILED'}")
    sys.exit(1 if failed else 0)
