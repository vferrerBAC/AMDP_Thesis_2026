"""Golden tests for the joint-analysis -> connection-check bridge.

Covers the three things the re-key was for:
  1. the FE solve is per JOINT, the checks are per CONNECTION, and the demand
     gets from one to the other by equilibrium (not by copying);
  2. a connection that cannot get a demand says so instead of being handed one;
  3. a joint with several contact patches keeps all of them.
"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))

import pandas as pd

from engine.block3_solver import analyze_joints
from engine.bolted_capacity import default_bolted_row
from engine.connection_demand import (
    actions_by_combo, connection_demands, demand_coverage, member_alias_map,
)
from engine.connection_templates import connection_keys, connections_in_joint
from engine.joint_checks import (
    critical_joint_summary, ensure_connection_columns, evaluate_connection_demands,
)
from engine.joint_recommendation import build_joint_verdicts
from engine.patch_groups import build_patch_groups, unresolvable_connections

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(f"{'PASS' if cond else 'FAIL'}  {name}" + (f"   {detail}" if detail and not cond else ""))


SEC = {"section_type": "HSS", "depth": 4.0, "width": 4.0, "wall_thickness": 0.25,
       "A": 3.59, "Iy": 8.22, "Iz": 8.22, "J": 13.4}


def model():
    """Column + beam meeting at J2 through TWO cleats straddling the work point,
    plus one centerline-only connection that must stay unchecked."""
    return {
        "members": [
            {"occurrence_name": "COL", "occurrence_path": "ASM:1/COL:1",
             "start_point": [0, 0, 0], "end_point": [0, 48, 0],
             "cross_section": dict(SEC), "dry_mass": 40.0},
            {"occurrence_name": "BEAM", "occurrence_path": "ASM:1/BEAM:1",
             "start_point": [0, 48, 0], "end_point": [60, 48, 0],
             "cross_section": dict(SEC), "dry_mass": 50.0},
        ],
        "joints": [
            {"joint_id": "J1", "location": [0, 0, 0], "member_names": ["COL"],
             "is_support_candidate": True},
            {"joint_id": "J2", "location": [0, 48, 0],
             "member_names": ["COL", "BEAM"], "is_support_candidate": False},
            {"joint_id": "J3", "location": [60, 48, 0], "member_names": ["BEAM"],
             "is_support_candidate": True},
        ],
        "connections": [
            {"connection_id": "C001", "member_a": "COL", "member_b": "BEAM",
             "joint_id": "J2", "location": [0, 46.0, 0], "contact_area_in2": 2.0,
             "contact_normal": [1.0, 0.0, 0.0], "weld_length_in": 6.0,
             "hole_count": 2, "detection_method": "face_contact"},
            {"connection_id": "C002", "member_a": "COL", "member_b": "BEAM",
             "joint_id": "J2", "location": [0, 50.0, 0], "contact_area_in2": 2.0,
             "contact_normal": [1.0, 0.0, 0.0], "weld_length_in": 6.0,
             "hole_count": 2, "detection_method": "face_contact"},
            {"connection_id": "C003", "member_a": "COL", "member_b": "BEAM",
             "joint_id": "J2", "location": [0, 48.0, 0], "contact_area_in2": None,
             "contact_normal": None, "detection_method": "centerline"},
        ],
    }


def solve(b1):
    return analyze_joints(b1, config={"source_mass_unit": "lb"},
                          loads={"wind": [{"node": "J2", "FX": 500.0}],
                                 "seismic_h": []})


def schedule(b1):
    s = ensure_connection_columns(pd.DataFrame(
        [default_bolted_row(cid, jid) for cid, jid in connection_keys(b1)]))
    s["connection_type"] = "bolted_bracket_joint"
    s["sheet_t_in"] = 0.1046
    s["n_fasteners"] = 2
    return s


b1 = model()
res = solve(b1)
table, detail, review = connection_demands(b1, res)
by_id = {r["connection_id"]: r for _, r in table.iterrows()}


# --- 1. the solver hands over GLOBAL actions, not member-local components ----
print("\n== solver handoff ==")
check("solver emits joint_member_actions", "joint_member_actions" in res)
acts = actions_by_combo(res)
check("actions keyed by combo then (joint, member)",
      ("J2", "COL") in acts["W1+"] and ("J2", "BEAM") in acts["W1+"])

# Newton's third law at the node: the two members deliver equal-and-opposite
# horizontal action into J2 apart from the 500 lbf applied there.
fx_sum = sum(acts["W1+"][k][0][0] for k in acts["W1+"] if k[0] == "J2")
check("member actions at J2 balance the applied 500 lbf",
      abs(fx_sum + 500.0) < 1.0, f"sum Fx = {fx_sum:.2f}, expected -500")


# --- 2. demand reaches the patches by equilibrium ----------------------------
print("\n== demand distribution ==")
check("both face-contact patches got a demand",
      by_id["C001"]["status"] == "Checked" and by_id["C002"]["status"] == "Checked")

# Two equal patches straddling the work point: the direct share is half the
# member-end force each, and the member-end moment is carried as a COUPLE
# (+d on one patch, -d on the other), not split by area. So the two normal
# demands must differ, and average to the direct share.
n1 = by_id["C001"]["tension_lbf"] - by_id["C001"]["bearing_lbf"]
n2 = by_id["C002"]["tension_lbf"] - by_id["C002"]["bearing_lbf"]
check("patch demands differ -> moment carried as a couple, not area-split",
      abs(n1 - n2) > 1.0, f"n1={n1}, n2={n2}")

# The two patches are enveloped over ALL combos, so each keeps its own worst
# case and they can govern under different (reversed) wind directions.
check("governing combo is recorded per connection",
      bool(by_id["C001"]["combo_id"]) and bool(by_id["C002"]["combo_id"]))
check("reversed wind governs the two patches oppositely",
      by_id["C001"]["combo_id"] != by_id["C002"]["combo_id"],
      f"{by_id['C001']['combo_id']} vs {by_id['C002']['combo_id']}")

# Sign convention: with the normal oriented out of each group's own member, a
# positive normal is tension. One patch must be in tension and one in bearing
# under a moment couple.
check("one patch in tension, the other in bearing",
      (n1 > 0) != (n2 > 0), f"n1={n1}, n2={n2}")


# --- 3. normals are oriented per member --------------------------------------
print("\n== normal orientation ==")
groups = build_patch_groups(b1["connections"])
na = groups[("J2", "COL")].normals[0]
nb = groups[("J2", "BEAM")].normals[0]
check("member_a and member_b groups get opposed normals",
      abs(float(na @ nb) + 1.0) < 1e-9, f"na.nb = {float(na @ nb)}")


# --- 4. unresolvable connections fail loud -----------------------------------
print("\n== unchecked, not invented ==")
check("centerline connection is flagged unresolvable",
      "C003" in unresolvable_connections(b1["connections"]))
check("centerline connection reports Unchecked",
      by_id["C003"]["status"] == "Unchecked")
check("Unchecked carries no demand number",
      pd.isna(by_id["C003"]["force_demand_lbf"]))
check("Unchecked states a reason", bool(by_id["C003"]["review_reason"]))
check("coverage counts the gap",
      demand_coverage(table) == {"total": 3, "checked": 2, "unchecked": 1},
      str(demand_coverage(table)))

# The whole point: it must NOT have been handed the joint's total force.
check("Unchecked did not inherit a sibling patch's demand",
      by_id["C003"]["force_demand_lbf"] is None
      or pd.isna(by_id["C003"]["force_demand_lbf"]))


# --- 5. checks are per connection, rollup is reporting only ------------------
print("\n== check and rollup ==")
sched = schedule(b1)
results = evaluate_connection_demands(table, sched)
check("one result row per connection", len(results) == 3, f"got {len(results)}")
check("Unchecked demand stays Unchecked after the capacity check",
      results.set_index("connection_id").loc["C003", "status"] == "Unchecked")

rollup = critical_joint_summary(results)
j2 = rollup.set_index("joint_id").loc["J2"]
check("rollup counts all 3 patches", int(j2["n_connections"]) == 3)
check("rollup counts the unchecked one", int(j2["n_unchecked"]) == 1)
check("joint takes its WORST connection's status", j2["worst_status"] == "Unchecked")
check("rollup names the governing connection", bool(j2["governing_connection"]))


# --- 6. the schedule is keyed by connection ----------------------------------
print("\n== schedule keying ==")
check("schedule has one row per connection", len(sched) == 3)
check("connection_id is the key", sched["connection_id"].is_unique)
check("joint_id rides along for rollup",
      set(sched["joint_id"]) == {"J2"})
check("connections_in_joint finds every patch",
      sorted(connections_in_joint(sched, "J2")) == ["C001", "C002", "C003"])

# Legacy joint-keyed CSVs must still load rather than being rejected outright.
legacy = ensure_connection_columns(pd.DataFrame([{"joint_id": "J2", "n_fasteners": 2}]))
check("legacy joint-keyed row gets a derived connection_id",
      legacy.iloc[0]["connection_id"] == "J2:1")


# --- 7. the last-wins bug is gone --------------------------------------------
print("\n== multi-patch joint no longer collapses ==")
verdicts = {v.joint_id: v for v in build_joint_verdicts(
    b1, critical_summary=[{"joint_id": "J2", "demand_lbf": 300.0,
                           "governing_combo": "W1+", "worst_status": "OK"}],
    connection_schedule=sched)}
check("all 3 connections reached the verdict (was 1 before)",
      verdicts["J2"].n_connections == 3,
      f"n_connections = {verdicts['J2'].n_connections}")
check("verdict names which connection governed",
      bool(verdicts["J2"].governing_connection_id))
check("multi-patch joint is flagged as such",
      any("contact patches" in r for r in verdicts["J2"].review_reasons))

# A joint whose patches differ must be governed by the WORSE one, not the last.
mixed = sched.copy()
mixed.loc[mixed["connection_id"] == "C002", "n_fasteners"] = 12   # much stronger
mixed.loc[mixed["connection_id"] == "C001", "n_fasteners"] = 1    # much weaker
v = {x.joint_id: x for x in build_joint_verdicts(
    b1, critical_summary=[{"joint_id": "J2", "demand_lbf": 300.0,
                           "governing_combo": "W1+", "worst_status": "OK"}],
    connection_schedule=mixed)}["J2"]
check("the WEAKEST patch governs the joint, not the last row",
      v.governing_connection_id == "C001", f"got {v.governing_connection_id!r}")


# --- 8. name aliasing ---------------------------------------------------------
print("\n== member alias resolution ==")
alias = member_alias_map(b1)
check("occurrence_path resolves to the solver's key",
      alias.get("ASM:1/COL:1") == "COL", str(alias))

aliased = model()
for c in aliased["connections"]:
    c["member_a"] = "ASM:1/COL:1"
    c["member_b"] = "ASM:1/BEAM:1"
t2, _, _ = connection_demands(aliased, solve(aliased))
check("path-named connections still get demand (no silent Unchecked)",
      demand_coverage(t2)["checked"] == 2, str(demand_coverage(t2)))


# --- 9. no FE solve at all ----------------------------------------------------
print("\n== no solve ==")
t3, _, r3 = connection_demands(b1, {"status": "not run"})
check("everything Unchecked when the solve has not run",
      (t3["status"] == "Unchecked").all())
check("and it says why", any("no_fe_actions" in f for f in r3))


print("\n" + "=" * 60)
print(f"{len(PASS)} passed, {len(FAIL)} failed")
if FAIL:
    for f in FAIL:
        print("  FAILED:", f)
    sys.exit(1)
print("ALL PASSED")
