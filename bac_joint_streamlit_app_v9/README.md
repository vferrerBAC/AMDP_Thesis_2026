# BAC Joint Check Assistant — Streamlit Prototype v2

This version is designed for design engineers with limited structural/civil background. It uses a guided workflow, Simple/Advanced modes, connection templates, validation checks, traffic-light results, suggested fixes, and professional exports.

## Run locally

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Or on Windows, double-click `run_local.bat`.

## Main workflow

1. Project setup
2. Upload or load sample CAD-derived joint geometry
3. Extract connections (real face contacts)
4. Fill in the bolted connection schedule — one row per contact patch
5. Run the structural FE solve (loads and combinations)
6. Validate inputs
7. Run connection screening checks
8. Export Excel, PDF, CSV templates, and JSON config

## Analysis by joint, checks by connection

The two levels are deliberate and different:

- **Analysis is per JOINT.** `engine/block3_solver.py` builds a centerline frame
  with one FE node per joint. A contact patch has no stiffness that can be
  defensibly assigned, so it cannot be a node.
- **Checks are per CONNECTION.** Every quantity a capacity equation needs —
  contact area, patch perimeter, hole count, thickness, bolt pattern — lives on
  the connection. A joint has a location and a member list; there is nothing
  there to check.

`engine/connection_demand.py` is the bridge. It takes the solver's per-member
global end actions and hands them to `engine/patch_groups.py`, which distributes
each one across the patches that physically carry it (area-weighted direct share
plus a rotational couple), per load combination, then envelopes.

A connection that cannot receive a demand — centerline-only detection, no contact
plane, or no member-end action from the solver — reports **Unchecked** with a
reason. It is never given a substitute number: loading every patch with the full
member-end force would oversize a four-patch joint by ~4x and destroy the tool's
ability to tell a good connection from a bad one.

Results roll up to the joint for reporting only. A joint's status is its worst
connection's status.

## Engineering limitation

The current joint capacities are **demo screening placeholders**. They are not final AISC/AISI/AWS code equations. Replace the placeholder formulas in `engine/joint_checks.py` with approved BAC/company/code equations before design release.

## Privacy recommendation

Start local-only. Uploaded files stay in the running Streamlit session unless the code is modified to save them. Do not deploy real BAC/customer CAD data to a cloud service without an approved security, authentication, storage, and deletion plan.
