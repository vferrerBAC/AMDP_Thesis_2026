"""
support_axis.py - vertical-axis-correct support selection.

Block 1 (frozen) flags `is_support_candidate` from the lowest-Z joints, but the
SAC BB2 master model is authored Y-up (Inventor default; gravity acts -Y),
matching block3_solver's `vertical_axis = "Y"`. On a Y-up model the Z-based
flag lands the fixed base on a depth/side face instead of the ground, which
invalidates the reactions and all downstream demand.

This module re-derives the support set on the correct vertical axis directly
from the existing Block 1 JSON, leaving the frozen extraction untouched. The
returned node list is passed explicitly into Block 3, which honours an explicit
`support_nodes` list over `is_support_candidate`.
"""

from typing import Any, Dict, List

_AXIS_INDEX = {"X": 0, "Y": 1, "Z": 2}

# Matches DEFAULT_SUPPORT_Z_TOL_IN in block1.py so the band width is unchanged;
# only the axis it is applied to changes.
DEFAULT_SUPPORT_TOL_IN = 1.0


def select_support_nodes(
    block1_data: Dict[str, Any],
    *,
    vertical_axis: str = "Y",
    tol_in: float = DEFAULT_SUPPORT_TOL_IN,
) -> Dict[str, Any]:
    """Re-flag base/ground joints as the lowest ones along the TRUE vertical axis.

    Mirrors Block 1's rule (lowest coordinate + inch tolerance) but on the
    correct axis index. Returns a small dict so the result can be surfaced and
    confirmed the same way Block 1's candidates are:

        support_nodes : sorted list[str]  -> pass to run_lrfd_joint_analysis(...)
        vertical_axis : "X"|"Y"|"Z"
        axis_min      : lowest coordinate on that axis (source length units)
        tol_in        : tolerance band used
        n_candidates  : how many joints fell in the band
        needs_review  : True if the set looks unable to stabilise a 3D frame
    """
    axis = vertical_axis.strip().upper()
    idx = _AXIS_INDEX.get(axis)
    if idx is None:
        raise ValueError(f"vertical_axis must be X, Y or Z; got {vertical_axis!r}")

    joints = [j for j in block1_data.get("joints", []) if j.get("location")]
    if not joints:
        return {"support_nodes": [], "vertical_axis": axis, "axis_min": None,
                "tol_in": tol_in, "n_candidates": 0, "needs_review": True}

    vmin = min(j["location"][idx] for j in joints)
    nodes: List[str] = [j["joint_id"] for j in joints
                        if j["location"][idx] <= vmin + tol_in]

    # Conservative: fewer than 3 base nodes cannot restrain a general 3D frame,
    # so route to confirmation rather than silently solving an unstable model.
    needs_review = len(nodes) < 3
    return {"support_nodes": sorted(nodes), "vertical_axis": axis,
            "axis_min": vmin, "tol_in": tol_in,
            "n_candidates": len(nodes), "needs_review": needs_review}
