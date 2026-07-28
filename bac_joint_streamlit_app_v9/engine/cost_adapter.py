"""engine/cost_adapter.py — Block 1 extraction -> cost.py contract (pure transform).

Block 1 (with the cost extension) now emits the nine finished cost fields for
BOTH part classes, so this adapter derives nothing: it renames Block 1 keys into
cost.py's contract, derives part_set/quantity, and flags anything missing.
"""
from __future__ import annotations
import os
from collections import Counter

# logical_name -> Block 1 key (all nine cost fields are emitted by block1_cost_extract)
FIELD_MAP = {
    "part_identifier":     "part_number",
    "part_class":          "part_class",
    "ncx_material":        "NCx_Material",
    "gauge":               "Gauge",
    "assembly_category":   "CostDataAssemblyCategory",
    "pierce_count":        "CostDataPierceCount",
    "cut_distance_inches": "CostDataCutDistanceInches",
    "unique_bends":        "CostDataUniqueBends",
    "corner_weld":         "Corner Weld",
    "flat_length_inches":  "CostDataFlatLengthInches",
    "flat_width_inches":   "CostDataFlatWidthInches",
    # joints
    "joint_part_a":        "part_a",
    "joint_part_b":        "part_b",
    "joint_length_inches": "joint_length_inches",
}
TUBE_FAMILIES = {"closed", "hss", "tube", "rect_tube", "round_tube"}


def _src(rec, logical, default=None):
    return rec.get(FIELD_MAP[logical], default)


def _flag(nr, scope, ident, field, issue):
    nr.append({"scope": scope, "identifier": ident, "field": field, "issue": issue})


def _classify(rec):
    pc = _src(rec, "part_class")
    if isinstance(pc, str) and pc.strip():
        return "tube" if pc.strip().lower() == "tube" else "sheet_metal"
    fam = (rec.get("cross_section", {}) or {}).get("family")   # fallback for pre-extension JSON
    if isinstance(fam, str) and fam.strip():
        return "tube" if fam.strip().lower() in TUBE_FAMILIES else "sheet_metal"
    return "sheet_metal"


def part_to_cost_input(rec, nr):
    rec = {**rec, **(rec.get("cost") or {})}   # lift Block 1 nested cost.* to top level
    ident = _src(rec, "part_identifier", "<unknown>")

    def field(logical, label, numeric=True):
        v = rec.get(FIELD_MAP[logical])
        if v is None or (isinstance(v, str) and v.strip() == ""):
            _flag(nr, "part", ident, label, "missing (Block 1 cost extension not run / iProperty absent)")
            return None
        if numeric:
            try:
                return float(v)
            except (TypeError, ValueError):
                _flag(nr, "part", ident, label, f"unparseable numeric: {v!r}")
                return None
        return v

    return {
        "part_set":          rec.get("part_set"),
        "part_identifier":   ident,
        "part_quantity":     rec.get("quantity", 1),
        "part_class":        _classify(rec),
        "cutting_method":    None,   # cost.py applies class-aware default
        "forming_method":    None,   # cost.py applies Manual Press Brake default
        "ncx_material":      field("ncx_material", "ncx_material", numeric=False),
        "gauge":             field("gauge", "gauge"),
        "assembly_category": field("assembly_category", "assembly_category", numeric=False),
        "pierce_count":        field("pierce_count", "pierce_count"),
        "cut_distance_inches": field("cut_distance_inches", "cut_distance_inches"),
        "unique_bends":        field("unique_bends", "unique_bends"),
        "corner_weld":         field("corner_weld", "corner_weld"),
        "flat_length_inches":  field("flat_length_inches", "flat_length_inches"),
        "flat_width_inches":   field("flat_width_inches", "flat_width_inches"),
    }


def joints_to_cost_inputs(block1_data, nr):
    """Cost 'joints' are the CONNECTIONS from Block 1 — each contact patch is a
    part-to-part joint whose fastener count the workbook derives from its length.

    Face-contact connections carry ``joint_length_in`` (the span of the contact
    overlap); centerline connections do not, and are flagged rather than costed.
    Without a length the workbook cannot compute fastener quantity, which is why
    the Summary HARDWARE columns read 0 until face-contact extraction (Step 3)
    has run. ``part_a``/``part_b`` keep the occurrence name so the workbook's
    Part-A-base-name lookup resolves each joint's part set from the BAC Part List.
    """
    rows = []
    for c in (block1_data.get("connections") or []):
        if not isinstance(c, dict):
            continue
        # Respect the engineer's pruning: skip removed / unreachable connections.
        if c.get("manually_removed"):
            continue
        if c.get("pruned_noncontact") and not c.get("prune_override"):
            continue

        cid = c.get("connection_id", "<connection>")
        a, b = c.get("member_a"), c.get("member_b")
        length = c.get("joint_length_in")
        if not a or not b:
            _flag(nr, "joint", cid, "member_a/member_b", "connection is missing a member reference")
            continue
        if length is None:
            _flag(nr, "joint", cid, "joint_length_inches",
                  "centerline connection has no contact length; run face-contact extraction (Step 3)")
            continue
        rows.append({"part_a": a, "part_b": b, "joint_length_inches": length})
    return rows


def build_cost_inputs(block1_data, aggregate_by_identifier=True):
    nr = []
    src = block1_data.get("source_document")
    part_set = os.path.splitext(os.path.basename(src.replace("\\", "/")))[0] if src else None

    members = block1_data.get("members", block1_data.get("parts", []))
    counts = Counter(m.get("part_number") for m in members)

    merged, order, flat = {}, [], []
    for m in members:
        rec = dict(m)
        rec["part_set"] = part_set
        rec["quantity"] = counts.get(rec.get("part_number"), 1)
        p = part_to_cost_input(rec, nr)
        pid = p["part_identifier"]
        if aggregate_by_identifier:
            if pid not in merged:
                merged[pid] = p
                order.append(pid)
        else:
            flat.append(p)
    parts = [merged[pid] for pid in order] if aggregate_by_identifier else flat

    joints = joints_to_cost_inputs(block1_data, nr)
    return parts, joints, nr
