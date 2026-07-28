"""
engine/block1_cost_extract.py  —  Block 1 cost-input extension

Additive to block1.py: for each part occurrence, reads the BAC cost iProperties
(and, for tube parts, measures cut geometry) so the Block 1 JSON carries the nine
cost fields the template needs. This is a faithful Python/COM port of Steve's
CostEstimationV4 routines (BuildSheetMetalData / BuildTubePartData /
MeasureTubeGeometry), with ONE intentional change per IJET's rules:

    Missing/unparseable iProperties -> needs_review + None,
    NOT the macro's silent MsgBox-then-"1" / TryParseDoubleSafe-then-1.0.

INTEGRATION (surgical, additive — inside block1.py's existing occurrence loop):

    from engine.block1_cost_extract import extract_part_cost_fields
    ...
    part_doc = occurrence.Definition.Document          # already how you get it
    cost_fields, cost_review = extract_part_cost_fields(part_doc, member["part_number"])
    member.update(cost_fields)                          # merges the 9 fields + part_class
    all_needs_review.extend(cost_review)

Emitted keys (both paths), matching cost_adapter.FIELD_MAP:
    NCx_Material, Gauge, CostDataAssemblyCategory, CostDataPierceCount,
    CostDataCutDistanceInches, CostDataUniqueBends, Corner Weld,
    CostDataFlatLengthInches, CostDataFlatWidthInches, part_class
    (+ tube_geometry: raw measurements, kept for Milestone 3 tube-laser use)

ENV: Windows + Inventor + pywin32. Tube geometry needs the Inventor dispatch to
expose type-lib constants (EnsureDispatch), OR set PLANE_SURFACE_ENUM to the
numeric SurfaceTypeEnum.kPlaneSurface value. The sheet-metal path uses no
geometry, so it runs without any of that.
"""
from __future__ import annotations
import math

CM_TO_IN = 1.0 / 2.54
IPROP_SET = "User Defined Properties"      # Steve's set name; some installs: "Inventor User Defined Properties"
DESIGN_TRACKING = "Design Tracking Properties"
SHEET_METAL_PROBE = "CostDataFlatLengthInches"

# Resolved lazily from the live Inventor type library; or set explicitly.
PLANE_SURFACE_ENUM = None


def _plane_enum():
    global PLANE_SURFACE_ENUM
    if PLANE_SURFACE_ENUM is None:
        try:
            from win32com.client import constants
            PLANE_SURFACE_ENUM = constants.kPlaneSurface
        except Exception:
            pass
    return PLANE_SURFACE_ENUM


def _iter(coll):
    """Iterate a COM collection (1-indexed) or a plain list uniformly."""
    try:
        return list(coll)
    except TypeError:
        return [coll.Item(i) for i in range(1, coll.Count + 1)]


def _flag(nr, ident, field, issue):
    nr.append({"scope": "part", "identifier": ident, "field": field, "issue": issue})


# ---------------------------------------------------------------------------
# iProperty reads (== GetRequiredCustomiProp / TryGetCustomiProp, no silent "1")
# ---------------------------------------------------------------------------
def _read_iprop(part_doc, name):
    try:
        return part_doc.PropertySets.Item(IPROP_SET).Item(name).Value, True
    except Exception:
        return None, False


def _req_str(part_doc, name, ident, nr):
    raw, present = _read_iprop(part_doc, name)
    if not present or raw in (None, ""):
        _flag(nr, ident, name, "missing required iProperty (not defaulted)")
        return None
    return str(raw)


def _req_num(part_doc, name, ident, nr):
    raw, present = _read_iprop(part_doc, name)
    if not present:
        _flag(nr, ident, name, "missing required iProperty (not defaulted)")
        return None
    try:
        return float(raw)
    except (TypeError, ValueError):
        _flag(nr, ident, name, f"unparseable numeric iProperty: {raw!r} (not defaulted)")
        return None


# ---------------------------------------------------------------------------
# Sheet-metal path (== BuildSheetMetalData: read all nine directly)
# ---------------------------------------------------------------------------
def build_sheet_metal_fields(part_doc, ident, nr):
    return {
        "NCx_Material":              _req_str(part_doc, "NCx_Material", ident, nr),
        "Gauge":                     _req_num(part_doc, "Gauge", ident, nr),
        "CostDataAssemblyCategory":  _req_str(part_doc, "CostDataAssemblyCategory", ident, nr),
        "CostDataPierceCount":       _req_num(part_doc, "CostDataPierceCount", ident, nr),
        "CostDataCutDistanceInches": _req_num(part_doc, "CostDataCutDistanceInches", ident, nr),
        "CostDataUniqueBends":       _req_num(part_doc, "CostDataUniqueBends", ident, nr),
        "Corner Weld":               _req_num(part_doc, "Corner Weld", ident, nr),
        "CostDataFlatLengthInches":  _req_num(part_doc, "CostDataFlatLengthInches", ident, nr),
        "CostDataFlatWidthInches":   _req_num(part_doc, "CostDataFlatWidthInches", ident, nr),
    }


# ---------------------------------------------------------------------------
# Tube path (== BuildTubePartData + MeasureTubeGeometry)
# ---------------------------------------------------------------------------
def build_tube_fields(part_doc, ident, nr):
    try:
        g = measure_tube_geometry(part_doc)
    except Exception as e:
        _flag(nr, ident, "tube_geometry", f"geometry measurement failed: {e!r}")
        g = None

    if g is None:
        flat_len = flat_wid = cut = pierce = None
    else:
        flat_len = g["longest_axis_in"]
        flat_wid = g["cross_section_perimeter_in"]
        ep = g["end_face_perimeter_sum_in"]
        cut = (ep + g["interior_cut_edge_sum_in"]) if ep is not None else None
        pierce = 2 + g["interior_cut_loop_count"]

    fields = {
        "NCx_Material":              _req_str(part_doc, "NCx_Material", ident, nr),
        "Gauge":                     _req_num(part_doc, "Gauge", ident, nr),
        "CostDataAssemblyCategory":  _req_str(part_doc, "CostDataAssemblyCategory", ident, nr),
        "CostDataFlatLengthInches":  flat_len,
        "CostDataFlatWidthInches":   flat_wid,
        "CostDataCutDistanceInches": cut,
        "CostDataPierceCount":       pierce,
        "CostDataUniqueBends":       1,
        "Corner Weld":               1,
    }
    if g is not None:
        fields["tube_geometry"] = g
    return fields


def measure_tube_geometry(part_doc):
    """One face+edge walk. Faithful port of MeasureTubeGeometry (single body)."""
    body = part_doc.ComponentDefinition.SurfaceBodies.Item(1)

    rb = body.RangeBox
    dims = [rb.MaxPoint.X - rb.MinPoint.X,
            rb.MaxPoint.Y - rb.MinPoint.Y,
            rb.MaxPoint.Z - rb.MinPoint.Z]
    axis_idx = max(range(3), key=lambda i: dims[i])
    longest_cm = dims[axis_idx]

    faces = _iter(body.Faces)
    i_min, i_max = _find_end_face_indices(faces, axis_idx)

    xsection = _outer_loop_length_in(faces[i_min]) if i_min is not None else 0.0
    endsum = 0.0
    for i in (i_min, i_max):
        if i is not None:
            endsum += _outer_loop_length_in(faces[i])

    end_set = {i for i in (i_min, i_max) if i is not None}
    interior_edge = 0.0
    interior_count = 0
    for fi, f in enumerate(faces):
        if fi in end_set:
            continue
        for lp in _iter(f.EdgeLoops):
            if not lp.IsOuterEdgeLoop:
                interior_count += 1
                for e in _iter(lp.Edges):
                    interior_edge += _edge_length_cm(e) * CM_TO_IN

    # Each interior loop is visited from both walls it pierces -> halve/ceiling.
    interior_edge /= 2.0
    interior_count = math.ceil(interior_count / 2.0)

    return {
        "longest_axis_in":            longest_cm * CM_TO_IN,
        "cross_section_perimeter_in": xsection,
        "end_face_perimeter_sum_in":  endsum,
        "interior_cut_edge_sum_in":   interior_edge,
        "interior_cut_loop_count":    interior_count,
    }


def _find_end_face_indices(faces, axis_idx):
    """Two planar faces whose centroids are the extremes along axis (miter-safe)."""
    plane = _plane_enum()
    min_c, max_c = float("inf"), float("-inf")
    i_min = i_max = None
    for i, f in enumerate(faces):
        if f.SurfaceType != plane:
            continue
        mid = _face_axis_mid(f, axis_idx)
        if mid is None:
            continue
        if mid < min_c:
            min_c, i_min = mid, i
        if mid > max_c:
            max_c, i_max = mid, i
    return i_min, i_max


def _face_axis_mid(f, axis_idx):
    try:
        rb = f.Evaluator.RangeBox
        lo = (rb.MinPoint.X, rb.MinPoint.Y, rb.MinPoint.Z)[axis_idx]
        hi = (rb.MaxPoint.X, rb.MaxPoint.Y, rb.MaxPoint.Z)[axis_idx]
        return (lo + hi) / 2.0
    except Exception:
        return None


def _outer_loop_length_in(f):
    for lp in _iter(f.EdgeLoops):
        if lp.IsOuterEdgeLoop:
            total = 0.0
            for e in _iter(lp.Edges):
                total += _edge_length_cm(e)
            return total * CM_TO_IN
    return 0.0


def _edge_length_cm(e):
    """Arc length in cm via the edge evaluator (lines, arcs, splines)."""
    try:
        min_p, max_p = e.Evaluator.GetParamExtents()
        return e.Evaluator.GetLengthAtParam(min_p, max_p)
    except Exception:
        return 0.0


# ---------------------------------------------------------------------------
# Orchestrator: probe -> path -> nine fields + part_class
# ---------------------------------------------------------------------------
def extract_part_cost_fields(part_doc, part_number=None):
    nr = []
    ident = part_number or "<unknown>"
    try:
        if not part_number:
            try:
                ident = part_doc.PropertySets.Item(DESIGN_TRACKING).Item("Part Number").Value
            except Exception:
                ident = "<unknown>"
        # Capability probe (== the macro): sheet-metal iff CostDataFlatLengthInches exists.
        _, is_sheet_metal = _read_iprop(part_doc, SHEET_METAL_PROBE)
        if is_sheet_metal:
            fields = build_sheet_metal_fields(part_doc, ident, nr)
            fields["part_class"] = "sheet_metal"
        else:
            fields = build_tube_fields(part_doc, ident, nr)
            fields["part_class"] = "tube"
        return fields, nr
    except Exception as e:
        _flag(nr, ident, "cost_extraction", f"failed: {e!r}")
        return {"part_class": None}, nr
