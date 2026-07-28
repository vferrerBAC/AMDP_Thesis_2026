"""IJET manufacturability screening engine.

This module is aligned with the existing IJET architecture:
- consumes the same Block 1 dictionary stored in ``st.session_state.block1_raw``;
- uses canonical Block 1 dimensions in inches;
- returns JSON-serializable records and review flags, like ``cost_adapter.py``;
- keeps machine limits in one configuration table for maintainability.

Machine capability rules are implemented exactly as provided. All inequalities
are strict. For APB, the orientation envelope and diagonal rule must both pass.
"""
from __future__ import annotations

import math
import re
from collections import Counter
from typing import Any, Dict, Iterable, Mapping, Optional, Tuple


# ---------------------------------------------------------------------------
# Machine capability constants (inches)
# ---------------------------------------------------------------------------

MPB_SHORT_MAX = 60.0
MPB_LONG_MAX = 168.0

TUBE_LASER_SHORT_MAX = 12.5
TUBE_LASER_LONG_MAX = 334.65

APB_SHORT_MIN = 18.3
APB_SHORT_MAX = 60.0
APB_LONG_MIN = 27.75
APB_DIAGONAL_MAX = 157.48

# Gauge/material-specific APB maximum long-side limit.
# Missing keys mean APB is not defined for that specification.
APB_LONG_MAX_BY_SPEC: Dict[Tuple[int, str], float] = {
    (10, "GLV"): 118.11,
    (10, "SST"): 82.67,
    (12, "GLV"): 149.60,
    (12, "SST"): 108.26,
    (14, "GLV"): 149.60,
    (14, "SST"): 118.11,
    (16, "GLV"): 149.60,
    (16, "SST"): 149.60,
}

SUPPORTED_SPECS = {
    (8, "GLV"), (8, "SST"),
    (10, "GLV"), (10, "SST"),
    (12, "GLV"), (12, "SST"),
    (14, "GLV"), (14, "SST"),
    (16, "GLV"), (16, "SST"),
}

SUPPORTED_GAUGES = (8, 10, 12, 14, 16)
SUPPORTED_MATERIALS = ("GLV", "SST")


# ---------------------------------------------------------------------------
# Normalization helpers
# ---------------------------------------------------------------------------

_MATERIAL_ALIASES = {
    "GLV": "GLV",
    "GALV": "GLV",
    "GALVANIZED": "GLV",
    "GALVANIZED STEEL": "GLV",
    "HDG": "GLV",
    "HOT DIP GALVANIZED": "GLV",
    "HOT-DIP GALVANIZED": "GLV",
    "GI": "GLV",
    "SST": "SST",
    "SS": "SST",
    "STAINLESS": "SST",
    "STAINLESS STEEL": "SST",
}


def canonical_material(value: Any) -> Optional[str]:
    """Return GLV/SST for recognized material labels, otherwise ``None``."""
    if value is None:
        return None
    text = str(value).strip().upper()
    if not text:
        return None

    if text in _MATERIAL_ALIASES:
        return _MATERIAL_ALIASES[text]

    # Tolerate longer Inventor material labels/descriptions.
    if "STAINLESS" in text or re.search(r"\bSST\b", text):
        return "SST"
    if "GALV" in text or re.search(r"\bHDG\b", text) or re.search(r"\bGLV\b", text):
        return "GLV"
    return None


def canonical_gauge(value: Any) -> Optional[int]:
    """Normalize numeric/string gauge representations to a supported integer."""
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, int):
        return value if value in SUPPORTED_GAUGES else None

    if isinstance(value, float):
        if math.isnan(value):
            return None
        if value.is_integer() and int(value) in SUPPORTED_GAUGES:
            return int(value)
        return None

    text = str(value).strip().upper()
    if not text:
        return None

    match = re.search(r"(?<!\d)(8|10|12|14|16)\s*(?:GAUGE|GA|G)?\b", text)
    if match:
        return int(match.group(1))

    try:
        number = int(float(text))
        return number if number in SUPPORTED_GAUGES else None
    except (TypeError, ValueError):
        return None


def _positive_float(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or number <= 0:
        return None
    return number


# ---------------------------------------------------------------------------
# Machine checks
# ---------------------------------------------------------------------------


def check_mpb(x: float, y: float) -> bool:
    """Manual Panel Bender envelope, orientation-independent."""
    return (
        (x < MPB_SHORT_MAX and y < MPB_LONG_MAX)
        or (x < MPB_LONG_MAX and y < MPB_SHORT_MAX)
    )


def check_tube_laser(x: float, y: float) -> bool:
    """Tube Laser envelope, orientation-independent."""
    return (
        (x < TUBE_LASER_SHORT_MAX and y < TUBE_LASER_LONG_MAX)
        or (x < TUBE_LASER_LONG_MAX and y < TUBE_LASER_SHORT_MAX)
    )


def check_apb(x: float, y: float, gauge: int, material: str) -> bool:
    """Automated Panel Bender envelope AND diagonal limit."""
    max_long = APB_LONG_MAX_BY_SPEC.get((gauge, material))
    if max_long is None:
        return False

    envelope_ok = (
        (
            x > APB_SHORT_MIN
            and x < APB_SHORT_MAX
            and y > APB_LONG_MIN
            and y < max_long
        )
        or (
            x > APB_LONG_MIN
            and x < max_long
            and y > APB_SHORT_MIN
            and y < APB_SHORT_MAX
        )
    )
    diagonal_ok = (x * x + y * y) < (APB_DIAGONAL_MAX * APB_DIAGONAL_MAX)
    return envelope_ok and diagonal_ok


def assess_part(
    x: Any,
    y: Any,
    gauge: Any,
    material: Any,
    part_identifier: Optional[str] = None,
    quantity: Any = 1,
) -> Dict[str, Any]:
    """Strictly assess one complete part input record.

    Raises ``ValueError`` when a required input is missing or unsupported. The
    Streamlit adapter uses ``evaluate_manufacturability_rows`` for tolerant,
    review-friendly batch evaluation.
    """
    x_value = _positive_float(x)
    y_value = _positive_float(y)
    gauge_value = canonical_gauge(gauge)
    material_value = canonical_material(material)

    if x_value is None or y_value is None:
        raise ValueError("x width and y length must be positive numbers in inches.")
    if gauge_value is None:
        raise ValueError(f"Unsupported or missing gauge: {gauge!r}.")
    if material_value is None:
        raise ValueError(f"Unsupported or missing material: {material!r}.")
    if (gauge_value, material_value) not in SUPPORTED_SPECS:
        raise ValueError(f"Unsupported specification: {gauge_value} Gauge {material_value}.")

    try:
        qty = int(float(quantity))
    except (TypeError, ValueError):
        qty = 1
    qty = max(qty, 1)

    mpb_ok = check_mpb(x_value, y_value)
    tube_laser_ok = check_tube_laser(x_value, y_value)
    apb_available = (gauge_value, material_value) in APB_LONG_MAX_BY_SPEC
    apb_ok = check_apb(x_value, y_value, gauge_value, material_value)

    # Eligible processes are listed in cost priority order (cheapest first).
    # Tube Laser is preferred whenever it applies; the Automated Panel Bender
    # is cheaper than the Manual Panel Bender. Tube Laser and the Automated
    # Panel Bender are mutually exclusive geometrically (APB requires both
    # sides > 18.3 in; TL requires one side < 12.5 in), so the TL-before-APB
    # order never causes a wrong selection.
    eligible = []
    if tube_laser_ok:
        eligible.append("Tube Laser")
    if apb_ok:
        eligible.append("Automated Panel Bender")
    if mpb_ok:
        eligible.append("Manual Panel Bender")

    selected_process = eligible[0] if eligible else None

    if tube_laser_ok:
        status = "TUBE LASER OK"
    elif apb_ok:
        status = "AUTOMATED PANEL BENDER OK"
    elif mpb_ok:
        status = "MANUAL PANEL BENDER OK"
    else:
        status = "NOT MANUFACTURABLE"

    failure_reasons: Dict[str, str] = {}
    if not tube_laser_ok:
        failure_reasons["Tube Laser"] = (
            "Requires one orientation with one side < 12.5 in and the other < 334.65 in."
        )
    if not mpb_ok:
        failure_reasons["Manual Panel Bender"] = (
            "Requires one orientation with one side < 60 in and the other < 168 in."
        )
    if not apb_available:
        failure_reasons["Automated Panel Bender"] = (
            f"Automated Panel Bender is not defined for {gauge_value} Gauge {material_value}."
        )
    elif not apb_ok:
        diagonal = math.hypot(x_value, y_value)
        failure_reasons["Automated Panel Bender"] = (
            "Fails the Automated Panel Bender orientation envelope and/or diagonal limit. "
            f"Calculated diagonal = {diagonal:.3f} in; required < {APB_DIAGONAL_MAX:g} in."
        )

    return {
        "part_identifier": part_identifier or "<unknown>",
        "quantity": qty,
        "x_width_in": x_value,
        "y_length_in": y_value,
        "gauge": gauge_value,
        "material": material_value,
        "tube_laser_ok": tube_laser_ok,
        "mpb_ok": mpb_ok,
        "apb_ok": apb_ok,
        "manufacturable_any_process": bool(eligible),
        "eligible_processes": eligible,
        "selected_process": selected_process,
        "status": status,
        "failure_reasons": failure_reasons,
    }


# ---------------------------------------------------------------------------
# Block 1 adapter
# ---------------------------------------------------------------------------


def _first_positive(rec: Mapping[str, Any], keys: Iterable[str]) -> tuple[Optional[float], str]:
    for key in keys:
        if key in rec:
            value = _positive_float(rec.get(key))
            if value is not None:
                return value, key
    return None, ""


def _part_identifier(rec: Mapping[str, Any], index: int) -> str:
    for key in ("part_number", "part_identifier", "occurrence_path", "occurrence_name", "member_id"):
        value = rec.get(key)
        if value not in (None, ""):
            return str(value)
    return f"PART_{index}"


def _extract_material(rec: Mapping[str, Any]) -> tuple[Optional[str], str]:
    for key in ("NCx_Material", "ncx_material", "material"):
        value = rec.get(key)
        normalized = canonical_material(value)
        if normalized:
            return normalized, key

    for key in ("bom_description", "description", "part_number"):
        normalized = canonical_material(rec.get(key))
        if normalized:
            return normalized, f"inferred from {key}"
    return None, ""


def _extract_gauge(rec: Mapping[str, Any]) -> tuple[Optional[int], str]:
    for key in ("Gauge", "gauge"):
        gauge = canonical_gauge(rec.get(key))
        if gauge is not None:
            return gauge, key

    # Inventor Gauge iProperty, read into Block 1's per-member cost extension
    # (member.cost.Gauge). This is the same authoritative source the structural
    # tab uses, and the usual place the gauge actually lives.
    cost = rec.get("cost") or {}
    if isinstance(cost, Mapping):
        gauge = canonical_gauge(cost.get("Gauge") or cost.get("gauge"))
        if gauge is not None:
            return gauge, "cost.Gauge (iProperty)"

    cs = rec.get("cross_section") or {}
    if isinstance(cs, Mapping):
        gauge = canonical_gauge(cs.get("gauge"))
        if gauge is not None:
            return gauge, "cross_section.gauge"

    # Last-resort metadata inference. This is surfaced in the source column so
    # the user can review it in the editable table.
    for key in ("bom_description", "description", "part_number"):
        gauge = canonical_gauge(rec.get(key))
        if gauge is not None:
            return gauge, f"inferred from {key}"
    return None, ""


def _extract_dimensions(rec: Mapping[str, Any]) -> tuple[Optional[float], Optional[float], str]:
    """Derive x=width and y=length from IJET/Block 1 fields.

    Priority:
    1. Explicit manufacturing flat-pattern dimensions.
    2. Explicit x/y fields.
    3. Block 1 member geometry: x uses the largest cross-section envelope
       dimension (conservative for Tube Laser screening); y uses member length.
    """
    x, x_src = _first_positive(
        rec,
        (
            "CostDataFlatWidthInches",
            "flat_width_inches",
            "x_width_in",
            "x",
        ),
    )
    y, y_src = _first_positive(
        rec,
        (
            "CostDataFlatLengthInches",
            "flat_length_inches",
            "y_length_in",
            "y",
            "length",
        ),
    )

    cs = rec.get("cross_section") or {}
    if x is None and isinstance(cs, Mapping):
        section_dims = [
            v for v in (
                _positive_float(cs.get("width")),
                _positive_float(cs.get("depth")),
            )
            if v is not None
        ]
        if section_dims:
            x = max(section_dims)
            x_src = "max(cross_section.width, cross_section.depth)"

    src = ", ".join(v for v in (x_src, y_src) if v)
    return x, y, src


def build_manufacturability_inputs(
    block1_data: Mapping[str, Any],
    aggregate_by_identifier: bool = True,
) -> tuple[list[Dict[str, Any]], list[Dict[str, Any]]]:
    """Adapt Block 1 JSON to editable manufacturability input rows.

    Returns ``(parts, review_flags)``. Missing fields are intentionally retained
    as ``None`` so the IJET data editor can be used to complete them.
    """
    raw_parts = block1_data.get("members", block1_data.get("parts", [])) or []
    counts = Counter(
        str(p.get("part_number"))
        for p in raw_parts
        if isinstance(p, Mapping) and p.get("part_number") not in (None, "")
    )

    rows: list[Dict[str, Any]] = []
    flags: list[Dict[str, Any]] = []
    seen: set[str] = set()

    for index, rec in enumerate(raw_parts, start=1):
        if not isinstance(rec, Mapping):
            continue

        ident = _part_identifier(rec, index)
        if aggregate_by_identifier and ident in seen:
            continue
        seen.add(ident)

        x, y, dim_source = _extract_dimensions(rec)
        gauge, gauge_source = _extract_gauge(rec)
        material, material_source = _extract_material(rec)

        qty = counts.get(str(rec.get("part_number")), 1) if rec.get("part_number") else 1

        row = {
            "part_identifier": ident,
            "quantity": qty,
            "x_width_in": x,
            "y_length_in": y,
            "gauge": gauge,
            "material": material,
            "dimension_source": dim_source or "missing",
            "gauge_source": gauge_source or "missing",
            "material_source": material_source or "missing",
        }
        rows.append(row)

        for field, value, issue in (
            ("x_width_in", x, "missing width; enter x in inches"),
            ("y_length_in", y, "missing length; enter y in inches"),
            ("gauge", gauge, "missing/unsupported gauge; select 8, 10, 12, 14, or 16"),
            ("material", material, "missing/unsupported material; select GLV or SST"),
        ):
            if value is None:
                flags.append({
                    "scope": "part",
                    "identifier": ident,
                    "field": field,
                    "issue": issue,
                })

    return rows, flags


def evaluate_manufacturability_rows(
    rows: Iterable[Mapping[str, Any]],
) -> tuple[list[Dict[str, Any]], list[Dict[str, Any]]]:
    """Tolerantly evaluate editable IJET rows.

    Invalid/incomplete rows become ``REVIEW REQUIRED`` rather than crashing the
    Streamlit tab. Returns ``(results, review_flags)``.
    """
    results: list[Dict[str, Any]] = []
    flags: list[Dict[str, Any]] = []

    for index, row in enumerate(rows, start=1):
        ident = str(row.get("part_identifier") or f"PART_{index}")
        try:
            result = assess_part(
                x=row.get("x_width_in"),
                y=row.get("y_length_in"),
                gauge=row.get("gauge"),
                material=row.get("material"),
                part_identifier=ident,
                quantity=row.get("quantity", 1),
            )
            results.append(result)
        except ValueError as exc:
            try:
                qty = max(int(float(row.get("quantity", 1))), 1)
            except (TypeError, ValueError):
                qty = 1
            results.append({
                "part_identifier": ident,
                "quantity": qty,
                "x_width_in": _positive_float(row.get("x_width_in")),
                "y_length_in": _positive_float(row.get("y_length_in")),
                "gauge": canonical_gauge(row.get("gauge")),
                "material": canonical_material(row.get("material")),
                "tube_laser_ok": None,
                "mpb_ok": None,
                "apb_ok": None,
                "manufacturable_any_process": None,
                "eligible_processes": [],
                "selected_process": None,
                "status": "REVIEW REQUIRED",
                "failure_reasons": {},
            })
            flags.append({
                "scope": "part",
                "identifier": ident,
                "field": "manufacturability_inputs",
                "issue": str(exc),
            })

    return results, flags


def summarize_results(results: Iterable[Mapping[str, Any]]) -> Dict[str, int]:
    """Quantity-weighted summary for IJET metric cards."""
    summary = {
        "part_types": 0,
        "total_quantity": 0,
        "tube_laser_ok_quantity": 0,
        "automated_panel_bender_ok_quantity": 0,
        "manual_panel_bender_ok_quantity": 0,
        "not_manufacturable_quantity": 0,
        "review_required_quantity": 0,
    }
    for row in results:
        summary["part_types"] += 1
        try:
            qty = max(int(float(row.get("quantity", 1))), 1)
        except (TypeError, ValueError):
            qty = 1
        summary["total_quantity"] += qty
        status = row.get("status")
        if status == "TUBE LASER OK":
            summary["tube_laser_ok_quantity"] += qty
        elif status == "AUTOMATED PANEL BENDER OK":
            summary["automated_panel_bender_ok_quantity"] += qty
        elif status == "MANUAL PANEL BENDER OK":
            summary["manual_panel_bender_ok_quantity"] += qty
        elif status == "NOT MANUFACTURABLE":
            summary["not_manufacturable_quantity"] += qty
        else:
            summary["review_required_quantity"] += qty
    return summary
