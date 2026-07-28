"""Joining cost — bolted vs welded, per joint.

Pure-Python port of the joining-cost columns of the "Joints List" sheet in
``cost_calculator - Clean - 12FEB26.xlsx``. Every line maps to a column so
results stay traceable to the validated workbook:

    R  Fastener Qty                = ROUND(joint length / fastener spacing)
    U  Extended Fastener Time      = ATF(type) * qty * PFD                 [sec]
    V  Extended Fastener Material  = purchase_cost(type) * qty             [USD]
    W  Extended Fastener Labor     = Joint_LaborRate * U                   [USD]
    X  Fully Burdened Fastener Cost= V + W                                 [USD]

    Z  Weld Setup Time             = Joint_WeldSetupTime (if welded)       [sec]
    AA Welding Time                = weld inches * Joint_WeldRate          [sec]
    AB Post-Weld Time              = weld inches * Joint_PostWeldRate      [sec]
    AC Fully Burdened Welding Cost = (Z + AA + AB) * PFD * Joint_LaborRate [USD]

    AD Fully Burdened Joining Cost = X + AC

Why this module exists separately from ``cost.py``: ``cost.py`` drives the whole
workbook over COM and therefore needs Excel on Windows. The joining-cost columns
above are a small, closed set of arithmetic with no chained lookups, so porting
them lets the joint recommendation compute capacity-per-dollar anywhere --
including in tests and on the thesis machine -- without a live Excel session.
Part-level cost (the 175-column BAC Part List) is NOT ported and still belongs
to ``cost.py``.

THE KEY LINK: the template hardcodes "Weld Inches" (col Y) to 100. It should be
the weld length actually implied by the joint geometry, which the capacity
template already computes as column P = branch width + branch height. Passing
that value here is what makes capacity and cost consistent for the same joint,
and is what makes a capacity-per-dollar comparison meaningful.

Note (flagged, not changed): the Personal Fatigue and Delay factor (col T) is a
regional VLOOKUP in the workbook. It is exposed here as a parameter defaulting
to 1.0 (no allowance). Set ``pfd`` to the regional value to match the sheet
exactly. It scales welding cost fully but scales only the LABOR half of fastener
cost, so it is not neutral in a bolted-vs-welded comparison.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional

# --- Constants (verbatim from the workbook's Joint_* named ranges) ----------

LABOR_RATE_USD_PER_SEC = 0.03866666666666666   # Joint_LaborRate
WELD_SETUP_TIME_SEC = 36.0                     # Joint_WeldSetupTime
WELD_RATE_SEC_PER_IN = 10.0                    # Joint_WeldRate
POST_WELD_RATE_SEC_PER_IN = 6.3                # Joint_PostWeldRate

# Joint_FastenerTypes -> (Joint_ATF sec each, Joint_FastenerPurchaseCost USD each)
FASTENER_DATA: Dict[str, tuple] = {
    'HW_TAPPER     5/16"':          (7.0, 0.3797),
    'HW_BOLT_ASSY     5/16"_DIA':   (30.0, 0.5233),
}

# The connection types IJET models map onto the workbook's two fastener types.
BOLT_FASTENER = 'HW_BOLT_ASSY     5/16"_DIA'
SCREW_FASTENER = 'HW_TAPPER     5/16"'

_CONNECTION_TO_FASTENER = {
    "bolted_bracket_joint": BOLT_FASTENER,
    "screwed_sheet_joint": SCREW_FASTENER,
}

DEFAULT_PFD = 1.0  # Personal Fatigue and Delay allowance; see module docstring


def _num(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except Exception:
        return default


@dataclass
class JoiningCost:
    """Cost of realising one joint with one connection method."""
    connection_type: str            # "bolted" | "screwed" | "welded"
    cost_usd: Optional[float]
    material_usd: Optional[float]
    labor_usd: Optional[float]
    time_sec: Optional[float]
    driver: str                     # what the cost scales with, for the UI
    needs_review: bool
    basis: str


# --- Fastener path (cols R, U, V, W, X) ------------------------------------

def fastener_cost(n_fasteners: Any,
                  fastener_type: str = BOLT_FASTENER,
                  pfd: float = DEFAULT_PFD) -> JoiningCost:
    """Cols U/V/W/X. Fastener quantity is taken DIRECTLY from the connection
    schedule (``n_fasteners``) rather than derived from joint length / spacing
    (col R), because the connection schedule is the authoritative source for the
    bolt group that the capacity check was run against. Using a different bolt
    count for cost than for capacity would make capacity-per-dollar meaningless.
    """
    n = _num(n_fasteners)
    if n <= 0:
        return JoiningCost("bolted", None, None, None, None, "fastener count",
                           True, "Fastener count is zero/missing. Routed to review.")

    if fastener_type not in FASTENER_DATA:
        return JoiningCost("bolted", None, None, None, None, "fastener count",
                           True, f"Fastener type {fastener_type!r} is not in the "
                                 f"workbook's Joint_FastenerTypes table. Routed to review.")

    atf_sec, purchase_usd = FASTENER_DATA[fastener_type]
    time_sec = atf_sec * n * pfd                       # U
    material = purchase_usd * n                        # V
    labor = LABOR_RATE_USD_PER_SEC * time_sec          # W
    total = material + labor                           # X

    kind = "screwed" if fastener_type == SCREW_FASTENER else "bolted"
    basis = (f"Joints List cols U-X: {n:g} x {fastener_type} "
             f"(ATF {atf_sec:g} s ea, ${purchase_usd:.4f} ea), PFD {pfd:g}. "
             f"Material ${material:.2f} + labor ${labor:.2f}.")
    return JoiningCost(kind, total, material, labor, time_sec,
                       "fastener count", False, basis)


# --- Weld path (cols Z, AA, AB, AC) ----------------------------------------

def welding_cost(weld_inches: Any,
                 include_setup: bool = True,
                 pfd: float = DEFAULT_PFD) -> JoiningCost:
    """Cols Z/AA/AB/AC.

    ``weld_inches`` should be the weld length implied by the joint geometry --
    i.e. column P of the capacity template (branch width + branch height), NOT
    the workbook's hardcoded 100. Pass ``WeldCapacity.weld_length_in`` straight
    in and capacity and cost describe the same weld.
    """
    length = _num(weld_inches)
    if length <= 0:
        return JoiningCost("welded", None, None, None, None, "weld inches",
                           True, "Weld length is zero/missing. Routed to review.")

    setup_sec = WELD_SETUP_TIME_SEC if include_setup else 0.0   # Z
    weld_sec = length * WELD_RATE_SEC_PER_IN                    # AA
    post_sec = length * POST_WELD_RATE_SEC_PER_IN               # AB
    time_sec = setup_sec + weld_sec + post_sec
    total = time_sec * pfd * LABOR_RATE_USD_PER_SEC             # AC

    basis = (f"Joints List cols Z-AC: {length:g} in of seam "
             f"(setup {setup_sec:g} s + weld {weld_sec:g} s + post-weld {post_sec:g} s "
             f"= {time_sec:g} s), PFD {pfd:g}, ${LABOR_RATE_USD_PER_SEC:.5f}/s.")
    return JoiningCost("welded", total, 0.0, total, time_sec,
                       "weld inches", False, basis)


# --- Convenience: cost a joint both ways -----------------------------------

def fastener_type_for(connection_type: Any) -> str:
    """Map an IJET connection_type to the workbook's fastener type."""
    return _CONNECTION_TO_FASTENER.get(str(connection_type or "").strip(), BOLT_FASTENER)


def joining_cost_total(fastener: Optional[JoiningCost],
                       weld: Optional[JoiningCost]) -> Optional[float]:
    """Col AD. Note the workbook SUMS both -- a real BAC joint may be fastened
    AND welded. The recommendation engine deliberately keeps them separate and
    does not call this; it is here for parity with the sheet."""
    values = [c.cost_usd for c in (fastener, weld) if c and c.cost_usd is not None]
    return sum(values) if values else None
