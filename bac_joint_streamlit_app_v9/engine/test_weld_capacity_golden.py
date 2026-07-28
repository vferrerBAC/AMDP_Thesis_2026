"""Golden test: the Python port must reproduce the validated welded template.

Counterpart of ``test_bolted_capacity_golden.py``. Expected values were obtained
by populating the "Welded Connections" sheet of
``Capacity Analysis Template.xlsx`` with the five tube-to-tube cases below and
recalculating with a spreadsheet engine, then reading back columns:

    O  Base metal area          in^2
    P  Weld length              in
    Q  Fexx                     ksi
    R  Weld leg size            in
    U  Max Shear Strength       kips   <-- governing capacity
    V  Value Used                      <-- governing mode

Template values are in kips; the module returns lbf, so U expectations are x1000.

Run:  pytest test_weld_capacity_golden.py   (or  python test_weld_capacity_golden.py)
"""

import math

from weld_capacity import (
    capacity_from_template_inputs,
    weld_capacity,
    weld_shear_capacity_lbf,
    fexx_from_electrode,
    weld_leg_size_in,
    thickness_from_gauge,
    DEFAULT_ELECTRODE,
)

# (branch_material, branch_gauge, branch_w, branch_h, chord_material, chord_gauge, electrode)
CASES = {
    "J001": ("Carbon Steel Tube", 12, 4.0, 4.0, "Carbon Steel Tube", 10, "E70XX"),
    "J002": ("Carbon Steel Tube", 16, 2.0, 3.0, "Carbon Steel Tube", 16, "E70XX"),
    "J003": ("Stainless Steel Tube", 10, 6.0, 4.0, "Stainless Steel Tube", 8, "E80XX"),
    "J004": ("Carbon Steel Tube", 8, 8.0, 6.0, "Stainless Steel Tube", 8, "E60XX"),
    "J005": ("Carbon Steel Tube", 14, 3.0, 2.0, "Carbon Steel Tube", 12, "E90XX"),
}

# col:  O area,     P len, Q Fexx, R leg,  U kips,       V mode
EXPECTED = {
    "J001": (1.483900, 8.0, 70.0, 0.125, 22.2705000, "Weld Shear Strength"),
    "J002": (0.528336, 5.0, 70.0, 0.125, 13.9190625, "Weld Shear Strength"),
    "J003": (2.475484, 10.0, 80.0, 0.125, 31.8150000, "Weld Shear Strength"),
    "J004": (4.190364, 14.0, 60.0, 0.125, 33.4057500, "Weld Shear Strength"),
    "J005": (0.652044, 5.0, 90.0, 0.125, 17.6051880, "Base Material Strength"),
}

REL = 1e-6


def test_golden_all_columns():
    """Every intermediate column must match the recalculated sheet, not just
    the final number -- an accidental compensating error would otherwise pass."""
    for jid, args in CASES.items():
        area, wlen, fexx, leg, u_kips, mode = EXPECTED[jid]
        r = capacity_from_template_inputs(*args)
        assert math.isclose(r.base_metal_area_in2, area, rel_tol=REL), f"{jid} col O"
        assert math.isclose(r.weld_length_in, wlen, rel_tol=REL), f"{jid} col P"
        assert math.isclose(r.fexx_ksi, fexx, rel_tol=REL), f"{jid} col Q"
        assert math.isclose(r.weld_leg_size_in, leg, rel_tol=REL), f"{jid} col R"
        assert math.isclose(r.shear_lbf, u_kips * 1000.0, rel_tol=REL), f"{jid} col U"
        assert r.governing_mode == mode, f"{jid} col V"


def test_golden_base_metal_can_govern():
    """J005 is the case where column S < column T. If a refactor ever makes
    weld metal always govern, this catches it."""
    r = capacity_from_template_inputs(*CASES["J005"])
    assert r.governing_mode == "Base Material Strength"


# --- Column-level unit checks ----------------------------------------------

def test_col_Q_fexx_two_digit_only():
    assert fexx_from_electrode("E60XX") == 60.0
    assert fexx_from_electrode("E70XX") == 70.0
    assert fexx_from_electrode("E90XX") == 90.0
    assert fexx_from_electrode("e70") == 70.0
    # The sheet would return 10 ksi for E100XX (MID reads one digit). We refuse.
    assert fexx_from_electrode("E100XX") is None
    assert fexx_from_electrode("fillet_weld") is None
    assert fexx_from_electrode(None) is None


def test_col_R_leg_size_table():
    assert weld_leg_size_in(0.054) == 0.125     # min t <= 1/4
    assert weld_leg_size_in(0.25) == 0.125
    assert weld_leg_size_in(0.375) == 0.1875    # 1/4 < t <= 1/2
    assert weld_leg_size_in(0.625) == 0.25      # 1/2 < t <= 3/4
    assert weld_leg_size_in(1.00) == 0.3125     # t > 3/4


def test_gauge_thickness_table_includes_odd_gauges():
    """The welded data sheet carries gauges the bolted sheet does not."""
    assert thickness_from_gauge(12) == 0.095
    assert thickness_from_gauge(9) == 0.148
    assert thickness_from_gauge(11) == 0.12
    assert thickness_from_gauge(15) == 0.072
    assert thickness_from_gauge(99) is None


# --- Live connection-schedule path -----------------------------------------

def _row(**over):
    row = {
        "branch_material": "Carbon Steel Tube",
        "branch_gauge": 12,
        "branch_width_in": 4.0,
        "branch_height_in": 4.0,
        "chord_material": "Carbon Steel Tube",
        "chord_gauge": 10,
        "fastener_type": "E70XX",
    }
    row.update(over)
    return row


def test_live_path_matches_template_path():
    r_live = weld_capacity(_row())
    r_tmpl = capacity_from_template_inputs(*CASES["J001"])
    assert math.isclose(r_live.shear_lbf, r_tmpl.shear_lbf, rel_tol=REL)
    assert r_live.governing_mode == r_tmpl.governing_mode
    assert r_live.needs_review is False


def test_explicit_thickness_overrides_gauge():
    r = weld_capacity(_row(branch_t_in=0.095, chord_t_in=0.127))
    r_gauge = weld_capacity(_row())
    assert math.isclose(r.shear_lbf, r_gauge.shear_lbf, rel_tol=REL)


def test_non_tube_material_flags_review():
    """The template's IFS resolves tubes only; a plate would be #N/A."""
    r = weld_capacity(_row(branch_material="Carbon Steel Plate"))
    assert r.needs_review is True
    assert "#N/A" in r.basis


def test_unknown_electrode_assumes_default_and_flags():
    r = weld_capacity(_row(fastener_type="fillet_weld"))
    assert r.electrode == DEFAULT_ELECTRODE
    assert r.needs_review is True
    assert r.shear_lbf > 0


def test_missing_geometry_routes_to_review():
    for bad in (_row(branch_width_in=0), _row(branch_height_in=0),
                _row(branch_gauge=None, branch_t_in=0)):
        r = weld_capacity(bad)
        assert r.shear_lbf is None
        assert r.needs_review is True


def test_wall_thicker_than_section_routes_to_review():
    """Column O would go negative; must not propagate a garbage area."""
    r = weld_capacity(_row(branch_width_in=0.15, branch_height_in=0.15, branch_gauge=8))
    assert r.shear_lbf is None
    assert r.needs_review is True


def test_shim_returns_zero_not_none_on_review():
    cap, basis = weld_shear_capacity_lbf(_row(branch_width_in=0))
    assert cap == 0.0
    assert "review" in basis.lower()


def test_shim_matches_module():
    cap, _ = weld_shear_capacity_lbf(_row())
    assert cap == weld_capacity(_row()).shear_lbf


# --- Integration with the screening pipeline -------------------------------

def test_joint_checks_routes_welded_joints():
    """Regression: welded_joint used to fall through to 0.0 / 'Unsupported'."""
    from joint_checks import estimate_screening_capacity_lbf
    cap, basis = estimate_screening_capacity_lbf(
        {"connection_type": "welded_joint", **_row()}
    )
    assert cap > 0.0
    assert "Unsupported" not in basis
    assert "welded template" in basis


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
