"""Golden test: the Python port must reproduce the validated template.

Expected values were obtained by populating
``Capacity Analysis Template.xlsx`` and recalculating with a spreadsheet
engine, then reading back columns S (MaxShearStrength), T (governing mode)
and V (MaxBoltTension). Template values are in kips; the module returns lbf,
so expectations are x1000.

Run:  pytest test_bolted_capacity_golden.py   (or  python test_bolted_capacity_golden.py)
"""

from bolted_capacity import (
    capacity_from_template_inputs,
    bolted_capacity,
    map_material,
    map_bolt,
)

TOL_LBF = 1.0  # < 1 lbf agreement with the spreadsheet

# (material, gauge, n, bolt_used, dia_in, Le_in, S_kips, mode, V_kips)
GOLDEN = [
    ("Galvanized Sheet Steel", 14, 2, "ASTM A307, Grade A", 0.3125, 1.5,
     1.43213, "Bolt Bearing", 5.17719),
    ("Carbon Steel Tube", 12, 2, "ASTM A307, Grade A", 0.3125, 1.5,
     2.88563, "Bolt Bearing", 5.17719),
    ("Carbon Steel Plate", 8, 2, "SAE J429, Grade 8", 0.75, 2.0,
     8.67510, "Bolt Bearing", 74.88281),
]


def test_matches_template():
    for mat, ga, n, bolt, dia, le, s_kips, mode, v_kips in GOLDEN:
        r = capacity_from_template_inputs(mat, ga, n, bolt, dia, le)
        assert abs(r.shear_lbf - s_kips * 1000.0) < TOL_LBF, (mat, ga, r.shear_lbf)
        assert abs(r.tension_lbf - v_kips * 1000.0) < TOL_LBF, (mat, ga, r.tension_lbf)
        assert r.governing_mode == mode, (mat, ga, r.governing_mode)
        assert not r.needs_review


def test_material_mapping():
    assert map_material("GLV-M5") == "Galvanized Sheet Steel"
    assert map_material("Galvanized steel") == "Galvanized Sheet Steel"
    assert map_material("carbon steel tube") == "Carbon Steel Tube"
    assert map_material("Unknown") is None


def test_unknown_material_routes_to_review():
    r = bolted_capacity({"material_grade": "Inconel 718", "n_fasteners": 2,
                         "diameter_in": 0.3125, "sheet_t_in": 0.067, "edge_dist_in": 1.5})
    assert r.needs_review and r.shear_lbf is None


def test_unknown_bolt_falls_back_to_a307_and_flags():
    r = bolted_capacity({"material_grade": "Galvanized Sheet Steel", "fastener_type": "mystery",
                         "n_fasteners": 2, "diameter_in": 0.3125, "sheet_t_in": 0.067,
                         "edge_dist_in": 1.5})
    assert r.needs_review and r.category == "A307"
    # value still matches the A307 golden row (thickness given directly)
    assert abs(r.shear_lbf - 1432.13) < 1.0


def test_missing_geometry_routes_to_review():
    r = bolted_capacity({"material_grade": "Galvanized Sheet Steel", "n_fasteners": 0,
                         "diameter_in": 0.3125, "sheet_t_in": 0.067, "edge_dist_in": 1.5})
    assert r.needs_review and r.shear_lbf is None


if __name__ == "__main__":
    test_matches_template()
    test_material_mapping()
    test_unknown_material_routes_to_review()
    test_unknown_bolt_falls_back_to_a307_and_flags()
    test_missing_geometry_routes_to_review()
    print("All golden tests passed.")
