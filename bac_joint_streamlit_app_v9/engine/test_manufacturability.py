import unittest

from engine.manufacturability import (
    assess_part,
    build_manufacturability_inputs,
    check_apb,
    check_mpb,
    check_tube_laser,
)


class ManufacturabilityRulesTests(unittest.TestCase):
    def test_mpb_orientation(self):
        self.assertTrue(check_mpb(50, 160))
        self.assertTrue(check_mpb(160, 50))
        self.assertFalse(check_mpb(60, 160))  # strict boundary

    def test_tube_laser_orientation_and_boundary(self):
        self.assertTrue(check_tube_laser(10, 300))
        self.assertTrue(check_tube_laser(300, 10))
        self.assertFalse(check_tube_laser(12.5, 300))

    def test_apb_requires_envelope_and_diagonal(self):
        self.assertTrue(check_apb(40, 100, 10, "GLV"))
        self.assertFalse(check_apb(40, 120, 10, "GLV"))
        self.assertFalse(check_apb(40, 100, 8, "GLV"))

    def test_assess_part_material_alias(self):
        result = assess_part(10, 200, "12 Gauge", "HDG", "P1")
        self.assertEqual(result["material"], "GLV")
        self.assertTrue(result["tube_laser_ok"])

    def test_selected_process_prefers_tube_laser(self):
        # 10 x 100, 12 GLV -> TL passes (10 < 12.5, 100 < 334.65) and MPB
        # passes; APB fails (x=10 is not > 18.3). Should select Tube Laser.
        result = assess_part(10, 100, 12, "GLV", "P_TL_MPB")
        self.assertTrue(result["tube_laser_ok"])
        self.assertTrue(result["mpb_ok"])
        self.assertFalse(result["apb_ok"])
        self.assertEqual(result["selected_process"], "Tube Laser")
        self.assertEqual(result["eligible_processes"], ["Tube Laser", "Manual Panel Bender"])
        self.assertEqual(result["status"], "TUBE LASER OK")

    def test_selected_process_prefers_apb_over_mpb(self):
        # 40 x 100, 12 GLV -> APB passes and MPB passes; TL fails (40 >= 12.5).
        # Should select Automated Panel Bender (cheaper than manual).
        result = assess_part(40, 100, 12, "GLV", "P_APB_MPB")
        self.assertFalse(result["tube_laser_ok"])
        self.assertTrue(result["apb_ok"])
        self.assertTrue(result["mpb_ok"])
        self.assertEqual(result["selected_process"], "Automated Panel Bender")
        self.assertEqual(
            result["eligible_processes"],
            ["Automated Panel Bender", "Manual Panel Bender"],
        )
        self.assertEqual(result["status"], "AUTOMATED PANEL BENDER OK")

    def test_selected_process_falls_back_to_mpb(self):
        # 40 x 160, 12 GLV -> MPB passes (both sides < envelope);
        # TL fails (40 >= 12.5); APB fails (y=160 > 149.6 long-max for 12 GLV,
        # and diagonal sqrt(40^2 + 160^2) ~= 164.9 > 157.48). Selects Manual PB.
        result = assess_part(40, 160, 12, "GLV", "P_MPB_ONLY")
        self.assertFalse(result["tube_laser_ok"])
        self.assertFalse(result["apb_ok"])
        self.assertTrue(result["mpb_ok"])
        self.assertEqual(result["selected_process"], "Manual Panel Bender")
        self.assertEqual(result["status"], "MANUAL PANEL BENDER OK")

    def test_selected_process_none_when_not_manufacturable(self):
        # 200 x 200 is outside every envelope.
        result = assess_part(200, 200, 12, "GLV", "P_NONE")
        self.assertFalse(result["manufacturable_any_process"])
        self.assertIsNone(result["selected_process"])
        self.assertEqual(result["status"], "NOT MANUFACTURABLE")

    def test_block1_adapter_uses_member_geometry(self):
        data = {
            "members": [
                {
                    "part_number": "P1",
                    "material": "Galvanized Steel",
                    "length": 200,
                    "bom_description": "12 GA frame",
                    "cross_section": {"width": 2, "depth": 10},
                }
            ]
        }
        rows, flags = build_manufacturability_inputs(data)
        self.assertEqual(len(flags), 0)
        self.assertEqual(rows[0]["x_width_in"], 10)
        self.assertEqual(rows[0]["y_length_in"], 200)
        self.assertEqual(rows[0]["gauge"], 12)
        self.assertEqual(rows[0]["material"], "GLV")


if __name__ == "__main__":
    unittest.main()
