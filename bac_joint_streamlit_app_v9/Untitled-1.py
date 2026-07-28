# test_bbox.py  — run from the app root:  python test_bbox.py [path_to_block1.json]
import json, sys
from engine.loads.asce7_lrfd_loads import estimate_projected_area_ft2, has_inventor_envelope

path = sys.argv[1] if len(sys.argv) > 1 else "sample_data/block1_result.json"
UNIT = "in"   # <-- set to "cm" if your member coords are in cm, not inches

b1 = json.load(open(path))

print(f"\n--- {path} ---")
print("top-level keys:", list(b1.keys()))
print("envelope present?", has_inventor_envelope(b1))

# show the raw bbox keys it looks for
for k in ("bounding_box", "model_bounding_box", "envelope", "range_box"):
    if k in b1:
        print(f"  found '{k}':", b1[k])

# area in each wind direction with the CURRENT data
for d in ("X (+)", "Y (+)", "Z (+)"):
    a = estimate_projected_area_ft2(b1, d, UNIT)
    print(f"  area {d:6s} = {a:8.1f} ft^2")

# --- prove the wiring: inject a synthetic envelope only if none exists ---
if not has_inventor_envelope(b1):
    print("\nNo Inventor envelope -> currently on the FRAME-CENTERLINE fallback.")
    b1["bounding_box"] = {"min": [-150, -2, -20], "max": [150, 255, 20],
                          "source": "synthetic_test"}
    print("Injected a synthetic envelope to prove the app picks it up:")
    print("  envelope present?", has_inventor_envelope(b1))
    for d in ("X (+)", "Y (+)", "Z (+)"):
        a = estimate_projected_area_ft2(b1, d, UNIT)
        print(f"  area {d:6s} = {a:8.1f} ft^2  (from envelope)")