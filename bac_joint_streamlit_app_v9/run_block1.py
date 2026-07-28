"""
run_block1.py -- one-shot live extraction + connection viewer.

Requires: Inventor open, with the assembly as the ACTIVE document.

    conda activate COOP26
    python run_block1.py

Writes, next to this file:
    block1_result.json      full Block 1 output (members, joints, connections)
    connections.html        interactive 3D viewer -- open in any browser
    obb_divergence.csv      every connection where the bounding-box method
                            overstates contact area, worst first

Then prints the health checks that tell you whether the face-loop path actually
bound to Inventor, or whether it silently fell back to bounding boxes.
"""

import csv
import json
import sys

from engine.block1 import run
from engine.connection_viz import write_html, divergence_report


def main() -> int:
    print("Connecting to Inventor and running Block 1...")
    result = run()

    with open("block1_result.json", "w") as fh:
        fh.write(result.to_json())
    data = json.loads(result.to_json())
    print(f"  wrote block1_result.json")

    members = data.get("members", [])
    joints = data.get("joints", [])
    conns = data.get("connections", [])
    diag = data.get("connection_diagnostics", {}) or {}

    if diag.get("fatal_error"):
        print("\n!! FACE DETECTION FAILED -- connections are centerline-only.")
        print(f"   {diag['fatal_error']}")
        print(diag.get("trace", ""))
        return 1

    write_html(data, "connections.html")
    print("  wrote connections.html")

    rep = divergence_report(data)
    if rep:
        with open("obb_divergence.csv", "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=list(rep[0].keys()))
            w.writeheader()
            w.writerows(rep)
        print(f"  wrote obb_divergence.csv ({len(rep)} rows)")

    # ----------------------------------------------------------------- #
    print("\n" + "=" * 62)
    print("MODEL")
    print("=" * 62)
    n_face = sum(1 for c in conns if c.get("detection_method") == "face_contact")
    print(f"  members                      {len(members)}")
    print(f"  joints                       {len(joints)}")
    print(f"  joint cluster radius         {data.get('joint_cluster_tol_in')} in")
    print(f"  connections                  {len(conns)}  ({n_face} from real faces)")
    print(f"  member pairs expanded        {diag.get('pairs_expanded', 0)}")
    print(f"  ...of which MULTI-FACE       {diag.get('pairs_multi_face', 0)}"
          "   <- pairs v6 collapsed into one connection")

    print("\n  where the connection count comes from:")
    print(f"    face pairs examined        {diag.get('face_pairs_examined', 0)}")
    print(f"    parallel planes            {diag.get('pairs_parallel', 0)}")
    print(f"    REJECTED, same-side        {diag.get('pairs_same_side_rejected', 0)}"
          "   <- flush faces, not touching")
    print(f"    normals opposed            {diag.get('pairs_normals_opposed', 0)}")
    print(f"    close enough to touch      {diag.get('pairs_offset_ok', 0)}")
    print(f"    raw contact patches        {diag.get('contacts_raw', 0)}")
    print(f"    MERGED coplanar fragments  {diag.get('patches_merged_coplanar', 0)}"
          "   <- one flange split by features")
    print(f"    coplanar but kept separate {diag.get('coplanar_groups_kept_separate', 0)}")
    print(f"    below min area, dropped    {diag.get('contacts_below_min_area', 0)}")
    print(f"    = face connections         {diag.get('contacts_found', 0)}")
    print(f"    pairs DROPPED, no contact  {diag.get('pairs_dropped_no_contact', 0)}"
          "   <- near, but not touching")

    up = diag.get("unconfirmed_pairs") or []
    if up:
        print("\n  member pairs the centerline detector paired up but whose faces")
        print("  do not touch (NOT emitted as connections):")
        for u in up[:15]:
            print(f"    - {u}")
        if len(up) > 15:
            print(f"    ... {len(up) - 15} more")

    print("\n" + "=" * 62)
    print("HEALTH CHECKS  (read these before trusting any number)")
    print("=" * 62)
    ok = True

    planar = diag.get("planar_faces", 0)
    loops = diag.get("faces_with_loops", 0)
    obb_only = diag.get("faces_obb_only", 0)
    if planar and loops / planar < 0.9:
        ok = False
        print(f"  [FAIL] faces_with_loops {loops} / planar_faces {planar}")
        print("         GetStrokes did not bind. You are silently back on")
        print("         BOUNDING-BOX geometry -- bolt holes are NOT subtracted")
        print("         and contact areas are overstated. Send me this output.")
    else:
        print(f"  [ok]   face loops read      {loops} / {planar} planar faces")
    if obb_only:
        print(f"         ({obb_only} faces fell back to OBB)")

    unor = diag.get("normals_unoriented", 0)
    if unor:
        ok = False
        print(f"  [FAIL] normals_unoriented   {unor}")
        print("         Face.IsParamReversed did not read, so outward normals are")
        print("         unknown and flush-but-not-touching faces cannot be")
        print("         rejected. Connections will be OVERCOUNTED.")
    else:
        print(f"  [ok]   outward normals      {diag.get('normals_oriented', 0)} oriented")

    if diag.get("loops_via_edges_unordered"):
        print(f"  [WARN] loops via unordered Edges "
              f"{diag['loops_via_edges_unordered']} (EdgeUses unavailable)")

    chorded = diag.get("curved_edges_chorded", 0)
    if chorded:
        ok = False
        print(f"  [WARN] curved edges chorded {chorded}")
        print("         Bolt holes are being flattened into polygons. Hole area")
        print("         is understated, so bearing area is overstated.")
    else:
        print(f"  [ok]   curved edges chorded 0")

    holes = diag.get("faces_with_holes", 0)
    print(f"  [--]   faces with holes     {holes}"
          "   <- 0 here means no bolt holes were found at all")

    if diag.get("pairs_unconfirmed"):
        print(f"  [WARN] unconfirmed pairs    {diag['pairs_unconfirmed']}")
        print("         Centerline says these members touch; the solids do not.")

    if diag.get("errors"):
        ok = False
        print("\n  COM errors:")
        for e in diag["errors"]:
            print(f"    - {e}")

    print("\n" + "=" * 62)
    print("OBB DIVERGENCE  (the thesis result)")
    print("=" * 62)
    if not rep:
        print("  no comparable contacts")
    else:
        bad = [r for r in rep if r["overstated_pct"] > 20]
        print(f"  {len(bad)} of {len(rep)} contacts overstated by >20% "
              f"by the V16 bounding-box method")
        print(f"\n  {'conn':8}{'joint':8}{'true':>8}{'OBB':>8}{'over%':>8}  members")
        for r in rep[:10]:
            print(f"  {r['connection_id']:8}{str(r['joint_id']):8}"
                  f"{r['true_area_in2']:8.2f}{r['obb_area_in2']:8.2f}"
                  f"{r['overstated_pct']:8.1f}  {r['members']}")
        if len(rep) > 10:
            print(f"  ... {len(rep) - 10} more in obb_divergence.csv")

    print("\nOpen connections.html to see it.")
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
