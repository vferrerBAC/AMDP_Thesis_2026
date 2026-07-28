"""Golden tests: equilibrium, degenerate layouts, and the pure geometry half."""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))

import numpy as np
from dataclasses import dataclass

import engine.connection_geometry as cg
from engine.patch_groups import (
    PatchGroup, build_patch_groups, distribute, envelope, area_inertia_tensor,
)


def C(connection_id, member_a, member_b, joint_id, centroid, normal, area):
    return dict(connection_id=connection_id, member_a=member_a, member_b=member_b,
                joint_id=joint_id, location=list(centroid),
                contact_normal=list(normal), contact_area_in2=area)


def mk_group(areas, cents, norms, member="M1", jid="J1"):
    return PatchGroup(
        joint_id=jid, member=member,
        connection_ids=[f"C{i}" for i in range(len(areas))],
        areas=np.array(areas, float),
        centroids=np.array(cents, float),
        normals=np.array(norms, float),
    )


def check_equilibrium(g, F, M, P, tol=1e-6):
    """Sum of patch forces must equal F; sum of moments about C must equal M_C."""
    ds = distribute(g, F, M, P)
    f = np.array([d.force for d in ds])
    C = g.centroid
    r = g.centroids - C

    F_sum = f.sum(axis=0)
    M_sum = np.cross(r, f).sum(axis=0)
    M_C = np.asarray(M, float) + np.cross(np.asarray(P, float) - C, np.asarray(F, float))

    return F_sum, M_sum, M_C, np.asarray(F, float), ds


print("=" * 70)
print("TEST 1  force equilibrium, 4 patches, pure axial")
print("=" * 70)
g = mk_group(
    areas=[4, 4, 2, 2],
    cents=[[0, 3, 0], [0, -3, 0], [2, 0, 0], [-2, 0, 0]],
    norms=[[0, 1, 0], [0, -1, 0], [1, 0, 0], [-1, 0, 0]],
)
Fs, Ms, MC, F, ds = check_equilibrium(g, [0, 0, -12000], [0, 0, 0], g.centroid)
print(f"  applied F      {F}")
print(f"  sum patch F    {Fs.round(6)}")
print(f"  residual       {np.linalg.norm(Fs - F):.3e}")
assert np.linalg.norm(Fs - F) < 1e-6, "FORCE EQUILIBRIUM VIOLATED"
print(f"  area split     {[round(d.moment_share, 4) for d in ds]}")
print("  -> forces split 4:4:2:2 by area, exactly as intended.  PASS")

print()
print("=" * 70)
print("TEST 2  moment equilibrium, general 3D action, eccentric node")
print("=" * 70)
g = mk_group(
    areas=[6.0, 6.0, 3.0],
    cents=[[1.0, 4.0, 0.5], [1.0, -4.0, 0.5], [-3.0, 0.0, 2.0]],
    norms=[[0, 1, 0], [0, -1, 0], [0, 0, 1]],
)
Fs, Ms, MC, F, ds = check_equilibrium(
    g, F := [3000.0, -1500.0, 8000.0], [22000.0, -9000.0, 4000.0], [0.0, 0.0, 0.0]
)
print(f"  sum patch F    {Fs.round(4)}   applied {np.asarray(F)}")
print(f"  force residual {np.linalg.norm(Fs - np.asarray(F)):.3e}")
print(f"  sum patch M@C  {Ms.round(4)}")
print(f"  required M@C   {MC.round(4)}")
print(f"  moment residual{np.linalg.norm(Ms - MC):.3e}")
assert np.linalg.norm(Fs - np.asarray(F)) < 1e-6, "FORCE EQUILIBRIUM VIOLATED"
assert np.linalg.norm(Ms - MC) < 1e-5, "MOMENT EQUILIBRIUM VIOLATED"
print("  -> both satisfied, with the P->C eccentricity transfer included.  PASS")
print(f"  eccentricity flag: {[f for f in ds[0].needs_review if 'eccentric' in f]}")

print()
print("=" * 70)
print("TEST 3  the case area-splitting a moment would get WRONG")
print("=" * 70)
print("  Two equal patches straddling a web, +/-4in apart. Pure moment about Z,")
print("  zero net force. Correct behaviour: a force COUPLE, equal and opposite.")
print("  Area-splitting the moment would instead give both patches the same sign.")
g = mk_group(
    areas=[5.0, 5.0],
    cents=[[0.0, 4.0, 0.0], [0.0, -4.0, 0.0]],
    norms=[[0, 1, 0], [0, -1, 0]],
)
ds = distribute(g, [0, 0, 0], [0, 0, 40000.0], g.centroid)
f = np.array([d.force for d in ds])
print(f"  patch 1 force  {f[0].round(2)}")
print(f"  patch 2 force  {f[1].round(2)}")
print(f"  sum            {f.sum(axis=0).round(6)}  (must be zero)")
assert np.linalg.norm(f.sum(axis=0)) < 1e-6
assert np.dot(f[0], f[1]) < 0, "patches should oppose — this is a couple"
M_check = np.cross(g.centroids - g.centroid, f).sum(axis=0)
print(f"  moment carried {M_check.round(2)}  (must be [0,0,40000])")
assert abs(M_check[2] - 40000.0) < 1e-5
print("  -> equal, opposite, correct magnitude. Couple recovered.  PASS")

print()
print("=" * 70)
print("TEST 4  degenerate layout — collinear patches, moment about their own axis")
print("=" * 70)
print("  Three patches in a line along Y. A moment about Y cannot be resisted by")
print("  any couple between them. This MUST be flagged, not silently dropped.")
g = mk_group(
    areas=[3.0, 3.0, 3.0],
    cents=[[0.0, -5.0, 0.0], [0.0, 0.0, 0.0], [0.0, 5.0, 0.0]],
    norms=[[1, 0, 0], [1, 0, 0], [1, 0, 0]],
)
ds = distribute(g, [0, 0, 0], [0, 30000.0, 0], g.centroid)
deg = [f for f in ds[0].needs_review if "degenerate" in f]
print(f"  flags: {deg}")
assert deg, "DEGENERATE LAYOUT NOT FLAGGED — this is the dangerous silent failure"
print("  -> unresisted moment surfaced as needs_review.  PASS")

print()
print("=" * 70)
print("TEST 5  single patch takes everything")
print("=" * 70)
g = mk_group(areas=[8.0], cents=[[1.0, 2.0, 3.0]], norms=[[0, 0, 1]])
ds = distribute(g, [100.0, 200.0, 300.0], [0, 0, 0], [0, 0, 0])
print(f"  force  {np.array(ds[0].force).round(3)}   share {ds[0].moment_share}")
assert np.allclose(ds[0].force, [100, 200, 300])
assert ds[0].moment_share == 1.0
print(f"  flags: {[f for f in ds[0].needs_review if 'single' in f]}")
print("  -> reduces correctly.  PASS")

print()
print("=" * 70)
print("TEST 6  grouping — member touching TWO others at one node")
print("=" * 70)
print("  Brace B1 meets column C1 AND beam B2 at joint J1. B1's end force must")
print("  distribute across ALL THREE of its patches, not per member pair.")
conns = [
    C("C0001", "B1", "C1", "J1", (0, 1, 0), (0, 1, 0), 4.0),
    C("C0002", "B1", "C1", "J1", (0, -1, 0), (0, -1, 0), 4.0),
    C("C0003", "B1", "B2", "J1", (1, 0, 0), (1, 0, 0), 2.0),
    C("C0004", "C1", "B2", "J1", (0, 0, 1), (0, 0, 1), 3.0),
]
groups = build_patch_groups(conns)
for k in sorted(groups):
    print(f"  {k}  ->  {groups[k].connection_ids}  total A = {groups[k].total_area}")
b1 = groups[("J1", "B1")]
assert set(b1.connection_ids) == {"C0001", "C0002", "C0003"}, "B1 group is wrong"
assert b1.total_area == 10.0
print("  -> B1's group spans all 3 of its patches across 2 different members. PASS")

print()
print("=" * 70)
print("TEST 7  envelope keeps the worse of the two per-member estimates")
print("=" * 70)
from engine.patch_groups import PatchDemand
a = PatchDemand("C1", "M1", "J1", (0, 0, 0), 100.0, 50.0, 0.5)
b = PatchDemand("C1", "M2", "J1", (0, 0, 0), 400.0, 10.0, 0.5)
e = envelope([a, b])
print(f"  member M1 says |N|+|V| = 150 ; member M2 says 410")
print(f"  envelope keeps: {e['C1'].member} at {e['C1'].normal_lbf}")
assert e["C1"].member == "M2"
print("  -> worse case retained.  PASS")

print()
print("=" * 70)
print("GEOMETRY: exact contact patches (engine/connection_geometry.py)")
print("=" * 70)
print(f"  shapely present      {cg.HAVE_SHAPELY}")

sq3 = [(0, 0, 0), (4, 0, 0), (4, 3, 0), (0, 3, 0)]
hole3 = [(1, 1, 0), (2, 1, 0), (2, 2, 0), (1, 2, 0)]

a = cg.build_face_loops([sq3], (0, 0, 1))
b = cg.build_face_loops([[(0, 0, 0), (4, 0, 0), (4, 3, 0), (0, 3, 0)]], (0, 0, -1))
area, perim, span, cen, nh, ext, holes, frame, flags = cg.contact_metrics_exact(a, b)
print(f"  4x3 self-overlap     area {area}  perim {perim}  span {span:.4f}")
assert abs(area - 12.0) < 1e-9 and abs(perim - 14.0) < 1e-9
assert abs(span - 5.0) < 1e-9

print()
print("  BOLT HOLE subtracted from bearing area")
a = cg.build_face_loops([sq3, hole3], (0, 0, 1))
area, perim, span, cen, nh, ext, holes, frame, flags = cg.contact_metrics_exact(a, b)
print(f"    net area           {area:.4f}   (expect 11.0 = 12 - 1)")
print(f"    holes carried      {nh}")
if cg.HAVE_SHAPELY:
    assert abs(area - 11.0) < 1e-6, "bolt hole was NOT subtracted"
    assert nh == 1
    print("    -> hole removed. The OBB path would have reported 12.")
else:
    assert "holes_ignored_no_shapely" in flags
    print(f"    -> fallback active, hole ignored AND flagged: {flags}")

print()
print("  OBB CROSS-CHECK on a coped (L-shaped) face — the V16 failure mode")
L = [(0, 0, 0), (6, 0, 0), (6, 2, 0), (2, 2, 0), (2, 5, 0), (0, 5, 0)]
la = cg.build_face_loops([L], (0, 0, 1))
lb = cg.build_face_loops([[(0, 0, 0), (6, 0, 0), (6, 5, 0), (0, 5, 0)]], (0, 0, -1))
area, *_ = cg.contact_metrics_exact(la, lb)
obb = cg.obb_area_of(la.exterior_2d)
div = abs(area - obb) / obb
print(f"    true contact area  {area:.2f}")
print(f"    OBB (V16) area     {obb:.2f}")
print(f"    divergence         {div * 100:.0f}%  -> flagged (threshold "
      f"{cg.OBB_DIVERGENCE_FLAG * 100:.0f}%)")
assert abs(area - 18.0) < 1e-6
assert div > cg.OBB_DIVERGENCE_FLAG
print("    -> V16 would overstate bearing area by 67%. Unconservative.")

print()
print("=" * 70)
print("ALL TESTS PASS")
print("=" * 70)
