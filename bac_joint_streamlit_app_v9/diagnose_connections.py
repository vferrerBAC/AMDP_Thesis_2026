"""
diagnose_connections.py -- explain why a Block 1 result has the connection count it does.

Usage:
    python diagnose_connections.py block1_result.json

Get the JSON from the app: Step 2 (Geometry) -> "Download Block 1 JSON".
Then run this and paste the output. It classifies every connection so we can see
whether the extra count is duplicate member occurrences, phantom near-contacts,
or legitimate multi-face expansion -- concentrated by joint (J003 / J004 / ...).
"""
from __future__ import annotations

import json
import sys
from collections import Counter, defaultdict


def main(path: str) -> int:
    data = json.load(open(path, encoding="utf-8"))
    members = data.get("members", []) or []
    conns = data.get("connections", []) or []

    print(f"\nFile: {path}")
    print(f"Members: {len(members)}   Joints: {len(data.get('joints', []))}   "
          f"Connections: {len(conns)}")

    # --- 1) duplicate member occurrences (same physical part, >1 occurrence) ---
    # Group members by (part_number, bom_description); >1 in a group with distinct
    # occurrence_path means the part is instanced more than once -> doubles contacts.
    by_part = defaultdict(list)
    for m in members:
        key = (m.get("part_number", ""), m.get("bom_description", ""))
        by_part[key].append(m.get("occurrence_path") or m.get("occurrence_name"))
    dup_parts = {k: v for k, v in by_part.items() if len(v) > 1}
    print(f"\n[1] Parts appearing as more than one occurrence: {len(dup_parts)}")
    for (pn, desc), paths in list(dup_parts.items())[:15]:
        print(f"    x{len(paths)}  {pn or '(no part#)'}  {desc[:40]!r}")

    # --- 2) detection method + per-joint breakdown ---------------------------
    by_method = Counter(c.get("detection_method", "?") for c in conns)
    print(f"\n[2] By detection method: {dict(by_method)}")

    per_joint = Counter(c.get("joint_id", "?") for c in conns)
    print("\n[3] Connections per joint (worst first):")
    for jid, n in per_joint.most_common():
        marker = "  <-- heavy" if n >= 4 else ""
        print(f"    {jid}: {n}{marker}")

    # --- 4) exact duplicate member-pairs within a joint ----------------------
    pair_key = Counter(
        (c.get("joint_id"), frozenset((c.get("member_a"), c.get("member_b"))))
        for c in conns
    )
    exact_dups = {k: n for k, n in pair_key.items() if n > 1}
    print(f"\n[4] Same member-pair repeated in the same joint: {len(exact_dups)}")
    for (jid, pair), n in list(exact_dups.items())[:15]:
        print(f"    {jid}: x{n}  {sorted(pair)}")

    # --- 5) coincident locations (overlapping diamonds) ----------------------
    loc_key = Counter(
        (c.get("joint_id"),
         tuple(round(float(v), 1) for v in (c.get("location") or [0, 0, 0])[:3]))
        for c in conns
    )
    coincident = {k: n for k, n in loc_key.items() if n > 1}
    print(f"\n[5] Connections stacked at the same point (rounded 0.1 in): "
          f"{len(coincident)}")
    for (jid, loc), n in list(coincident.items())[:15]:
        print(f"    {jid}: x{n} at {loc}")

    # --- 6) inferred / near-but-not-touching ---------------------------------
    n_inferred = sum(1 for c in conns if c.get("is_inferred"))
    print(f"\n[6] Connections flagged 'inferred' (near, not confirmed touching): "
          f"{n_inferred}")

    print("\nInterpretation:")
    print("  section [1] non-empty  -> duplicate occurrences are doubling contacts")
    print("  section [4] non-empty  -> literal duplicate connections in the data")
    print("  section [5] non-empty  -> stacked diamonds (the visual duplication)")
    print("  method has face_contact + multi -> multi-face expansion is expected\n")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python diagnose_connections.py block1_result.json")
        raise SystemExit(2)
    raise SystemExit(main(sys.argv[1]))
