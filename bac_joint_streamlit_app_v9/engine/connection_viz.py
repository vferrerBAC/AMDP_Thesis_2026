"""
engine/connection_viz.py
========================
3D viewer for the connection layer. Works straight off ``block1_result.json`` --
no Inventor session, no COM, nothing to install beyond plotly (already a
dependency). Renders to an interactive HTML file or into the Streamlit app.

It draws the three geometries that the pipeline actually builds, which are easy
to conflate:

  1. CONTACT PATCH (solid, teal)
     The true clipped face outline, bolt holes removed. THIS IS THE CONNECTION.
     One per touching face pair.

  2. OBB RING (dashed, amber)
     What JointLocatorV16 measures instead: face A's bounding rectangle clipped
     by face B's. Drawn on top of the patch, so the visible gap between amber and
     teal IS the error the bounding-box method introduces. On a square butt joint
     they coincide exactly. On a coped, mitred or angled contact -- every diagonal
     brace in the frame -- the amber ring is visibly larger, and a larger bearing
     area makes a capacity check UNCONSERVATIVE.

  3. CLUSTER SPHERE (wireframe, gray)
     Joints are NOT boxed. detect_joints merges contact points when they fall
     within ``joint_cluster_tol_in`` of the point that seeded the joint -- a
     sphere. This is the radius that decides whether two connections belong to
     one joint or to two, and it is the thing to look at when a joint has
     swallowed contacts that should have stayed separate (sphere visibly
     spanning two physical nodes) or split one node in two (two spheres where
     the model has one connection region).

Usage
-----
    from engine.connection_viz import figure_from_result, write_html

    data = json.load(open("block1_result.json"))
    write_html(data, "connections.html")          # open in any browser

    # or inside Streamlit:
    import streamlit as st
    st.plotly_chart(figure_from_result(data), use_container_width=True)
"""

from __future__ import annotations

import json
import math
from typing import Any, Mapping, Optional, Sequence

import plotly.graph_objects as go

# Thesis palette: teal = truth, amber = the approximation, gray = structure.
TEAL = "#1b9e93"
TEAL_FILL = "rgba(27,158,147,0.35)"
AMBER = "#e8a33d"
GRAY = "#8a8f98"
GRAY_FAINT = "rgba(138,143,152,0.22)"
RED = "#d1495b"


# --------------------------------------------------------------------------- #
def _get(obj: Any, key: str, default=None):
    if isinstance(obj, Mapping):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _patch_to_world(conn) -> tuple[list, list]:
    """Map a connection's 2D patch outline back into world coordinates.

    Returns (exterior_3d, [hole_3d, ...]). Empty if the connection is a
    centerline fallback with no real patch.
    """
    frame = _get(conn, "patch_frame")
    ext2d = _get(conn, "patch_exterior_2d") or []
    if not frame or not ext2d:
        return [], []

    o = frame["origin"] if isinstance(frame, Mapping) else frame[0]
    xa = frame["x_axis"] if isinstance(frame, Mapping) else frame[1]
    ya = frame["y_axis"] if isinstance(frame, Mapping) else frame[2]

    def m(ring):
        return [
            (o[0] + xa[0] * u + ya[0] * v,
             o[1] + xa[1] * u + ya[1] * v,
             o[2] + xa[2] * u + ya[2] * v)
            for (u, v) in ring
        ]

    return m(ext2d), [m(h) for h in (_get(conn, "patch_holes_2d") or [])]


def _is_centerline_only(conn) -> bool:
    """True when a connection has neither a true contact patch nor an OBB ring,
    so there is nothing real to draw for it -- a pair the face pass never
    confirmed (typically angled braces at a corner)."""
    ext, _ = _patch_to_world(conn)
    return len(ext) < 3 and len(_get(conn, "obb_polygon_3d") or []) < 3


def _fan_mesh(ring: Sequence[Sequence[float]]):
    """Triangle-fan indices for a ring. Fine for the convex-ish contact patches
    we get; a concave patch renders slightly optimistically but the outline
    (drawn separately) still tells the truth."""
    n = len(ring)
    if n < 3:
        return None
    i = [0] * (n - 2)
    j = list(range(1, n - 1))
    k = list(range(2, n))
    return i, j, k


def _closed(ring):
    return list(ring) + [ring[0]] if ring else []


def _sphere(center, r, n=14):
    """Wireframe sphere as a set of latitude/longitude line segments."""
    xs, ys, zs = [], [], []
    for a in range(n):
        th = 2 * math.pi * a / n
        for b in range(n + 1):
            ph = math.pi * b / n
            xs.append(center[0] + r * math.sin(ph) * math.cos(th))
            ys.append(center[1] + r * math.sin(ph) * math.sin(th))
            zs.append(center[2] + r * math.cos(ph))
        xs.append(None); ys.append(None); zs.append(None)
    for b in range(1, n):
        ph = math.pi * b / n
        for a in range(n + 1):
            th = 2 * math.pi * a / n
            xs.append(center[0] + r * math.sin(ph) * math.cos(th))
            ys.append(center[1] + r * math.sin(ph) * math.sin(th))
            zs.append(center[2] + r * math.cos(ph))
        xs.append(None); ys.append(None); zs.append(None)
    return xs, ys, zs


# --------------------------------------------------------------------------- #
def figure_from_result(
    data: Mapping[str, Any],
    *,
    show_members: bool = True,
    show_patches: bool = True,
    show_obb: bool = True,
    show_spheres: bool = True,
    show_centroids: bool = True,
    show_centerline_only: bool = False,
    joint_filter: Optional[str] = None,
) -> go.Figure:
    """Build the interactive 3D figure from a Block 1 result dict.

    ``show_centerline_only`` controls the red ✕ markers for pairs the face pass
    never confirmed. Off by default: those connections are dropped from the
    figure entirely, so they are not drawn and not counted in the subtitle.
    """
    members = data.get("members", [])
    joints = data.get("joints", [])
    conns = data.get("connections", [])
    cluster_tol = data.get("joint_cluster_tol_in")

    if joint_filter:
        joints = [j for j in joints if j.get("joint_id") == joint_filter]
        conns = [c for c in conns if c.get("joint_id") == joint_filter]

    if not show_centerline_only:
        conns = [c for c in conns if not _is_centerline_only(c)]

    fig = go.Figure()

    # ---- members: centerlines ------------------------------------------- #
    if show_members and members:
        xs, ys, zs = [], [], []
        for m in members:
            a, b = m.get("start_point"), m.get("end_point")
            if not a or not b:
                continue
            xs += [a[0], b[0], None]
            ys += [a[1], b[1], None]
            zs += [a[2], b[2], None]
        fig.add_trace(go.Scatter3d(
            x=xs, y=ys, z=zs, mode="lines", name="member centerlines",
            line=dict(color=GRAY, width=4), hoverinfo="skip",
        ))

    # ---- joint cluster spheres ------------------------------------------ #
    if show_spheres and cluster_tol and joints:
        xs, ys, zs = [], [], []
        for j in joints:
            sx, sy, sz = _sphere(j["location"], cluster_tol)
            xs += sx; ys += sy; zs += sz
        fig.add_trace(go.Scatter3d(
            x=xs, y=ys, z=zs, mode="lines",
            name=f"joint cluster radius ({cluster_tol:.2f} in)",
            line=dict(color=GRAY_FAINT, width=1), hoverinfo="skip",
        ))

    # ---- joint nodes ----------------------------------------------------- #
    if joints:
        fig.add_trace(go.Scatter3d(
            x=[j["location"][0] for j in joints],
            y=[j["location"][1] for j in joints],
            z=[j["location"][2] for j in joints],
            mode="markers+text", name="joints",
            marker=dict(size=6, color=GRAY, symbol="diamond"),
            text=[j.get("joint_id", "") for j in joints],
            textposition="top center", textfont=dict(size=9, color=GRAY),
            hovertext=[
                f"{j.get('joint_id')}<br>{j.get('joint_type', '')} "
                f"{j.get('geom_descriptor', '')}<br>"
                f"{len(j.get('member_names', []))} members"
                for j in joints
            ],
            hoverinfo="text",
        ))

    # ---- contact patches (the connections) ------------------------------- #
    n_patch = n_obb = n_div = 0
    first_patch = first_obb = True
    # Connections that draw neither a patch nor an OBB ring (centerline-only
    # pairs the face pass could not confirm -- typically angled braces at a
    # corner). Collected here so they are shown as markers instead of being
    # silently dropped from the model.
    centerline_only: list = []

    for c in conns:
        ext, holes = _patch_to_world(c)
        cid = _get(c, "connection_id")
        div = _get(c, "obb_divergence")
        review = _get(c, "needs_review")

        hover = (
            f"<b>{cid}</b><br>"
            f"{_get(c, 'member_a')}<br>&nbsp;&nbsp;+ {_get(c, 'member_b')}<br>"
            f"joint {_get(c, 'joint_id')} &middot; face pair "
            f"{_get(c, 'face_pair_index')}<br>"
            f"type {_get(c, 'connection_type')}<br>"
            f"area {_get(c, 'contact_area_in2')} in²<br>"
            f"OBB area {_get(c, 'obb_area_in2')} in²<br>"
            f"divergence {'-' if div is None else f'{div * 100:.0f}%'}<br>"
            f"holes {_get(c, 'hole_count')}<br>"
            f"weld path {_get(c, 'weld_length_in')} in"
        )
        if review:
            hover += f"<br><b>review:</b> {_get(c, 'review_reason')}"

        if show_patches and len(ext) >= 3:
            n_patch += 1
            tri = _fan_mesh(ext)
            if tri:
                i, j, k = tri
                fig.add_trace(go.Mesh3d(
                    x=[p[0] for p in ext], y=[p[1] for p in ext],
                    z=[p[2] for p in ext], i=i, j=j, k=k,
                    color=TEAL, opacity=0.35,
                    name="contact patch (true)", legendgroup="patch",
                    showlegend=first_patch, hovertext=hover, hoverinfo="text",
                ))
                first_patch = False
            ring = _closed(ext)
            fig.add_trace(go.Scatter3d(
                x=[p[0] for p in ring], y=[p[1] for p in ring],
                z=[p[2] for p in ring], mode="lines",
                line=dict(color=TEAL, width=4), legendgroup="patch",
                showlegend=False, hovertext=hover, hoverinfo="text",
            ))
            for h in holes:
                hr = _closed(h)
                fig.add_trace(go.Scatter3d(
                    x=[p[0] for p in hr], y=[p[1] for p in hr],
                    z=[p[2] for p in hr], mode="lines",
                    line=dict(color=RED, width=3), legendgroup="patch",
                    showlegend=False, hoverinfo="skip",
                ))

        # ---- the V16 bounding-box ring, drawn on top --------------------- #
        obb = _get(c, "obb_polygon_3d") or []
        if show_obb and len(obb) >= 3:
            n_obb += 1
            if div is not None and div > 0.20:
                n_div += 1
            ring = _closed([tuple(p) for p in obb])
            fig.add_trace(go.Scatter3d(
                x=[p[0] for p in ring], y=[p[1] for p in ring],
                z=[p[2] for p in ring], mode="lines",
                line=dict(color=AMBER, width=3, dash="dash"),
                name="OBB ring (JointLocatorV16)", legendgroup="obb",
                showlegend=first_obb,
                hovertext=(
                    f"<b>{cid}</b> OBB<br>area {_get(c, 'obb_area_in2')} in²"
                    f"<br>vs true {_get(c, 'contact_area_in2')} in²"
                    f"<br>overstated by "
                    f"{'-' if div is None else f'{div * 100:.0f}%'}"
                ),
                hoverinfo="text",
            ))
            first_obb = False

        # ---- centerline-only connections ---------------------------------- #
        # No true contact patch AND no OBB ring: the face pass never confirmed
        # this pair (common on angled braces at a corner), so nothing above
        # drew it. Record its location so it shows as a marker instead of
        # disappearing from the model. Existence is checked directly -- NOT via
        # the show_* toggles -- so toggling patches/OBB off does not turn every
        # connection into a marker.
        if len(ext) < 3 and len(obb) < 3:
            loc = _get(c, "location")
            if loc:
                centerline_only.append((loc, hover))

    # ---- centerline-only connection markers ------------------------------- #
    if centerline_only:
        fig.add_trace(go.Scatter3d(
            x=[loc[0] for loc, _ in centerline_only],
            y=[loc[1] for loc, _ in centerline_only],
            z=[loc[2] for loc, _ in centerline_only],
            mode="markers",
            name=f"centerline connection ({len(centerline_only)}, face contact unconfirmed)",
            marker=dict(size=6, color=RED, symbol="x",
                        line=dict(color=RED, width=1)),
            hovertext=[h for _, h in centerline_only], hoverinfo="text",
        ))

    # ---- contact centroids (the yellow dots in the Inventor macro) -------- #
    if show_centroids and conns:
        face = [c for c in conns if _get(c, "detection_method") == "face_contact"]
        if face:
            fig.add_trace(go.Scatter3d(
                x=[_get(c, "location")[0] for c in face],
                y=[_get(c, "location")[1] for c in face],
                z=[_get(c, "location")[2] for c in face],
                mode="markers", name="contact centroids",
                marker=dict(size=4, color=AMBER),
                hovertext=[_get(c, "connection_id") for c in face],
                hoverinfo="text",
            ))

    subtitle = (
        f"{len(conns)} connections &middot; {len(joints)} joints &middot; "
        f"{n_div} of {n_obb} contacts where the OBB overstates area by >20%"
    )
    if centerline_only:
        subtitle += (
            f" &middot; {len(centerline_only)} centerline-only "
            "(face contact unconfirmed, shown as red ✕)"
        )
    fig.update_layout(
        title=dict(text=f"IJET connection layer<br><sup>{subtitle}</sup>"),
        scene=dict(
            aspectmode="data",
            xaxis_title="X (in)", yaxis_title="Y (in)", zaxis_title="Z (in)",
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.0),
        margin=dict(l=0, r=0, t=70, b=0),
        height=780,
    )
    return fig


def write_html(data: Mapping[str, Any], path: str = "connections.html", **kw) -> str:
    """Render to a standalone interactive HTML file. Opens in any browser --
    no Streamlit, no Inventor."""
    fig = figure_from_result(data, **kw)
    fig.write_html(path, include_plotlyjs="cdn")
    return path


def divergence_report(data: Mapping[str, Any]) -> list[dict]:
    """Every connection where the bounding-box method overstates contact area,
    worst first. This is the table for the thesis: each row is a joint whose
    capacity V16 would have computed against too much bearing area."""
    rows = []
    for c in data.get("connections", []):
        div = c.get("obb_divergence")
        if div is None:
            continue
        rows.append({
            "connection_id": c.get("connection_id"),
            "joint_id": c.get("joint_id"),
            "members": f"{c.get('member_a')} + {c.get('member_b')}",
            "true_area_in2": c.get("contact_area_in2"),
            "obb_area_in2": c.get("obb_area_in2"),
            "overstated_pct": round(div * 100, 1),
            "holes": c.get("hole_count"),
        })
    rows.sort(key=lambda r: r["overstated_pct"], reverse=True)
    return rows


if __name__ == "__main__":
    import sys

    src = sys.argv[1] if len(sys.argv) > 1 else "block1_result.json"
    out = sys.argv[2] if len(sys.argv) > 2 else "connections.html"
    with open(src) as fh:
        d = json.load(fh)
    print(f"wrote {write_html(d, out)}")
    rep = divergence_report(d)
    if rep:
        worst = rep[0]
        print(f"worst OBB divergence: {worst['connection_id']} "
              f"({worst['overstated_pct']}% overstated)")
