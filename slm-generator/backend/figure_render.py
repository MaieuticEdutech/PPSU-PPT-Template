#!/usr/bin/env python3
"""figure_render.py — deterministic diagram renderer for figure blocks.

The AI never draws: it proposes a small structured spec (schemas.
FIGURE_SPEC — a kind plus labelled items), and THIS module turns the spec
into a PNG with matplotlib. Same division of labour as the docx builder:
model proposes content, code owns layout, so a rendered figure can never
be malformed in ways a drawing model could produce. A spec that cannot be
rendered (e.g. a bar chart with non-numeric values) raises ValueError and
the caller keeps the DTP placeholder box instead — figures degrade, they
never break a unit.

Four diagram kinds cover the shapes SLM figures actually take:
  flow      — a sequential process (boxes joined by arrows)
  cycle     — a repeating process (boxes around a circle)
  hierarchy — one parent concept branching into children
  bar_chart — a quantitative comparison (requires numeric values)
"""
import textwrap

import matplotlib
matplotlib.use("Agg")            # headless — server threads, no display
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

import styles as st

KINDS = ("flow", "cycle", "hierarchy", "bar_chart")

_DPI = 150


def _hex(c):
    # styles constants are python-docx RGBColor objects whose str() is the
    # bare hex ("0E2841"); plain hex strings pass through unchanged
    return "#" + str(c)


def _wrap(text, width=16):
    return textwrap.fill(str(text), width)


def _box(ax, x, y, w, h, text, *, fill, text_color, fontsize=10,
         detail=None):
    ax.add_patch(FancyBboxPatch(
        (x - w / 2, y - h / 2), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.04",
        facecolor=fill, edgecolor="none"))
    if detail:
        ax.text(x, y + h * 0.16, _wrap(text), ha="center", va="center",
                fontsize=fontsize, color=text_color, fontweight="bold")
        ax.text(x, y - h * 0.22, _wrap(detail, 26), ha="center",
                va="center", fontsize=fontsize - 2.5, color=text_color)
    else:
        ax.text(x, y, _wrap(text), ha="center", va="center",
                fontsize=fontsize, color=text_color, fontweight="bold")


def _arrow(ax, xy_from, xy_to, color, curved=0.0):
    ax.add_patch(FancyArrowPatch(
        xy_from, xy_to, arrowstyle="-|>", mutation_scale=18,
        linewidth=2.2, color=color,
        connectionstyle=f"arc3,rad={curved}"))


def _new_ax(width, height):
    fig, ax = plt.subplots(figsize=(width, height), dpi=_DPI)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    return fig, ax


def _render_flow(spec, navy, orange):
    items = spec["items"]
    n = len(items)
    horizontal = n <= 4
    if horizontal:
        fig, ax = _new_ax(2.1 * n, 2.4)
        w, h = 0.82 / n, 0.62
        xs = [(i + 0.5) / n for i in range(n)]
        for i, (x, it) in enumerate(zip(xs, items)):
            _box(ax, x, 0.5, w, h, it["label"], fill=_hex(navy),
                 text_color="white", detail=it.get("detail"))
            if i:
                _arrow(ax, (xs[i - 1] + w / 2, 0.5), (x - w / 2, 0.5),
                       _hex(orange))
    else:
        fig, ax = _new_ax(5.4, 1.15 * n)
        w, h = 0.62, 0.8 / n
        ys = [1 - (i + 0.5) / n for i in range(n)]
        for i, (y, it) in enumerate(zip(ys, items)):
            _box(ax, 0.5, y, w, h, it["label"], fill=_hex(navy),
                 text_color="white", detail=it.get("detail"))
            if i:
                _arrow(ax, (0.5, ys[i - 1] - h / 2), (0.5, y + h / 2),
                       _hex(orange))
    return fig


def _render_cycle(spec, navy, orange):
    import math
    items = spec["items"]
    n = len(items)
    fig, ax = _new_ax(6.2, 6.2)
    r = 0.34
    pts = []
    for i in range(n):
        a = math.pi / 2 - 2 * math.pi * i / n     # clockwise from the top
        pts.append((0.5 + r * math.cos(a), 0.5 + r * math.sin(a)))
    for (x, y), it in zip(pts, items):
        _box(ax, x, y, 0.30, 0.14, it["label"], fill=_hex(navy),
             text_color="white", fontsize=9.5)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        # start/end just outside the boxes, curved with the circle
        dx, dy = x2 - x1, y2 - y1
        d = (dx * dx + dy * dy) ** 0.5 or 1.0
        pad = 0.09
        _arrow(ax, (x1 + dx / d * pad, y1 + dy / d * pad),
               (x2 - dx / d * pad, y2 - dy / d * pad),
               _hex(orange), curved=-0.25)
    return fig


def _render_hierarchy(spec, navy, orange):
    items = spec["items"]
    root = spec.get("root") or spec.get("title") or "Concept"
    n = len(items)
    fig, ax = _new_ax(max(5.5, 1.9 * n), 3.4)
    _box(ax, 0.5, 0.82, min(0.4, 2.2 / n if n else 0.4), 0.22, root,
         fill=_hex(navy), text_color="white")
    w = 0.86 / n
    for i, it in enumerate(items):
        x = (i + 0.5) / n
        _box(ax, x, 0.28, w * 0.92, 0.34, it["label"], fill="#FEF0E6",
             text_color=_hex(navy), fontsize=9.5, detail=it.get("detail"))
        ax.plot([0.5, x], [0.71, 0.46], color=_hex(orange), linewidth=2)
    return fig


def _render_bar_chart(spec, navy, orange):
    items = spec["items"]
    values = []
    for it in items:
        v = it.get("value")
        if not isinstance(v, (int, float)) or isinstance(v, bool):
            raise ValueError("bar_chart items need numeric 'value' fields")
        values.append(float(v))
    labels = [_wrap(it["label"], 12) for it in items]
    fig, ax = plt.subplots(figsize=(max(5.5, 1.4 * len(items)), 3.6),
                           dpi=_DPI)
    bars = ax.bar(labels, values, color=_hex(navy), width=0.55)
    ax.bar_label(bars, fmt="%g", padding=2, fontsize=9,
                 color=_hex(navy), fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#CCCCCC")
    ax.tick_params(labelsize=9, color="#CCCCCC")
    ax.margins(y=0.15)
    return fig


def render(spec, out_path):
    """Render one FIGURE_SPEC dict to a PNG at out_path. Raises ValueError
    for an unrenderable spec (caller falls back to the placeholder box)."""
    kind = spec.get("kind")
    items = spec.get("items") or []
    if kind not in KINDS:
        raise ValueError(f"unknown figure kind: {kind!r}")
    if not 2 <= len(items) <= 8:
        raise ValueError(f"figure needs 2-8 items, got {len(items)}")
    for it in items:
        if not str(it.get("label", "")).strip():
            raise ValueError("figure item with an empty label")

    navy, orange = st.NAVY, st.ORANGE
    fig = {"flow": _render_flow, "cycle": _render_cycle,
           "hierarchy": _render_hierarchy,
           "bar_chart": _render_bar_chart}[kind](spec, navy, orange)
    try:
        fig.savefig(str(out_path), bbox_inches="tight", facecolor="white")
    finally:
        plt.close(fig)
    return out_path
