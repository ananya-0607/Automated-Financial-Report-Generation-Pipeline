"""
chart_builder.py — NIIF PPT chart renderer
"""

import io
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.font_manager as _fm
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

from .chart_theme import ChartTheme, resolve, make_tick_formatter
from .ppt_logger  import get_logger, log_exceptions


# Active font file path — set by register_font(), used by _fp() helper.
# Using the file path directly (FontProperties(fname=...)) is more reliable
# than the family-name lookup (rcParams["font.family"]) which can silently
# fall back to DejaVu Sans when the name is not found.
_FONT_FILE: str = ""


def register_font(font_path: str) -> str:
    """
    Register a font file (.ttf / .otf) with matplotlib and store its path
    so all chart text uses it via FontProperties(fname=...) — bypassing the
    unreliable family-name lookup that can fall back to DejaVu Sans.

    Called from main.py / run_pipeline() using proj_font from the YAML.
    Returns the font family name, or "" on failure.
    """
    global _FONT_FILE
    if not font_path:
        return ""
    font_path = font_path.strip()
    if not os.path.exists(font_path):
        print(f"  [chart_builder] font not found, skipping: {font_path}")
        return ""
    try:
        _fm.fontManager.addfont(font_path)
        name = _fm.FontProperties(fname=font_path).get_name()
        _FONT_FILE = font_path          # store for _fp() helper below
        print(f"  [chart_builder] registered font '{name}' from {font_path}")
        return name
    except Exception as exc:
        print(f"  [chart_builder] could not register font {font_path}: {exc}")
        return ""


def _fp(size: float, weight: str = "normal") -> _fm.FontProperties:
    """Return FontProperties using the registered font file (reliable) or
    fall back to the theme font-family name if no file was registered."""
    if _FONT_FILE:
        return _fm.FontProperties(fname=_FONT_FILE, size=size)
    return _fm.FontProperties(size=size, weight=weight)

# Font sizes (scaled by slot size in build_chart)
BASE_TICK   = 11
BASE_LABEL  = 11
BASE_LEGEND = 10
BASE_LINE   = 1.8
DPI         = 200
SCALE_FACTOR = 2.0

# working values overwritten per chart
TICK_SIZE   = 11
LABEL_SIZE  = 11
LEGEND_SIZE = 10
LINE_W      = 1.8
LEG_LINE_W  = 1.8   # legend swatch thickness (= LINE_W unless legend_linewidth set)

plt.rcParams.update({
    "font.family":       ["Calibri"],
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.grid":         False,
    "figure.facecolor":  "white",
    "axes.facecolor":    "white",
})


def _apply_grid(ax, th, cfg=None):
    # Per-chart show_grid (from Excel) overrides YAML theme's grid_show.
    # cfg.show_grid is a bool set in config_reader; None/missing falls back to theme.
    if cfg is not None and hasattr(cfg, "show_grid") and cfg.show_grid is not None:
        grid_on = cfg.show_grid
    else:
        grid_on = th.grid_show
    if grid_on:
        ax.yaxis.grid(True, linestyle="-", linewidth=th.grid_width,
                      color=th.grid_color, alpha=th.grid_alpha)
    else:
        ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.set_axisbelow(True)


def _bar_width(n):
    if n <= 12:  return 0.55
    if n <= 24:  return 0.65
    if n <= 60:  return 0.75
    if n <= 120: return 0.82
    return 0.88


def _fmt(v, parens=True, decimals=1):
    negative = v < 0
    av = abs(v)
    if av == int(av) and av < 1e6:
        s = f"{int(av):,}"
    elif av >= 1e12: s = f"{av/1e12:.{decimals}f}T"
    elif av >= 1e9:  s = f"{av/1e9:.{decimals}f}B"
    elif av >= 1e6:  s = f"{av/1e6:.{decimals}f}M"
    else:            s = f"{av:,.{decimals}f}"
    if negative and parens: return f"({s})"
    if negative:            return f"-{s}"
    return s


def _tick_fmt(v, _):
    return _fmt(v, parens=True)


def _apply_yaxis(ax, y_min, y_max, y_breaks, th=None, secondary=False):
    th = th or ChartTheme()
    ticks = np.linspace(y_min, y_max, y_breaks + 1)
    pad = (y_max - y_min) * 0.03
    ax.set_ylim(y_min - pad, y_max + pad)
    ax.set_yticks(ticks)

    fmt_kind = th.y2_fmt if secondary else th.y_fmt
    tick_len  = th.y2_tick_length if secondary else th.y_tick_length
    ax.tick_params(axis="y", colors=th.tick_color,
                   length=tick_len, pad=15)

    show_labels = th.y2_show_labels if secondary else th.y_show_labels
    bold = th.y2_tick_bold if secondary else th.y_tick_bold
    fw_y = "bold" if bold else getattr(th, "font_weight", "normal")
    fp_y = _fp(TICK_SIZE, fw_y)

    if not show_labels:
        ax.set_yticklabels([])
    else:
        fmt_func   = make_tick_formatter(fmt_kind, th.data_lbl_decimals)
        label_strs = [fmt_func(v, None) for v in ticks]
        ax.set_yticklabels(label_strs, fontproperties=fp_y, color=th.tick_color)

    # rotated unit label
    if secondary:
        lbl = th.y2_axis_label
        lsz = th.y2_axis_label_size
    else:
        lbl = th.y_axis_label
        lsz = th.y_axis_label_size

    if lbl:
        rotation  = 270 if secondary else 90
        labelpad  = 20 if secondary else 10
        ax.set_ylabel(lbl, fontproperties=_fp(lsz), rotation=rotation,
                      labelpad=labelpad, color=th.tick_color)
    else:
        ax.set_ylabel("")


def _apply_xaxis(ax, x_labels, th=None, cfg=None):
    th = th or ChartTheme()
    positions  = [i for i, l in enumerate(x_labels) if l]
    labels_vis = [x_labels[i] for i in positions]
    if not labels_vis:
        ax.set_xticks([])
        return
    ax.set_xticks(positions)
    # Per-chart Excel x_tick_angle wins over YAML theme; None = use theme
    cfg_angle = getattr(cfg, "x_tick_angle", None)
    effective_angle = cfg_angle if cfg_angle is not None else th.x_tick_angle
    if effective_angle >= 0:
        rot = effective_angle
        ha  = "right" if rot > 0 else "center"
    else:
        rot, ha = (45, "right") if len(labels_vis) > 9 else (0, "center")
    fw_x = "bold" if th.x_tick_bold else getattr(th, "font_weight", "normal")
    fp_x = _fp(TICK_SIZE, fw_x)
    ax.set_xticklabels(labels_vis, rotation=rot, ha=ha,
                       color=th.tick_color, fontproperties=fp_x)
    show = th.show_x_axis
    # X-spine is fully independent of grid — always uses its own dedicated fields.
    spine_color = th.axis_color
    spine_lw    = th.axis_x_linewidth if show else 0

    ax.spines["bottom"].set_visible(show)
    if show:
        explicit_at_min = getattr(th, "x_axis_at_data_min", False)
        # Auto-fallback: when y_min > 0 the default spine at y=0 sits BELOW
        # the visible range and is invisible. Silently treat as x_axis_at_data_min
        # so the spine always appears at the bottom of the visible chart area.
        auto_at_min = (cfg is not None and cfg.y_min > 0 and not explicit_at_min)
        at_data_min = explicit_at_min or auto_at_min

        if at_data_min and cfg is not None:
            # Move the spine so tick marks anchor at y_min, then hide the spine
            # line itself and replace it with an axhline — this guarantees the
            # line renders on top of bar patches regardless of zorder/clipping.
            ax.spines["bottom"].set_position(("data", cfg.y_min))
            ax.spines["bottom"].set_visible(False)
            ax.axhline(y=cfg.y_min, color=spine_color,
                       linewidth=spine_lw, zorder=5, clip_on=True)
            # _apply_yaxis adds 3 % padding below y_min (ylim bottom < y_min).
            # That gap is real axes space — data/lines inside it are visible
            # below the spine.  Snap ylim bottom back to exactly y_min so the
            # spine is the absolute bottom of the visible area.
            _, hi = ax.get_ylim()
            ax.set_ylim(cfg.y_min, hi)
        else:
            ax.spines["bottom"].set_linewidth(spine_lw)
            ax.spines["bottom"].set_color(spine_color)
    # Tick mark width — fully independent of spine/grid linewidth.
    tick_lw = getattr(th, "x_tick_width", 0.5) if show else 0
    ax.tick_params(axis="x", length=th.x_tick_length if show else 0,
                   width=tick_lw, color=th.tick_color,
                   pad=getattr(th, "x_tick_pad", 4))

    # Direction: per-chart Excel wins → YAML theme → default (oldest→newest)
    cfg_rev = getattr(cfg, "x_axis_reverse", None)
    reverse = cfg_rev if isinstance(cfg_rev, bool) else th.x_axis_reverse
    if reverse:
        ax.invert_xaxis()


def _resolve_data_lbl_color(th, series_color):
    if th.data_lbl_color_match_series:
        return series_color
    return th.data_lbl_color


def _annotate_bar_value(ax, bar, val, th, color, y_min, y_max):
    if val == 0:
        return
    fw     = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
    col    = _resolve_data_lbl_color(th, color)
    dec    = th.data_lbl_decimals
    offset = (y_max - y_min) * 0.02
    va     = "bottom" if val >= 0 else "top"
    ypos   = bar.get_height() + offset if val >= 0 else bar.get_height() - offset
    ax.text(bar.get_x() + bar.get_width() / 2, ypos,
            _fmt(val, decimals=dec),
            ha="center", va=va,
            fontproperties=_fp(LABEL_SIZE - 2, fw), color=col,
            clip_on=True)


def _draw_bar(ax, x, y, color, show_values, y_min, y_max, n, bw=None, th=None,
              tick_positions=None, annotate=True):
    th = th or ChartTheme()
    bw_actual = bw or (th.bar_width if th.bar_width > 0 else _bar_width(n))
    bars = ax.bar(x, y, color=color, width=bw_actual, edgecolor="none", zorder=3)
    # show_values_all=True → annotate bars at x-axis tick positions (synced with visible labels)
    # show_values_all=False → last bar only (handled by caller via _annotate_last_bar)
    # annotate=False → skip all annotation here (caller will do joint spread with y2)
    do_all = bool(show_values)
    if do_all and annotate:
        tick_set = set(tick_positions) if tick_positions is not None else set(range(n))
        for xi, bar, val in zip(x, bars, y):
            if xi in tick_set:
                _annotate_bar_value(ax, bar, val, th, color, y_min, y_max)
    return bars, do_all


def _draw_line(ax, x, y_raw, color, n, th=None, series_idx=0):
    th = th or ChartTheme()
    y  = np.array([float(v) if v is not None else float("nan") for v in y_raw])
    ls_list = th.line_styles
    ls = ls_list[series_idx % len(ls_list)] if ls_list else th.line_style
    if th.markers:
        shapes = th.marker_shapes or ["o", "s", "^", "D"]
        mk     = shapes[series_idx % len(shapes)]
        ax.plot(x, y, color=color, linewidth=LINE_W, linestyle=ls,
                marker=mk, markersize=th.marker_size,
                zorder=4, solid_capstyle="round")
    else:
        ax.plot(x, y, color=color, linewidth=LINE_W, linestyle=ls,
                zorder=4, solid_capstyle="round")
    return y


def _annotate_last_bar(ax, bars, y1, th=None, color="#1a1a1a"):
    th = th or ChartTheme()
    if not bars:
        return
    val = y1[-1]
    if val == 0:
        return
    last_bar = bars[-1]
    fw  = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
    col = _resolve_data_lbl_color(th, color)
    dec = th.data_lbl_decimals
    ax.annotate(_fmt(val, decimals=dec),
        xy=(last_bar.get_x() + last_bar.get_width() / 2, last_bar.get_height()),
        xytext=(0, 8), textcoords="offset points",
        ha="center", va="bottom",
        fontproperties=_fp(LABEL_SIZE, fw),
        color=col, clip_on=False, annotation_clip=False, zorder=5)


def _annotate_all_line(ax, x, y_arr, th=None, color="#1a1a1a", tick_positions=None):
    """Annotate line data points — synced to x-axis tick positions when show_values_all=True."""
    th  = th or ChartTheme()
    fw  = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
    col = _resolve_data_lbl_color(th, color)
    dec = th.data_lbl_decimals
    tick_set = set(tick_positions) if tick_positions is not None else set(range(len(x)))
    for xi, val in zip(x, y_arr):
        if np.isnan(float(val)):
            continue
        if xi not in tick_set:
            continue
        ax.annotate(_fmt(val, decimals=dec),
            xy=(xi, val), xytext=(0, 6), textcoords="offset points",
            ha="center", va="bottom",
            fontproperties=_fp(max(7, LABEL_SIZE - 2), fw),
            color=col, clip_on=False, annotation_clip=False, zorder=5)


def _annotate_last_line(ax, x, y_arr, th=None, color="#1a1a1a", y_min=None, y_max=None):
    th = th or ChartTheme()
    valid = [(i, v) for i, v in enumerate(y_arr) if not np.isnan(v)]
    if not valid:
        return
    last_xi, last_val = valid[-1]
    fw  = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
    col = _resolve_data_lbl_color(th, color)
    dec = th.data_lbl_decimals
    display_y = last_val
    if getattr(th, "annot_avoid_xaxis", False) and y_min is not None and y_max is not None:
        pad = (y_max - y_min) * 0.03
        display_y = max(display_y, y_min + pad)
    ax.annotate(_fmt(last_val, decimals=dec),
        xy=(last_xi, display_y),
        xytext=(8, 0), textcoords="offset points",
        ha="left", va="center",
        fontproperties=_fp(LABEL_SIZE, fw),
        color=col, clip_on=False, annotation_clip=False, zorder=5)


def _annotate_multi_last(ax, series_list, y_min, y_max, th=None, colors=None, fh=None):
    th     = th or ChartTheme()
    colors = colors or []
    fw     = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
    dec    = th.data_lbl_decimals

    last_points = []
    for i, ya_raw in enumerate(series_list):
        ya    = np.array(ya_raw, dtype=float)
        valid = [(j, v) for j, v in enumerate(ya) if not np.isnan(v)]
        if valid:
            last_points.append((valid[-1][0], valid[-1][1], i))

    if not last_points:
        return

    last_points.sort(key=lambda p: p[1])
    n_pts   = len(last_points)
    fsz     = LABEL_SIZE if n_pts <= 4 else max(7, LABEL_SIZE - 1)
    min_gap = _annot_min_gap(y_min, y_max, fh)

    # Optional: clamp below x-axis
    avoid    = getattr(th, "annot_avoid_xaxis", False)
    axis_pad = (y_max - y_min) * 0.03 if avoid else 0.0
    raw_ys   = [max(p[1], y_min + axis_pad) if avoid else p[1]
                for p in last_points]

    # Bidirectional spread so labels never overlap
    display_ys = _spread_labels(raw_ys, min_gap, y_lo=y_min, y_hi=y_max)

    for (xi, val, si), dy in zip(last_points, display_ys):
        col = colors[si] if (th.data_lbl_color_match_series and si < len(colors)) \
              else th.data_lbl_color
        ax.annotate(_fmt(val, decimals=dec),
            xy=(xi, dy), xytext=(8, 0), textcoords="offset points",
            ha="left", va="center",
            fontproperties=_fp(fsz, fw),
            color=col, clip_on=False, annotation_clip=False, zorder=5)


def _annot_min_gap(y_min, y_max, fh=None):
    """Minimum y-gap between annotation labels in data units.
    When fh (figure height, inches) is supplied we derive the gap from the
    actual rendered font height so it's never smaller than one line of text.
    Falls back to 3 % of the y-range when fh is unknown."""
    y_range = max(y_max - y_min, 1e-9)
    if fh and fh > 0:
        # LABEL_SIZE is in points (already at the render scale factor).
        # Convert: pts → inches → fraction of chart height → data units.
        pad          = y_range * 0.03
        data_per_in  = (y_range + 2 * pad) / fh
        gap          = (LABEL_SIZE / 72.0) * data_per_in * 1.0   # 1× line-height (was 1.5)
        return max(gap, y_range * 0.01)                           # ≥ 1 % floor (was 2 %)
    return y_range * 0.03                                         # fallback (was 6 %)


def _spread_labels(desired_ys, min_gap, y_lo=None, y_hi=None, max_iter=120):
    """Bidirectional iterative label-spread.
    Pushes overlapping labels apart symmetrically (half up, half down),
    then clamps the whole group to [y_lo, y_hi] without compressing gaps.
    Returns a new list of resolved y positions in the same order as input."""
    ys = list(desired_ys)
    n  = len(ys)
    if n <= 1:
        return ys
    for _ in range(max_iter):
        moved = False
        for i in range(n - 1):
            overlap = min_gap - (ys[i + 1] - ys[i])
            if overlap > 1e-9:
                ys[i]     -= overlap / 2
                ys[i + 1] += overlap / 2
                moved = True
        # Clamp bottom: shift entire stack up if it went below y_lo
        if y_lo is not None and ys[0] < y_lo:
            ys = [y + (y_lo - ys[0]) for y in ys]
        # Clamp top: shift entire stack down if it went above y_hi
        if y_hi is not None and ys[-1] > y_hi:
            ys = [y - (ys[-1] - y_hi) for y in ys]
        if not moved:
            break
    return ys


def _spread_tick_annots(ax1, ax2, xi, items, y_min, y_max, from_ax1_fn, th, fh, dec):
    """Place annotations at a single x-tick with vertical collision avoidance.

    items — list of (ax1_y, display_val, target, col) where target is 'ax1' or 'ax2'.
    from_ax1_fn — converts an ax1 y-value back to ax2 data coords (None if no ax2).
    All spreading is done in ax1 space so both axes stay comparable."""
    if not items:
        return
    fw  = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
    fsz = max(7, getattr(th, "data_lbl_size", LABEL_SIZE) - 2)
    # Sort ascending by ax1_y, spread, then map back to original order
    order      = sorted(range(len(items)), key=lambda i: items[i][0])
    raw_sorted = [items[i][0] for i in order]
    min_gap    = _annot_min_gap(y_min, y_max, fh)
    spread     = _spread_labels(raw_sorted, min_gap, y_min, y_max)
    display    = [0.0] * len(items)
    for k, orig_i in enumerate(order):
        display[orig_i] = spread[k]

    for (ax1_y, val, target, col), dy in zip(items, display):
        if target == 'ax1':
            ax1.annotate(_fmt(val, decimals=dec),
                xy=(xi, dy), xytext=(0, 6), textcoords="offset points",
                ha="center", va="bottom", fontproperties=_fp(fsz, fw),
                color=col, clip_on=False, annotation_clip=False, zorder=5)
        else:
            dy2 = from_ax1_fn(dy) if from_ax1_fn else dy
            ax2.annotate(_fmt(val, decimals=dec),
                xy=(xi, dy2), xytext=(0, 6), textcoords="offset points",
                ha="center", va="bottom", fontproperties=_fp(fsz, fw),
                color=col, clip_on=False, annotation_clip=False, zorder=5)


def _auto_annot_pad(fig, ax):
    """Extend ax xlim so all end-of-series annotation text fits inside the
    axes area.  Measures actual rendered text extents via the canvas renderer,
    then extends the x data range by exactly the overhang + 5% breathing room.
    Call AFTER all annotations are placed and BEFORE _save()."""
    try:
        fig.canvas.draw()                              # force layout → real extents
        renderer = fig.canvas.get_renderer()
        ax_bb    = ax.get_window_extent(renderer)
        ax_right = ax_bb.x1                           # axes right edge, display px

        ann_right = ax_right
        for child in ax.get_children():
            if isinstance(child, (matplotlib.text.Text,
                                  matplotlib.text.Annotation)):
                try:
                    bb = child.get_window_extent(renderer)
                    if bb.x1 > ann_right:
                        ann_right = bb.x1
                except Exception:
                    pass

        extra_px = ann_right - ax_right
        if extra_px > 2:                              # >2 px overhang → extend
            xlo, xhi    = ax.get_xlim()
            data_per_px = (xhi - xlo) / ax_bb.width
            ax.set_xlim(xlo, xhi + extra_px * data_per_px * 1.05)
    except Exception:
        pass                                           # silent fallback


def _save(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=DPI, bbox_inches="tight", pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf



def _place_legend(ax, items, th, cfg):
    """Place legend using per-chart cfg fields (Excel) falling back to th (YAML).

    cfg fields (all optional — read from slides-config Excel):
      legend_loc  : "top" | "right" | "none" | "" (= use theme)
      legend_ncol : int > 0  explicit column count   | 0 = auto
      legend_nrow : int > 0  explicit row count       | 0 = auto (computes ncol)
    """
    if not items:
        return

    # Location: per-chart Excel wins over YAML theme
    loc = (getattr(cfg, "legend_loc", "") or "").strip()
    if not loc:
        loc = (getattr(th, "legend_loc", "top") or "top").strip()
    if loc == "none":
        return

    n = len(items)

    # Number of columns
    ncol_cfg = int(getattr(cfg, "legend_ncol", 0) or 0)
    nrow_cfg = int(getattr(cfg, "legend_nrow", 0) or 0)
    if ncol_cfg > 0:
        ncol = ncol_cfg
    elif nrow_cfg > 0:
        ncol = max(1, int(np.ceil(n / nrow_cfg)))
    else:
        # auto: 1 column for right-side; up to 4 per row for top
        ncol = 1 if loc == "right" else min(n, 4)

    leg_fp  = _fp(LEGEND_SIZE)
    lw      = getattr(th, "legend_width",    2.0)
    lh      = getattr(th, "legend_height",   0.7)
    lg      = getattr(th, "legend_label_n_tax_gap", 0.8)
    if loc == "right":
        ax.legend(handles=items, prop=leg_fp,
                  loc="upper left", bbox_to_anchor=(1.02, 1.0),
                  ncol=ncol, frameon=False,
                  handlelength=lw, handleheight=lh, handletextpad=lg,
                  columnspacing=0.6, borderaxespad=0)
    elif loc == "bottom":
        nrows_leg = max(1, int(np.ceil(n / ncol)))
        fig_h_pts = ax.figure.get_figheight() * 72
        leg_pts   = LEGEND_SIZE * 1.6 * nrows_leg
        anchor_y  = -(leg_pts / fig_h_pts) - 0.02
        ax.legend(handles=items, prop=leg_fp,
                  loc="upper center", bbox_to_anchor=(0.5, anchor_y),
                  ncol=ncol, frameon=False,
                  handlelength=lw, handleheight=lh, handletextpad=lg,
                  columnspacing=0.6)
    else:  # "top" (default) or anything unrecognised
        nrows_leg = max(1, int(np.ceil(n / ncol)))
        fig_h_pts = ax.figure.get_figheight() * 72
        leg_pts   = LEGEND_SIZE * 1.6 * nrows_leg
        offset    = getattr(th, "legend_offset", 0.02)
        anchor_y  = 1.0 + (leg_pts / fig_h_pts) + offset
        ax.legend(handles=items, prop=leg_fp,
                  loc="lower center", bbox_to_anchor=(0.5, anchor_y),
                  ncol=ncol, frameon=False,
                  handlelength=lw, handleheight=lh, handletextpad=lg,
                  columnspacing=0.6)

@log_exceptions
def build_chart(cfg, fig_width=5.5, fig_height=3.5, scale=1.0, theme=None):
    global TICK_SIZE, LABEL_SIZE, LEGEND_SIZE, LINE_W, LEG_LINE_W

    if theme is None:
        theme = getattr(cfg, "theme", None) or ChartTheme()
    cfg._theme = theme

    render_w = max(fig_width  * SCALE_FACTOR, 8.0)
    render_h = max(fig_height * SCALE_FACTOR, 5.0)

    MAX_PIXELS  = 80_000_000
    pixel_count = (render_w * DPI) * (render_h * DPI)
    if pixel_count > MAX_PIXELS:
        shrink   = (MAX_PIXELS / pixel_count) ** 0.5
        render_w *= shrink
        render_h *= shrink

    sf = min(render_w / fig_width, render_h / fig_height)

    base_tick   = getattr(theme, "tick_label_size", BASE_TICK)  or BASE_TICK
    base_label  = getattr(theme, "data_lbl_size",      BASE_LABEL) or BASE_LABEL
    base_legend = getattr(theme, "legend_fontsize",  BASE_LEGEND) or BASE_LEGEND
    base_line   = getattr(theme, "line_width",       BASE_LINE) or BASE_LINE

    TICK_SIZE   = int(base_tick   * sf)
    LABEL_SIZE  = int(base_label  * sf)
    LEGEND_SIZE = int(base_legend * sf)
    LINE_W      = max(0.5, base_line * sf)
    leg_lw_cfg  = getattr(theme, "legend_linewidth", 0.0) or 0.0
    LEG_LINE_W  = max(0.5, leg_lw_cfg * sf) if leg_lw_cfg > 0 else LINE_W

    plt.rcParams["font.family"]       = [theme.font_family]
    plt.rcParams["font.weight"]       = getattr(theme, "font_weight", "normal")
    plt.rcParams["xtick.labelsize"]   = TICK_SIZE
    plt.rcParams["ytick.labelsize"]   = TICK_SIZE

    n  = len(cfg.y1_values) if not isinstance(
             cfg.y1_values[0] if cfg.y1_values else 0, list) \
         else len(cfg.y1_values[0])
    x  = np.arange(n)
    y1 = np.array(cfg.y1_values) if not (
             cfg.y1_values and isinstance(cfg.y1_values[0], list)) \
         else cfg.y1_values

    if cfg.is_dual_axis:
        return _build_dual(cfg, x, y1, n, render_w, render_h)
    return _build_single(cfg, x, y1, n, render_w, render_h)


@log_exceptions
def _build_single(cfg, x, y1, n, fw, fh):
    th = getattr(cfg, "_theme", None) or ChartTheme()
    fig, ax = plt.subplots(figsize=(fw, fh))
    if th.transparent:
        fig.patch.set_alpha(0)
        ax.patch.set_alpha(0)

    # Resolve annotation mode:
    #   show_values_all=TRUE  → annotate EVERY data point (highest priority)
    #   show_values_latest    → annotate LATEST point (blank Excel = TRUE by default)
    #   show_values_all=FALSE AND show_values_latest=FALSE → no annotation
    _cfg_all    = getattr(cfg, "show_values_all",    None)          # None if Excel blank
    _cfg_latest = bool(getattr(cfg, "show_values_latest", True))    # True if Excel blank
    # show_values_all: Excel wins → YAML → False
    if _cfg_all is not None:
        _show_all = bool(_cfg_all)
    else:
        _show_all = bool(getattr(th, "show_values_all", False))
    # show_values_all=True always overrides show_values_latest
    _show_latest = _cfg_latest and not _show_all
    # Tick positions: indices where x-axis labels are visible — annotations sync to these
    _tick_positions = [i for i, l in enumerate(cfg.x_labels) if l] if cfg.x_labels else list(range(n))

    if cfg.y1_chart_type == "stacked_bar":
        colors      = getattr(cfg, "stack_colors", None) or ["#FFC125","#EE7600","#FFD39B","#E0E0E0"]
        labels      = getattr(cfg, "stack_labels", None) or []
        bw_base     = th.bar_width if th.bar_width > 0 else _bar_width(n)
        bottom_pos  = np.zeros(n)
        bottom_neg  = np.zeros(n)
        series_list = y1 if isinstance(y1, list) and isinstance(y1[0], list) else [y1]
        lbl_list    = labels if labels else [f"S{i+1}" for i in range(len(series_list))]
        stacked_annot = []
        seg_annot_data = []   # (xi, val, btm_val, color) — one entry per segment per x
        for i, series in enumerate(series_list):
            color = colors[i % len(colors)]
            ys    = np.array(series, dtype=float)
            btm   = np.where(ys >= 0, bottom_pos, bottom_neg)
            ax.bar(x, ys, bottom=btm, color=color, width=bw_base,
                   edgecolor="none", zorder=3, label=lbl_list[i])
            last_val = float(ys[-1])
            last_btm = float(btm[-1])
            stacked_annot.append((int(x[-1]), last_val, last_btm + last_val, color))
            for xi, val, b in zip(x, ys, btm):
                if float(val) != 0:
                    seg_annot_data.append((int(xi), float(val), float(b), color))
            bottom_pos = np.where(ys >= 0, bottom_pos + ys, bottom_pos)
            bottom_neg = np.where(ys <  0, bottom_neg + ys, bottom_neg)
        patches_stk = [Patch(facecolor=colors[i % len(colors)],
                             label=lbl_list[i]) for i in range(len(series_list))]
        _place_legend(ax, patches_stk, th, cfg)
        fw_txt = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
        if _show_all:
            # Annotate every segment individually at its midpoint (tick positions only)
            tick_set = set(_tick_positions)
            for xi, val, btm_val, col in seg_annot_data:
                if xi not in tick_set:
                    continue
                mid_y = btm_val + val / 2
                ax.annotate(_fmt(val, decimals=th.data_lbl_decimals),
                    xy=(xi, mid_y), xytext=(0, 0), textcoords="offset points",
                    ha="center", va="center",
                    fontproperties=_fp(max(7, LABEL_SIZE - 2), fw_txt),
                    color=_resolve_data_lbl_color(th, col),
                    clip_on=True, annotation_clip=False, zorder=5)
        elif _show_latest:
            # annotate last column only, with spread to avoid overlap
            stacked_annot.sort(key=lambda p: p[2])
            raw_ys     = [p[2] for p in stacked_annot]
            min_gap    = _annot_min_gap(cfg.y_min, cfg.y_max, fh)
            display_ys = _spread_labels(raw_ys, min_gap, cfg.y_min, cfg.y_max)
            for (xi, val, _, sc), dy in zip(stacked_annot, display_ys):
                if val == 0:
                    continue
                col = _resolve_data_lbl_color(th, sc)
                ax.annotate(_fmt(val, decimals=th.data_lbl_decimals),
                    xy=(xi, dy), xytext=(8, 0), textcoords="offset points",
                    ha="left", va="center", fontproperties=_fp(LABEL_SIZE, fw_txt),
                    color=col, clip_on=True, annotation_clip=False, zorder=5)

    elif cfg.y1_chart_type == "bar":
        bars, did_all = _draw_bar(ax, x, y1, cfg.bar_color, _show_all,
                                  cfg.y_min, cfg.y_max, n, th=th, tick_positions=_tick_positions)
        if not did_all and _show_latest:
            _annotate_last_bar(ax, bars, y1, th=th, color=cfg.bar_color)
        if cfg.y1_label:
            _place_legend(ax,
                [Patch(facecolor=cfg.bar_color, label=cfg.y1_label)],
                th, cfg)

    elif cfg.y1_chart_type == "line":
        if isinstance(y1, list) and y1 and isinstance(y1[0], list):
            colors   = getattr(cfg, "stack_colors", None) or ["#888888","#FFC125","#E0E0E0","#EE7600"]
            labels   = getattr(cfg, "stack_labels", None) or []
            lbl_list = labels if labels else [f"L{i+1}" for i in range(len(y1))]
            for i, series in enumerate(y1):
                color = colors[i % len(colors)]
                lbl   = lbl_list[i] if i < len(lbl_list) else f"L{i+1}"
                _draw_line(ax, x, series, color, n, th=th, series_idx=i)
                ax.plot([], [], color=color, linewidth=LEG_LINE_W, label=lbl)
            if _show_all:
                # Per-tick spread across all series to avoid overlapping labels
                series_arrs = [np.array([float(v) if v is not None else float("nan") for v in s])
                               for s in y1]
                for xi in _tick_positions:
                    if xi >= n:
                        continue
                    items = []
                    for i, ya in enumerate(series_arrs):
                        val = float(ya[xi]) if not np.isnan(float(ya[xi])) else None
                        if val is not None:
                            col = _resolve_data_lbl_color(th, colors[i % len(colors)])
                            items.append((val, val, 'ax1', col))
                    _spread_tick_annots(ax, None, xi, items, cfg.y_min, cfg.y_max, None, th, fh,
                                        th.data_lbl_decimals)
            elif _show_latest:
                _annotate_multi_last(ax, y1, cfg.y_min, cfg.y_max, th=th, colors=colors, fh=fh)
            lines_lbl = [Line2D([0],[0], color=colors[i % len(colors)],
                                linewidth=LEG_LINE_W, label=lbl_list[i])
                         for i in range(len(y1))]
            _place_legend(ax, lines_lbl, th, cfg)
        else:
            ya = _draw_line(ax, x, y1, cfg.bar_color, n, th=th, series_idx=0)
            if _show_all:
                _annotate_all_line(ax, x, ya, th=th, color=cfg.bar_color, tick_positions=_tick_positions)
            elif _show_latest:
                _annotate_last_line(ax, x, ya, th=th, color=cfg.bar_color,
                                    y_min=cfg.y_min, y_max=cfg.y_max)
            if cfg.y1_label:
                _place_legend(ax,
                    [Line2D([0],[0], color=cfg.bar_color, linewidth=LEG_LINE_W,
                            label=cfg.y1_label)],
                    th, cfg)

    _apply_yaxis(ax, cfg.y_min, cfg.y_max, cfg.y_breaks, th)
    _apply_xaxis(ax, cfg.x_labels, th, cfg)
    _apply_grid(ax, th, cfg)

    # Always hide top and right spines — open / frameless chart look.
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax.spines["left"].set_visible(th.show_y_axis)
    if th.show_y_axis:
        ax.spines["left"].set_linewidth(th.axis_y_linewidth)
        ax.spines["left"].set_color(th.axis_color)

    # Zero-crossing line: only draw when x_axis_at_data_min is NOT explicitly
    # set.  When the user places the x-axis at y_min (e.g. y_min=-100), the
    # y=0 crossing is already visible from the bars/lines — the auto axhline
    # at y=0 is redundant and looks like an unexpected extra line.
    show_zero = (not th.zero_line_with_grid) or th.grid_show
    if show_zero and cfg.y_min < 0 < cfg.y_max \
            and not getattr(th, "x_axis_at_data_min", False):
        ax.axhline(0, color=th.axis_color, linewidth=0.8, zorder=2)

    _auto_annot_pad(fig, ax)   # extend xlim to fit annotation labels exactly
    return _save(fig)


@log_exceptions
def _build_dual(cfg, x, y1, n, fw, fh):
    th = getattr(cfg, "_theme", None) or ChartTheme()
    fig, ax1 = plt.subplots(figsize=(fw, fh))
    if th.transparent:
        fig.patch.set_alpha(0)
        ax1.patch.set_alpha(0)
    ax2 = ax1.twinx()
    if th.transparent:
        ax2.patch.set_alpha(0)
    y2_raw = cfg.y2_values if cfg.y2_values else [0] * n
    bw_base = th.bar_width if th.bar_width > 0 else _bar_width(n)
    bw      = bw_base * 0.85

    # Resolve annotation mode:
    #   show_values_all=TRUE  → annotate EVERY data point (highest priority)
    #   show_values_latest    → annotate LATEST point (blank Excel = TRUE by default)
    #   show_values_all=FALSE AND show_values_latest=FALSE → no annotation
    _cfg_all    = getattr(cfg, "show_values_all",    None)          # None if Excel blank
    _cfg_latest = bool(getattr(cfg, "show_values_latest", True))    # True if Excel blank
    # show_values_all: Excel wins → YAML → False
    if _cfg_all is not None:
        _show_all = bool(_cfg_all)
    else:
        _show_all = bool(getattr(th, "show_values_all", False))
    # show_values_all=True always overrides show_values_latest
    _show_latest = _cfg_latest and not _show_all
    # Tick positions: indices where x-axis labels are visible — annotations sync to these
    _tick_positions = [i for i, l in enumerate(cfg.x_labels) if l] if cfg.x_labels else list(range(n))

    if cfg.y1_chart_type == "stacked_bar":
        colors      = getattr(cfg, "stack_colors", None) or ["#FFC125","#EE7600","#FFD39B","#E0E0E0"]
        labels      = getattr(cfg, "stack_labels", None) or []
        bottom_pos  = np.zeros(n)
        bottom_neg  = np.zeros(n)
        series_list = y1 if isinstance(y1, list) and isinstance(y1[0], list) else [y1]
        lbl_list    = labels if labels else [f"S{i+1}" for i in range(len(series_list))]
        patches     = []
        stacked_annot = []
        seg_annot_data = []   # (xi, val, btm_val, color) — one entry per segment per x
        for i, series in enumerate(series_list):
            color = colors[i % len(colors)]
            ys    = np.array(series, dtype=float)
            btm   = np.where(ys >= 0, bottom_pos, bottom_neg)
            ax1.bar(x, ys, bottom=btm, color=color, width=bw, edgecolor="none", zorder=3)
            last_val = float(ys[-1])
            last_btm = float(btm[-1])
            stacked_annot.append((int(x[-1]), last_val, last_btm + last_val, color))
            for xi, val, b in zip(x, ys, btm):
                if float(val) != 0:
                    seg_annot_data.append((int(xi), float(val), float(b), color))
            bottom_pos = np.where(ys >= 0, bottom_pos + ys, bottom_pos)
            bottom_neg = np.where(ys <  0, bottom_neg + ys, bottom_neg)
            patches.append(Patch(facecolor=color, label=lbl_list[i]))
        stacked_annot.sort(key=lambda p: p[2])
        # Annotations placed later (jointly with y2 line) to avoid cross-axis collision.

    elif cfg.y1_chart_type == "bar":
        # annotate=False → we do joint spread with y2 after y2 is drawn
        bars, did_all = _draw_bar(ax1, x, y1, cfg.bar_color, _show_all,
                                  cfg.y_min, cfg.y_max, n, bw=bw, th=th,
                                  tick_positions=_tick_positions, annotate=False)
        if not did_all and _show_latest:
            _annotate_last_bar(ax1, bars, y1, th=th, color=cfg.bar_color)
        patches = [Patch(facecolor=cfg.bar_color, label=cfg.y1_label or "Primary")]

    else:
        if isinstance(y1, list) and y1 and isinstance(y1[0], list):
            colors   = getattr(cfg, "stack_colors", None) or ["#888888","#FFC125","#E0E0E0","#EE7600"]
            labels   = getattr(cfg, "stack_labels", None) or []
            lbl_list = labels if labels else [f"L{i+1}" for i in range(len(y1))]
            patches  = []
            for i, series in enumerate(y1):
                color = colors[i % len(colors)]
                lbl   = lbl_list[i] if i < len(lbl_list) else f"L{i+1}"
                _draw_line(ax1, x, series, color, n, th=th, series_idx=i)
                patches.append(Line2D([0],[0], color=color, linewidth=LEG_LINE_W, label=lbl))
            if _show_all:
                for i, series in enumerate(y1):
                    color = colors[i % len(colors)]
                    _annotate_all_line(ax1, x, series, th=th, color=color, tick_positions=_tick_positions)
            elif _show_latest:
                _annotate_multi_last(ax1, y1, cfg.y_min, cfg.y_max, th=th, colors=colors, fh=fh)
        else:
            ya = _draw_line(ax1, x, y1, cfg.bar_color, n, th=th, series_idx=0)
            if _show_all:
                _annotate_all_line(ax1, x, ya, th=th, color=cfg.bar_color, tick_positions=_tick_positions)
            elif _show_latest:
                _annotate_last_line(ax1, x, ya, th=th, color=cfg.bar_color, y_min=cfg.y_min, y_max=cfg.y_max)
            patches = [Line2D([0],[0], color=cfg.bar_color, linewidth=LEG_LINE_W,
                              label=cfg.y1_label or "Primary")]

    y2a = _draw_line(ax2, x, y2_raw, cfg.y2_color, n, th=th, series_idx=1)

    if cfg.y1_chart_type == "stacked_bar":
        # Scale functions for cross-axis comparison
        _y1_range = max(cfg.y_max  - cfg.y_min,  1e-9)
        _y2_range = max(cfg.y2_max - cfg.y2_min, 1e-9)
        def _to_ax1(v2):
            return cfg.y_min + (v2 - cfg.y2_min) / _y2_range * _y1_range
        def _from_ax1(v1):
            return cfg.y2_min + (v1 - cfg.y_min) / _y1_range * _y2_range

        if _show_all:
            # ── Per-segment midpoint labels + y2 line at tick positions ──
            # Each stack segment is labeled at its vertical midpoint (natural no-overlap).
            # y2 line is annotated separately via _annotate_all_line.
            tick_set = set(_tick_positions)
            dec = th.data_lbl_decimals
            fw_txt2 = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
            for xi, val, btm_val, col in seg_annot_data:
                if xi not in tick_set:
                    continue
                mid_y = btm_val + val / 2
                ax1.annotate(_fmt(val, decimals=dec),
                    xy=(xi, mid_y), xytext=(0, 0), textcoords="offset points",
                    ha="center", va="center",
                    fontproperties=_fp(max(7, LABEL_SIZE - 2), fw_txt2),
                    color=_resolve_data_lbl_color(th, col),
                    clip_on=True, annotation_clip=False, zorder=5)
            # y2 line — annotate at tick positions, clipped to ax2
            _annotate_all_line(ax2, x, y2a, th=th, color=cfg.y2_color,
                               tick_positions=_tick_positions)
        elif _show_latest:
            # ── Joint cross-axis label spread (last-point only) ───────────────
            # Bar tops live in ax1 coords; the line's last point lives in ax2 coords.
            # Convert everything to ax1 data space, spread together so labels from
            # both series never collide visually, then convert y2's position back.
            _y1_range = max(cfg.y_max  - cfg.y_min,  1e-9)
            _y2_range = max(cfg.y2_max - cfg.y2_min, 1e-9)

            def _to_ax1(v2):
                return cfg.y_min + (v2 - cfg.y2_min) / _y2_range * _y1_range

            def _from_ax1(v1):
                return cfg.y2_min + (v1 - cfg.y_min) / _y1_range * _y2_range

            _y2_valid = [(i, v) for i, v in enumerate(y2a) if not np.isnan(v)]
            _has_y2   = bool(_y2_valid)
            _y2_xi    = _y2_valid[-1][0] if _has_y2 else None
            _y2_val   = _y2_valid[-1][1] if _has_y2 else None

            # All last-point ys in ax1 space (bar tops already sorted asc, then y2)
            _all_raw = [p[2] for p in stacked_annot]
            if _has_y2:
                _all_raw.append(_to_ax1(_y2_val))

            # Sort for spread, remember original indices to map back
            _sort_idx      = sorted(range(len(_all_raw)), key=lambda i: _all_raw[i])
            _sorted_raw    = [_all_raw[i] for i in _sort_idx]
            _min_gap       = _annot_min_gap(cfg.y_min, cfg.y_max, fh)
            _spread_sorted = _spread_labels(_sorted_raw, _min_gap, cfg.y_min, cfg.y_max)

            _display_ax1 = [0.0] * len(_all_raw)
            for _k, _orig_i in enumerate(_sort_idx):
                _display_ax1[_orig_i] = _spread_sorted[_k]

            # Annotate bar tops (ax1 coords)
            _fw = "bold" if th.data_lbl_bold else getattr(th, "font_weight", "normal")
            for (xi, val, _, sc), dy in zip(stacked_annot, _display_ax1[:len(stacked_annot)]):
                if val == 0:
                    continue
                col = _resolve_data_lbl_color(th, sc)
                ax1.annotate(_fmt(val, decimals=th.data_lbl_decimals),
                    xy=(xi, dy), xytext=(8, 0), textcoords="offset points",
                    ha="left", va="center", fontproperties=_fp(LABEL_SIZE, _fw),
                    color=col, clip_on=False, annotation_clip=False, zorder=5)

            # Annotate y2 line — convert spread position from ax1 back to ax2 coords
            if _has_y2:
                _dy_ax2 = _from_ax1(_display_ax1[len(stacked_annot)])
                _col2   = _resolve_data_lbl_color(th, cfg.y2_color)
                ax2.annotate(_fmt(_y2_val, decimals=th.data_lbl_decimals),
                    xy=(_y2_xi, _dy_ax2), xytext=(8, 0), textcoords="offset points",
                    ha="left", va="center", fontproperties=_fp(LABEL_SIZE, _fw),
                    color=_col2, clip_on=False, annotation_clip=False, zorder=5)
    else:
        if _show_all:
            if cfg.y1_chart_type == "bar":
                # Joint per-tick spread: bar tops (ax1) + y2 line (ax2) — no overlaps
                _y1_range = max(cfg.y_max  - cfg.y_min,  1e-9)
                _y2_range = max(cfg.y2_max - cfg.y2_min, 1e-9)
                def _to_ax1_b(v2):
                    return cfg.y_min + (v2 - cfg.y2_min) / _y2_range * _y1_range
                def _from_ax1_b(v1):
                    return cfg.y2_min + (v1 - cfg.y_min) / _y1_range * _y2_range
                tick_set = set(_tick_positions)
                dec = th.data_lbl_decimals
                y1_arr = np.array([float(v) if v is not None else float("nan") for v in y1])
                for xi, bar, bar_val in zip(x, bars, y1_arr):
                    if xi not in tick_set or np.isnan(bar_val):
                        continue
                    items = []
                    col1 = _resolve_data_lbl_color(th, cfg.bar_color)
                    items.append((bar_val, bar_val, 'ax1', col1))
                    y2_here = float(y2a[xi]) if xi < len(y2a) and not np.isnan(float(y2a[xi])) else None
                    if y2_here is not None:
                        col2 = _resolve_data_lbl_color(th, cfg.y2_color)
                        items.append((_to_ax1_b(y2_here), y2_here, 'ax2', col2))
                    _spread_tick_annots(ax1, ax2, xi, items, cfg.y_min, cfg.y_max, _from_ax1_b, th, fh, dec)
            else:
                # line+line dual: per-tick spread across y1 and y2
                _y1_range = max(cfg.y_max  - cfg.y_min,  1e-9)
                _y2_range = max(cfg.y2_max - cfg.y2_min, 1e-9)
                def _to_ax1_l(v2):
                    return cfg.y_min + (v2 - cfg.y2_min) / _y2_range * _y1_range
                def _from_ax1_l(v1):
                    return cfg.y2_min + (v1 - cfg.y_min) / _y1_range * _y2_range
                tick_set = set(_tick_positions)
                dec = th.data_lbl_decimals
                y1_arr = np.array([float(v) if v is not None else float("nan") for v in y1])
                for xi in sorted(tick_set):
                    if xi >= n:
                        continue
                    items = []
                    v1 = float(y1_arr[xi]) if not np.isnan(float(y1_arr[xi])) else None
                    v2 = float(y2a[xi])    if not np.isnan(float(y2a[xi]))    else None
                    if v1 is not None:
                        col1 = _resolve_data_lbl_color(th, cfg.bar_color)
                        items.append((v1, v1, 'ax1', col1))
                    if v2 is not None:
                        col2 = _resolve_data_lbl_color(th, cfg.y2_color)
                        items.append((_to_ax1_l(v2), v2, 'ax2', col2))
                    _spread_tick_annots(ax1, ax2, xi, items, cfg.y_min, cfg.y_max, _from_ax1_l, th, fh, dec)
        elif _show_latest:
            _annotate_last_line(ax2, x, y2a, th=th, color=cfg.y2_color,
                                y_min=cfg.y2_min, y_max=cfg.y2_max)

    _apply_yaxis(ax1, cfg.y_min,  cfg.y_max,  cfg.y_breaks, th, secondary=False)
    _apply_yaxis(ax2, cfg.y2_min, cfg.y2_max, cfg.y_breaks, th, secondary=True)
    if cfg.y_min == cfg.y2_min and cfg.y_max == cfg.y2_max:
        ax2.set_yticklabels([])
    _apply_xaxis(ax1, cfg.x_labels, th, cfg)
    _apply_grid(ax1, th, cfg)
    ax2.grid(False)

    # Always hide top spines on both axes — open / frameless chart look.
    ax1.spines["top"].set_visible(False)
    ax2.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax2.spines["left"].set_visible(False)
    # ax2 has its own bottom spine object (separate from ax1's) that defaults
    # to black and overlays ax1's grid-coloured spine — always hide it.
    ax2.spines["bottom"].set_visible(False)
    ax1.spines["left"].set_visible(th.show_y_axis)
    if th.show_y_axis:
        ax1.spines["left"].set_linewidth(th.axis_y_linewidth)
        ax1.spines["left"].set_color(th.axis_color)
    _show_y2_spine = cfg.show_y2_axis if cfg.show_y2_axis is not None else th.show_y2_axis
    ax2.spines["right"].set_visible(bool(_show_y2_spine))
    if bool(_show_y2_spine):
        ax2.spines["right"].set_linewidth(th.axis_y_linewidth)
        ax2.spines["right"].set_color(th.axis_color)

    show_zero = (not th.zero_line_with_grid) or th.grid_show
    if show_zero and cfg.y_min < 0 < cfg.y_max             and not getattr(th, "x_axis_at_data_min", False):
        ax1.axhline(0, color=th.axis_color, linewidth=0.8, zorder=2)

    if th.y_axis_label:
        ax1.set_ylabel(th.y_axis_label, fontsize=th.y_axis_label_size, rotation=90,
                       labelpad=10, color=th.tick_color)
    if th.y2_axis_label:
        ax2.set_ylabel(th.y2_axis_label, fontsize=th.y2_axis_label_size, rotation=270,
                       labelpad=15, color=th.tick_color)

    y2_line_handle = Line2D([0], [0], color=cfg.y2_color, linewidth=LEG_LINE_W,
                            label=cfg.y2_label or "Secondary")
    _place_legend(ax1, patches + [y2_line_handle], th, cfg)

    _auto_annot_pad(fig, ax1)
    return _save(fig)


@log_exceptions
def build_chart(cfg, fig_width=None, fig_height=None, theme=None):
    """Public entry-point. Attaches theme, resolves geometry, dispatches to
    _build_single or _build_dual."""
    if theme is not None:
        cfg._theme = theme

    x = list(range(len(cfg.x_labels)))
    n = len(x)
    if n == 0:
        return None

    fw = fig_width  or DEFAULT_FW
    fh = fig_height or DEFAULT_FH

    if cfg.is_dual_axis:
        return _build_dual(cfg, x, cfg.y1_values, n, fw, fh)
    else:
        return _build_single(cfg, x, cfg.y1_values, n, fw, fh)
