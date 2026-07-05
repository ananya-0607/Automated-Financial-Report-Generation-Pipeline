"""
chart_theme.py — per-project chart styling themes.

Model:
  * A ChartTheme holds the reusable VISUAL identity of a project
    (grid, colours, number formats, legend position, etc.).
  * Themes are loaded from PPT_PROJECTS_ENV_CONFIG.yml under each
    project's `chart_theme:` block. Omitted keys fall back to the
    dataclass defaults below.
  * Resolution rule: cfg ALWAYS WINS. For any styling field, if cfg
    has it explicitly set, that value is used; otherwise the theme
    value is used. See resolve().
"""

from dataclasses import dataclass, field, fields
from typing import Optional, List
import matplotlib.ticker as ticker


# ── Theme definition ──────────────────────────────────────────────────
@dataclass
class ChartTheme:
    # grid
    grid_show:   bool  = True
    grid_color:  str   = "#000000"
    grid_alpha:  float = 0.4
    grid_width:  float = 1.0

    # y-axis number format: "plain" | "comma" | "percent" | "parens"
    y_fmt:          str  = "parens"
    y2_fmt:         str  = "parens"
    y_show_labels:  bool = True
    y2_show_labels: bool = True
    tick_color:     str  = "#000000"

    # series style
    bar_color:    str         = "#C00000"
    line_colors:  List[str]   = field(default_factory=lambda: ["#1f4e79"])
    line_width:   float       = 1.8
    line_style:   str         = "-"     # "-" solid | "--" dashed | ":" dotted
    markers:      bool        = False
    marker_size:  float       = 5.0

    # annotations (end-of-series value labels)
    data_lbl_size:     int  = 11    # font size of end-of-series data labels
    data_lbl_color:    str  = "#1a1a1a"
    data_lbl_decimals: int  = 1

    # legend: "top" | "right" | "none"
    legend_loc: str = "top"

    # background
    transparent: bool = True

    # baseline zero-line tied to grid toggle?  (when False, zero-line
    # always shows if data crosses zero; when True, hidden with grid)
    zero_line_with_grid: bool = False

    # ── Typography & axes ─────────────────────────────────────────────
    # Font family applied to ALL chart text (ticks, labels, legend, annotations).
    # Use any font installed on the system: "Calibri", "Times New Roman", "Arial", etc.
    font_family:      str   = "Calibri"

    # Font weight applied globally to ALL chart text.
    # "light" matches thin fonts like Calibri Light / Source Sans Pro Light.
    # "normal" is the default; "bold" for heavy fonts.
    # Must match the weight baked into the proj_font file.
    font_weight:      str   = "normal"

    # Axis tick label size (x-axis and y-axis numbers/dates).
    # Scaled proportionally to the chart slot size — this is the base size.
    tick_label_size:  int   = 11

    # Show / hide axis spine lines independently.
    # True = visible, False = hidden.
    show_x_axis:      bool  = True   # bottom x-axis line

    # X-axis spine position.
    # False (default) = x-axis line sits at y=0 (matplotlib default).
    # True            = x-axis line sits at y_min (bottom of the chart area).
    #                   Use when y_min is not 0 and you want the axis line
    #                   to coincide with the lowest tick (e.g. y_min=-20 or y_min=10).
    x_axis_at_data_min: bool = False
    show_y_axis:      bool  = False  # left y-axis line (hidden by default — clean look)
    show_y2_axis:     bool  = False  # right y-axis line for dual-axis charts

    # Thickness of the spine lines when shown (pts).
    axis_x_linewidth: float = 0.7
    axis_y_linewidth: float = 0.7

    # Tick mark lengths (pts). 0 = hidden.
    x_tick_length:    float = 5.0   # x-axis tick marks (only drawn when show_x_axis=True)
    y_tick_length:    float = 0.0   # primary y-axis tick marks
    y2_tick_length:   float = 0.0   # secondary y-axis tick marks

    # Tick mark width (pts) — independent of spine/grid linewidth.
    # Controls how thick the short vertical tick lines are.
    x_tick_width:     float = 0.5   # thin by default

    # Gap between tick line and tick label (pts).
    # x_tick_pad controls the space between the x-axis tick mark and the date/label below it.
    x_tick_pad:       float = 4.0   # matplotlib default is 4

    # Tick label bold.
    x_tick_bold:      bool  = False  # bold x-axis tick labels
    y_tick_bold:      bool  = False  # bold primary y-axis tick labels
    y2_tick_bold:     bool  = False  # bold secondary y-axis tick labels

    # X-axis tick label rotation angle (degrees).
    # -1 = auto (0° when ≤9 labels, 45° when >9 labels).
    # Any other value (e.g. 0, 30, 45, 90) overrides auto.
    x_tick_angle:     float = -1.0

    # X-axis direction.
    # False = oldest → newest (left to right, default).
    # True  = newest → oldest (left to right, reverse).
    x_axis_reverse:   bool  = False

    # X-axis label interval direction.
    # False = count from oldest  → oldest always labelled, then every x_interval steps forward.
    # True  = count from latest  → latest always labelled, then every x_interval steps back.
    x_label_from_latest: bool = False

    # Extra right-side x-axis padding for dual-axis charts (fraction of x range).
    # Creates a gap between the end-of-series annotation and the secondary Y-axis tick labels.
    # e.g. 0.08 = extend the x range by 8% on the right side.  0.0 = no padding.
    dual_annot_pad:    float = 0.08

    # Colour for spine lines and zero line.
    axis_color:       str   = "#888888"

    # ── Y-axis rotated unit label ─────────────────────────────────────
    # Rotated label on the primary / secondary y-axis.
    # e.g. "Percent" | "Growth Rate (%)" | "INR Lakh Crore" | "" (empty = hidden)
    y_axis_label:       str  = ""
    y_axis_label_size:  int  = 10
    y2_axis_label:      str  = ""
    y2_axis_label_size: int  = 10

    # ── Bar width ─────────────────────────────────────────────────────
    # Fraction of available slot width: 0.0 = auto (computed from n),
    # positive value overrides (e.g. 0.65).
    bar_width:        float = 0.0

    # ── Show values on every data point ──────────────────────────────
    # True  = annotate EVERY data point (all bars, all line points, all stacked columns).
    # False = annotate only the LAST / latest data point (default).
    # Can be overridden per-chart via the Excel show_values_all column.
    show_values_all:    bool  = False

    # ── Show only latest / last data point label ──────────────────────
    # True  = annotate the LAST/latest data point (default behaviour).
    # False = suppress latest-only label (only meaningful if show_values_all is also False).
    # show_values_all=True overrides this and annotates every point.
    # Can be overridden per-chart via the Excel show_values_latest column.
    show_values_latest: bool  = True

    # ── Annotation x-axis collision avoidance ─────────────────────────
    # True  = clamp end-of-series annotation y so it never falls below
    #         y_min (x-axis line). Prevents label overlapping the axis.
    # False = label sits exactly at the data value (may overlap x-axis).
    annot_avoid_xaxis: bool = False

    # ── Marker shapes ─────────────────────────────────────────────────
    # Shape per series when markers=True.  Matplotlib codes:
    # "o" circle, "s" square, "^" triangle, "D" diamond.  Cycles through list.
    marker_shapes:    List[str] = field(default_factory=lambda: ["o", "s", "^", "D"])

    # ── Per-series line dash styles ───────────────────────────────────
    # If non-empty, overrides line_style per series (cycles).
    # e.g. ["-", "--", ":"] means series 0 solid, 1 dashed, 2 dotted.
    # Empty list [] means all series use global line_style.
    line_styles:      List[str] = field(default_factory=list)

    # ── Annotation styling ────────────────────────────────────────────
    # True = bold annotation text.
    data_lbl_bold:               bool = False
    # True = annotation colour matches series colour (blue label on blue line).
    # False = data_lbl_color is used for all annotations.
    data_lbl_color_match_series: bool = False

    # ── Legend font size ──────────────────────────────────────────────
    # Base size for legend text; scaled to chart size like tick_label_size.
    legend_fontsize:  int   = 10

    # ── Legend offset ─────────────────────────────────────────────────
    # Extra vertical gap between the top of the chart area and the legend
    # (as a fraction of axes height). 0.02 = matplotlib default-ish.
    # Reduce to 0.0 or negative to bring legend closer to the chart.
    legend_offset:        float = 0.02

    # ── Legend handle styling (applies to ALL chart types: line, bar, stacked) ──
    # legend_width   : horizontal length of the swatch in font-size units.
    #                  R/ggplot2 default ≈ 1.5–2.0; matplotlib default = 2.0
    # legend_height  : vertical height of the swatch area in font-size units.
    #                  Controls bar patch height AND line handle height.
    #                  matplotlib default = 0.7; increase for taller swatches.
    # legend_label_n_tax_gap : gap between swatch and label text in font-size units.
    #                          R/ggplot2 default ≈ 0.5; matplotlib default = 0.8
    # legend_linewidth       : thickness of the LINE swatch (pts, before scale).
    #                          0.0 = use line_width (same as chart lines).
    #                          Bar/patch handles ignore this — their size is legend_height.
    legend_width:             float = 2.0
    legend_height:            float = 0.7   # matplotlib default
    legend_label_n_tax_gap:   float = 0.8   # gap between swatch and label text
    legend_linewidth:         float = 0.0   # 0 = match line_width; line charts only


# ── YAML loader ───────────────────────────────────────────────────────
def theme_from_dict(d: Optional[dict]) -> ChartTheme:
    """Build a ChartTheme from a YAML `chart_theme` block (or None)."""
    if not d:
        return ChartTheme()
    valid = {f.name for f in fields(ChartTheme)}
    clean = {k: v for k, v in d.items() if k in valid}
    return ChartTheme(**clean)


def load_theme(project_cfg: dict) -> ChartTheme:
    """
    Given one project's dict from PPT_PROJECTS_ENV_CONFIG.yml, pull its
    `chart_theme` block. Unknown / missing -> defaults.
    """
    return theme_from_dict((project_cfg or {}).get("chart_theme"))


# ── Number-format dispatcher ──────────────────────────────────────────
def make_tick_formatter(kind: str, decimals: int = 1):
    """Return a matplotlib FuncFormatter for the given format kind."""
    def f(v, _):
        if kind == "percent":
            return f"{v:.0f}%"
        if kind == "comma":
            return f"{v:,.0f}"
        if kind == "parens":
            neg = v < 0
            av = abs(v)
            s = f"{av:,.{decimals}f}".rstrip("0").rstrip(".") if decimals else f"{int(av):,}"
            return f"({s})" if neg else s
        # plain
        if v == int(v):
            return f"{int(v):,}"
        return f"{v:,.{decimals}f}"
    return ticker.FuncFormatter(f)


# ── cfg-wins resolver ─────────────────────────────────────────────────
_UNSET = object()  # sentinel for "cfg did not specify this field"


def resolve(theme: ChartTheme, cfg, name: str):
    """
    Return the effective value of styling field `name`.
    cfg ALWAYS WINS: if cfg has the attribute and it is not None / _UNSET,
    use it; otherwise fall back to the theme's value.
    """
    cfg_val = getattr(cfg, name, _UNSET)
    if cfg_val is not _UNSET and cfg_val is not None:
        return cfg_val
    return getattr(theme, name)
