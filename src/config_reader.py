"""
config_reader.py
Reads slides_config.xlsx. One row = one chart.
"""

import pandas as pd
from dataclasses import dataclass, field

REQUIRED_COLS = [
    "slide_number", "slide_title", "layout_name",
    "widget_id", "y1_label", "y1_chart_type", "y1_divisor",
    "bar_color", "y_min", "y_max", "y_breaks",
    "x_grouping", "x_interval", "show_grid"
    # Note: show_values_all / show_values_latest are optional per-chart columns.
    # Old Excel files using the legacy "show_values" column are also supported.
]


@dataclass
class ChartConfig:
    chart_position:    int
    slide_number:      int
    widget_id:         int
    widget_id_older:   int
    Fetching_Method:   str
    y1_label:          str
    y1_chart_type:     str
    y1_divisor:        float
    bar_color:         str
    y_min:             object  # float | None = auto from 1-yr window
    y_max:             object  # float | None = auto from 1-yr window
    y_breaks:          object  # int   | None = auto (default 6)
    y2_widget_id:      int
    y2_widget_id_older:int
    y2_label:          str
    y2_chart_type:     str
    y2_divisor:        float
    y2_color:          str
    y2_min:            object  # float | None = auto from 1-yr window
    y2_max:            object  # float | None = auto from 1-yr window
    y2_breaks:         object  # int   | None = auto (default 6)
    y2_start_date:     str
    x_grouping:        str
    x_interval:        int
    last_n_months:     int
    start_year:        int
    end_year:          int
    growth_type:       str
    aggregate:         str
    show_values_all:    object  # True=all points | None=use YAML
    show_values_latest: object  # True=latest point only | None=use YAML
    show_grid:         bool
    source_override:   str
    notes:             str
    chart_header:      str = ""
    chart_source:      str = ""
    chart_note:        str = ""
    chart_content:     str = ""
    subtitle_mode:     str = ""
    subtitle_prefix:   str = ""
    # stacked bar (multi-widget)
    stack_widgets:     str  = ""
    stack_labels:      list = None
    stack_colors:      list = None
    stack_signs:       str  = ""
    # legend placement (per-chart overrides; empty/0 = use chart_theme YAML default)
    legend_loc:        str  = ""   # "top" | "right" | "none" | "" (use YAML)
    legend_ncol:       int  = 0    # number of legend columns (0 = auto)
    legend_nrow:       int  = 0    # number of legend rows    (0 = auto)
    # x-axis direction (None = use YAML default)
    x_axis_reverse:    object = None  # True=newest→oldest | False=oldest→newest | None=use YAML
    # x-axis tick label angle (None = use YAML default)
    x_tick_angle:      object = None  # float e.g. 0, 45, 90 | None=use YAML
    # show secondary y-axis spine (None = use YAML default)
    show_y2_axis:      object = None  # True | False | None=use YAML
    # drop the most recent data point before plotting
    drop_last_period:  bool   = False
    # filled by main.py
    chart_title:       str  = ""
    chart_subheader:   str  = ""
    source:            str  = ""
    x_labels:          list = field(default_factory=list)
    y1_values:         list = field(default_factory=list)
    y2_values:         list = field(default_factory=list)

    @property
    def is_dual_axis(self):
        return self.y2_widget_id > 0


@dataclass
class SlideConfig:
    slide_number:       int
    slide_title:        str
    layout_name:        str
    charts:             list = field(default_factory=list)
    slide_heading:      str  = ""
    slide_sub_heading:  str  = ""

    @property
    def chart_count(self): return len(self.charts)


def _str(val, default=""):
    try:
        if pd.isna(val): return default
    except: pass
    s = str(val).strip()
    return s if s not in ("nan", "None", "") else default

def _float(val, default=0.0):
    try:    return float(val)
    except: return default

def _int(val, default=0):
    try:    return int(float(str(val)))
    except: return default

def _float_or_none(val):
    """Return float if the cell has a real value, else None (blank/nan = auto-range)."""
    try:
        if pd.isna(val): return None
    except: pass
    s = str(val).strip()
    if s in ("", "nan", "None"): return None
    try: return float(s)
    except: return None

def _bool(val, default=False):
    if isinstance(val, bool): return val
    if isinstance(val, (int, float)): return bool(val)   # 0.0→False, 1.0→True
    return str(val).strip().upper() not in ("FALSE", "0", "NO", "N")


def read_config(excel_path: str) -> list:
    df = pd.read_excel(excel_path, sheet_name="slides")
    df.columns = df.columns.str.strip()

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Excel missing columns: {missing}")

    slides = []
    for sn, grp in df.groupby("slide_number", sort=True):
        first = grp.iloc[0]
        slide = SlideConfig(
            slide_number      = int(sn),
            slide_title       = _str(first["slide_title"]),
            layout_name       = _str(first["layout_name"]),
            slide_heading     = _str(first.get("matched_slide_heading"), ""),
            slide_sub_heading = _str(first.get("matched_slide_sub_heading"), ""),
        )
        for pos, (_, row) in enumerate(grp.iterrows(), 1):
            slide.charts.append(ChartConfig(
                chart_position     = pos,
                slide_number       = int(sn),
                widget_id          = _int(row["widget_id"]),
                widget_id_older    = _int(row.get("widget_id_older"), 0),
                Fetching_Method    = _str(row.get("Fetching_Method"), "API"),
                y1_label           = _str(row.get("y1_label"), ""),
                y1_chart_type      = _str(row.get("y1_chart_type"), "bar").lower(),
                y1_divisor         = _float(row.get("y1_divisor"), 1.0),
                bar_color          = _str(row.get("bar_color"), "#FFC125"),
                y_min              = _float_or_none(row.get("y_min")),
                y_max              = _float_or_none(row.get("y_max")),
                y_breaks           = (_int(row.get("y_breaks"), 6)
                                      if _float_or_none(row.get("y_breaks")) is not None
                                      else None),
                y2_widget_id       = _int(row.get("y2_widget_id"), 0),
                y2_widget_id_older = _int(row.get("y2_widget_id_older"), 0),
                y2_label           = _str(row.get("y2_label"), ""),
                y2_chart_type      = _str(row.get("y2_chart_type"), "line").lower(),
                y2_divisor         = _float(row.get("y2_divisor"), 1.0),
                y2_color           = _str(row.get("y2_color"), "#888888"),
                y2_min             = _float_or_none(row.get("y2_min")),
                y2_max             = _float_or_none(row.get("y2_max")),
                y2_breaks          = (_int(row.get("y2_breaks"), 6)
                                      if _float_or_none(row.get("y2_breaks")) is not None
                                      else None),
                y2_start_date      = _str(row.get("y2_start_date"), ""),
                x_grouping         = _str(row.get("x_grouping"), "MY").upper(),
                x_interval         = _int(row.get("x_interval"), 12),
                last_n_months      = _int(row.get("last_n_months"), 0),
                start_year         = _int(row.get("start_year"), 0),
                end_year           = _int(row.get("end_year"), 0),
                growth_type        = _str(row.get("growth_type"), "none").lower(),
                aggregate          = _str(row.get("aggregate"), "none").lower(),
                show_values_all    = (
                    # Backward-compat: fall back to legacy "show_values" column if new one absent
                    None if pd.isna(row.get("show_values_all",
                                    row.get("show_values", float("nan"))))
                         or str(row.get("show_values_all",
                                row.get("show_values", ""))).strip() in ("", "nan", "None")
                    else _bool(row.get("show_values_all",
                               row.get("show_values")))
                ),
                show_values_latest = (
                    True if pd.isna(row.get("show_values_latest", float("nan")))
                         or str(row.get("show_values_latest", "")).strip() in ("", "nan", "None")
                    else _bool(row.get("show_values_latest"))
                ),
                show_grid          = _bool(row.get("show_grid"), True),
                source_override    = _str(row.get("source_override"), ""),
                notes              = _str(row.get("notes"), ""),
                chart_header       = _str(row.get("matched_chart_title"), ""),
                chart_source       = _str(row.get("matched_chart_source"), ""),
                chart_note         = _str(row.get("chart_note"), ""),
                chart_content      = _str(row.get("chart_content"), ""),
                subtitle_mode      = _str(row.get("subtitle_mode"), ""),
                subtitle_prefix    = _str(row.get("subtitle_prefix"), ""),
                stack_widgets      = _str(row.get("stack_widgets"), ""),
                stack_labels       = _str(row.get("stack_labels"), "").split("|") if _str(row.get("stack_labels"), "") else None,
                stack_colors       = _str(row.get("stack_colors"), "").split("|") if _str(row.get("stack_colors"), "") else None,
                stack_signs        = _str(row.get("stack_signs"), ""),
                legend_loc         = _str(row.get("legend_loc"), ""),
                legend_ncol        = _int(row.get("legend_ncol"), 0),
                legend_nrow        = _int(row.get("legend_nrow"), 0),
                x_axis_reverse     = (
                    None if pd.isna(row.get("x_axis_reverse", float("nan")))
                         or str(row.get("x_axis_reverse", "")).strip() in ("", "nan", "None")
                    else _bool(row.get("x_axis_reverse"))
                ),
                x_tick_angle       = (
                    None if pd.isna(row.get("x_tick_angle", float("nan")))
                         or str(row.get("x_tick_angle", "")).strip() in ("", "nan", "None")
                    else _float(row.get("x_tick_angle"))
                ),
                show_y2_axis       = (
                    None if pd.isna(row.get("show_y2_axis", float("nan")))
                         or str(row.get("show_y2_axis", "")).strip() in ("", "nan", "None")
                    else _bool(row.get("show_y2_axis"))
                ),
                drop_last_period   = _bool(row.get("drop_last_period"), False),
            ))
        slides.append(slide)

    print(f"\nConfig loaded: {len(slides)} slides, "
          f"{sum(s.chart_count for s in slides)} charts")
    for s in slides:
        print(f"  Slide {s.slide_number}: '{s.slide_title}'  "
              f"layout='{s.layout_name}'  {s.chart_count} chart(s)")
    return slides
