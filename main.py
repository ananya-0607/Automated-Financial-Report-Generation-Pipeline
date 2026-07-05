"""
main.py  —  run this to generate the PPT

    # resolve all paths from YAML by project name (recommended)
    python main.py --project NIIF_GRIX
    python main.py --project NCAER_MONTHLY_PPT

    # override any individual path after --project
    python main.py --project NIIF_GRIX --output D:/custom/out.pptx

    # fully manual (no YAML)
    python main.py --config input/slides_config_matched.xlsx \
                   --template input/template.pptx \
                   --output output/report.pptx \
                   --ai-titles input/ai_titles.xlsx
"""

import argparse
import os
import sys
import yaml
import pandas as pd
import numpy as np

# NOTE: src.* imports are done lazily inside run_pipeline() / main()
# so that sys.path is already set up by whoever imported this module.

# ── Fallback hardcoded paths (used only when --project is NOT set) ─────
DEFAULT_CONFIG    = "input/slides_config_matched.xlsx"
DEFAULT_TEMPLATE  = "input/template.pptx"
DEFAULT_OUTPUT    = "output/report.pptx"
DEFAULT_AI_TITLES = "input/ai_titles.xlsx"

# ── Default YAML location ─────────────────────────────────────────────
DEFAULT_YAML = r"D:\Santonu\Desktop\ADQvest\Error files\Modified(corr)\R_PPT\ppt_system\ppt_system\PPT_PROJECTS_ENV_CONFIG.yml"

# ── House name appended to every source string ────────────────────────
# Change this to "" if not needed, or to any other org name
HOUSE_NAME = "NIIF Research"


# ── YAML path loader ─────────────────────────────────────────────────
def _load_yaml_paths(yaml_path, project):
    """Return the paths block for one project from the YAML config."""
    with open(yaml_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    projects = cfg["default"]["projects"]
    if project not in projects:
        raise KeyError(
            f"Project '{project}' not found in YAML.\n"
            f"Available: {list(projects.keys())}"
        )
    return projects[project]


def _resolve_paths(args):
    """
    Return (config, template, output, ai_titles) in priority order:
      1. Explicit CLI override (--config / --template / --output / --ai-titles)
      2. YAML block resolved via --project
      3. Hardcoded defaults

    YAML key mapping:
      ai_slide_config → config   (matched slides config — output of match_titles.py)
      ppt_template    → template
      ppt_dir         → output dir  (filename = <project>_report.pptx)
      ai_title_dir    → ai_titles
    """
    yaml_config   = None
    yaml_template = None
    yaml_output   = None
    yaml_titles   = None

    if args.project:
        yaml_path = args.yaml or DEFAULT_YAML
        paths = _load_yaml_paths(yaml_path, args.project)

        yaml_config   = paths.get("ai_slide_config")
        yaml_template = paths.get("ppt_template")
        ppt_dir       = paths.get("ppt_dir")
        if ppt_dir:
            yaml_output = os.path.join(ppt_dir, f"{args.project}_report.pptx")
        yaml_titles   = paths.get("ai_title_dir")

        print(f"Project : {args.project}")
        print(f"  ai_slide_config : {yaml_config}")
        print(f"  ppt_template    : {yaml_template}")
        print(f"  ppt_dir (output): {yaml_output}")
        print(f"  ai_title_dir    : {yaml_titles}")

    config    = args.config    or yaml_config    or DEFAULT_CONFIG
    template  = args.template  or yaml_template  or DEFAULT_TEMPLATE
    output    = args.output    or yaml_output    or DEFAULT_OUTPUT
    ai_titles = args.ai_titles or yaml_titles    or DEFAULT_AI_TITLES

    return config, template, output, ai_titles


# ── AI title matching ─────────────────────────────────────────────────

def _clean_widget(val):
    """Convert widget ID to string, return None if blank/zero/invalid."""
    if val is None: return None
    if isinstance(val, float) and (np.isnan(val) or val == 0): return None
    s = str(val).strip()
    if s in ("", "0", "nan", "None", "NaN"): return None
    return s


def _load_ai_titles(path):
    """
    Load their AI titles Excel and build a lookup:
    frozenset(widget_ids) → (chart_title, chart_source, slide_heading, slide_sub_heading)
    """
    if not path or not os.path.exists(path):
        if path:
            print(f"  [AI titles] File not found: {path} — skipping title matching")
        else:
            print("  [AI titles] No file configured — skipping title matching")
        return {}

    df = pd.read_excel(path)
    lookup = {}

    for (slide, chart), grp in df.groupby(["Slide", "chart"], sort=False):
        # collect all widget IDs in this group
        widget_set = set()
        for w in grp["widget_id"].dropna():
            cw = _clean_widget(w)
            if cw and cw not in ("NO OUTPUT", "Data not Available",
                                  "Report break"):
                try:
                    cw = str(int(float(cw)))
                except:
                    pass
                widget_set.add(cw)

        if not widget_set:
            continue

        # take first non-empty chart_title
        title = ""
        for t in grp["chart_title"].dropna():
            t = str(t).strip()
            if t and t not in ("NO OUTPUT",
                               "Insufficient data to generate title",
                               "nan"):
                title = t
                break

        # take first non-empty chart_source
        source = ""
        for s in grp["chart_source"].dropna():
            s = str(s).strip()
            if s and s not in ("nan", ""):
                source = s
                break

        # take first non-empty slide_heading
        heading = ""
        if "slide_heading" in grp.columns:
            for h in grp["slide_heading"].dropna():
                h = str(h).strip()
                if h and h not in ("nan", "", "Insufficient data to generate a headline."):
                    heading = h
                    break

        # take first non-empty slide_sub_heading
        sub_heading = ""
        if "slide_sub_heading" in grp.columns:
            for sh in grp["slide_sub_heading"].dropna():
                sh = str(sh).strip()
                if sh and sh not in ("nan", ""):
                    sub_heading = sh
                    break

        lookup[frozenset(widget_set)] = (title, source, heading, sub_heading)

    print(f"  [AI titles] Loaded {len(lookup)} chart groups from {path}")
    return lookup


def _get_our_widget_set(chart):
    """Collect all widget IDs from one chart as a frozenset of strings."""
    ids = set()
    for val in [chart.widget_id, chart.widget_id_older,
                chart.y2_widget_id, chart.y2_widget_id_older]:
        cw = _clean_widget(val)
        if cw:
            # split stitched widgets e.g. "1306732+1542853"
            for part in cw.split("+"):
                part = part.strip()
                if part and part != "0":
                    try:
                        part = str(int(float(part)))
                    except:
                        pass
                    ids.add(part)
    # include stacked multi-widget IDs
    sw = getattr(chart, "stack_widgets", "") or ""
    for part in sw.split("|"):
        part = part.strip()
        if part and part != "0":
            try:
                part = str(int(float(part)))
            except:
                pass
            ids.add(part)
    return frozenset(ids)


def _build_subtitle(mode, prefix, x_labels, x_grouping="MY"):
    """
    Build chart subtitle from mode and actual plotted x_labels.
    x_labels: list of label strings from transform (e.g., "Jan-13", "FY13-14")
    """
    if not mode or mode.lower() == "none":
        return prefix or ""

    # get first and last non-empty labels
    visible = [l for l in x_labels if l and str(l).strip()]
    if not visible:
        return prefix or ""

    first_lbl = visible[0]
    last_lbl  = visible[-1]

    def _parse_label(lbl, grouping):
        """Parse x-label back to approximate date."""
        lbl = str(lbl).strip()
        if grouping == "FY" and lbl.startswith("FY"):
            # "FY13-14" → fiscal year ending 2014
            parts = lbl.replace("FY", "").split("-")
            try:
                yr = int(parts[-1])
                yr = yr + 2000 if yr < 100 else yr
                return pd.Timestamp(year=yr, month=3, day=31)
            except:
                return None
        else:
            # "Jan-13", "May-26" etc.
            try:
                return pd.to_datetime(lbl, format="%b-%y")
            except:
                try:
                    return pd.to_datetime(lbl)
                except:
                    return None

    first_dt = _parse_label(first_lbl, x_grouping)
    last_dt  = _parse_label(last_lbl,  x_grouping)

    if not first_dt or not last_dt:
        return prefix or ""

    last_mon = last_dt.strftime("%b")
    last_yy  = last_dt.strftime("%y")
    last_day = str(last_dt.day)

    def _to_fy(dt):
        return dt.year + 1 if dt.month >= 4 else dt.year

    mode_up = mode.strip().upper()
    if mode_up == "CY":
        date_part = (f"CY{first_dt.year}-CY{last_dt.year} "
                     f"({last_mon} '{last_yy})")
    elif mode_up == "FY_DATE":
        date_part = (f"FY{_to_fy(first_dt)}-FY{_to_fy(last_dt)} "
                     f"({last_day} {last_mon} '{last_yy})")
    elif mode_up == "MONTH_ONLY":
        date_part = f"({last_mon} '{last_yy})"
    else:  # FY
        date_part = (f"FY{_to_fy(first_dt)}-FY{_to_fy(last_dt)} "
                     f"({last_mon} '{last_yy})")

    if prefix:
        return f"{prefix}, {date_part}"
    return date_part


def _match_ai_title(chart, ai_lookup):
    """
    Collect ALL widget IDs for this chart → find the AI group
    with the EXACT same set → return its title, source, heading, sub_heading.
    """
    if not ai_lookup:
        return "", "", "", ""
    our_set = _get_our_widget_set(chart)
    if not our_set:
        return "", "", "", ""

    result = ai_lookup.get(our_set)
    if result:
        return result

    return "", "", "", ""


# ── Source building ───────────────────────────────────────────────────

def _build_source(raw1, raw2=None, source_override=None):
    if source_override:
        return source_override

    def get_latest(raw):
        if raw is None: return None, ""
        df  = raw.get("df")
        src = (raw.get("source") or "").strip()
        if df is None or df.empty: return None, src
        try:
            latest = pd.to_datetime(df["Relevant_Date"]).max()
            return latest, src
        except:
            return None, src

    def make_source(parts_str):
        parts = []
        seen  = {"Thurro"}
        for p in parts_str.split(","):
            p = p.strip()
            if p and p not in seen:
                seen.add(p)
                parts.append(p)
        all_parts = ["Thurro"] + parts
        if HOUSE_NAME and HOUSE_NAME not in seen:
            all_parts.append(HOUSE_NAME)
        return "Source: " + ", ".join(all_parts)

    d1, s1 = get_latest(raw1)
    d2, s2 = get_latest(raw2)

    if d2 is None:
        return make_source(s1) if s1 else "Source: Thurro"

    if d1 is not None and d2 is not None:
        if d1 > d2:   return make_source(s1)
        elif d2 > d1: return make_source(s2)
        else:         return make_source(s1 + ", " + s2)

    return make_source(s1) if s1 else "Source: Thurro"


# ── Auto axis ─────────────────────────────────────────────────────────

def _nice_step(data_range, n_breaks=6):
    """Return a clean round step that gives ≈ n_breaks intervals."""
    import math
    if data_range <= 0:
        return 1.0
    raw   = data_range / n_breaks
    mag   = 10 ** math.floor(math.log10(raw))
    for c in [1, 2, 2.5, 5, 10]:
        if c * mag >= raw:
            return c * mag
    return mag * 10


def _auto_axis(chart):
    """
    Auto-compute y_min / y_max from ALL plotted data values.

    Logic (per axis):
      1. Collect every non-null finite value across all series.
      2. Round the raw min/max outward to a clean step so tick labels land
         on nice round numbers (uses _nice_step for 1, 2, 2.5, 5, 10 × magnitude).

    Only runs when y_min or y_max is None (Excel cell was blank).
    Explicitly set values (even 0) are never overwritten.
    """
    import math

    def _all_values(values):
        """Return all non-null finite values across all series."""
        if not values:
            return []
        is_multi = isinstance(values[0], list) if values else False
        result   = []
        src      = values if not is_multi else [v for s in values for v in s]
        for v in src:
            if v is not None:
                try:
                    f = float(v)
                    if not math.isnan(f):
                        result.append(f)
                except (TypeError, ValueError):
                    pass
        return result

    def _compute_nice_range(vals, n_breaks=6):
        """Round raw min/max outward to clean step boundaries."""
        if not vals:
            return None, None
        d_min = min(vals)
        d_max = max(vals)
        if d_min == d_max:
            pad   = abs(d_max) * 0.10 or 1.0
            d_min -= pad
            d_max += pad
        step  = _nice_step(d_max - d_min, n_breaks)
        return (math.floor(d_min / step) * step,
                math.ceil(d_max  / step) * step)

    n_breaks = int(chart.y_breaks) if chart.y_breaks is not None else 6

    # ── Primary axis ──────────────────────────────────────────────────
    if (chart.y_min is None or chart.y_max is None) and chart.y1_values:
        vals = _all_values(chart.y1_values)
        if vals:
            auto_min, auto_max = _compute_nice_range(vals, n_breaks)
            if chart.y_min is None:
                chart.y_min = auto_min
            if chart.y_max is None:
                chart.y_max = auto_max

    # Ensure y_min / y_max are real numbers (last-resort defaults)
    if chart.y_min is None: chart.y_min = 0.0
    if chart.y_max is None: chart.y_max = 100.0
    if chart.y_breaks is None: chart.y_breaks = 6

    # ── Secondary axis (dual-axis charts) ─────────────────────────────
    if chart.is_dual_axis:
        if (chart.y2_min is None or chart.y2_max is None) and chart.y2_values:
            vals2 = _all_values(chart.y2_values)
            if vals2:
                n2 = int(chart.y2_breaks) if chart.y2_breaks is not None else 6
                auto_min2, auto_max2 = _compute_nice_range(vals2, n2)
                if chart.y2_min is None: chart.y2_min = auto_min2
                if chart.y2_max is None: chart.y2_max = auto_max2

        if chart.y2_min    is None: chart.y2_min    = 0.0
        if chart.y2_max    is None: chart.y2_max    = 100.0
        if chart.y2_breaks is None: chart.y2_breaks = 6


# ── Fetch and transform ───────────────────────────────────────────────

def fetch_and_transform(slides, ai_lookup, theme=None):
    from src.data_fetcher import fetch_and_stitch
    from src.transformer  import transform
    from src.chart_theme  import ChartTheme
    from src.ppt_logger   import get_logger
    _log = get_logger()
    theme = theme or ChartTheme()
    from_latest = getattr(theme, "x_label_from_latest", False)
    print("\nFetching data from API...")

    for slide in slides:
        _log.set_context(slide=slide.slide_number, chart="")
        # ── Slide fetch banner ────────────────────────────────────────
        print(f"\n{'='*60}")
        print(f"  SLIDE {slide.slide_number}  |  {len(slide.charts)} chart(s)")
        print(f"  Title : {getattr(slide, 'slide_title', '') or '(none)'}")
        print(f"{'='*60}")

        for ci, chart in enumerate(slide.charts):
            chart_lbl = chr(65 + ci)   # A, B, C, D …
            _log.set_context(chart=chart_lbl)

            # ── primary ───────────────────────────────────────────────
            print(f"\n  ── Chart {chart_lbl} ──────────────────────────────────")
            tag = f"{chart.widget_id}" + (f"+{chart.widget_id_older}"
                  if chart.widget_id_older else "")
            print(f"    [Primary]   widget {tag} ...", end=" ", flush=True)
            try:
                raw1 = fetch_and_stitch(chart.widget_id, chart.widget_id_older)
                print("──────> OK")
            except Exception as e:
                print(f"──────> FAILED ({e})")
                chart.x_labels    = ["N/A"]
                chart.y1_values   = [0]
                chart.chart_title = f"Widget {chart.widget_id} — fetch failed"
                chart.source      = ""
                continue

            t1 = transform(
                df                  = raw1["df"],
                x_grouping          = chart.x_grouping,
                x_interval          = chart.x_interval,
                last_n_months       = chart.last_n_months,
                start_year          = chart.start_year,
                end_year            = chart.end_year,
                growth_type         = chart.growth_type,
                aggregate           = chart.aggregate,
                divisor             = chart.y1_divisor,
                x_label_from_latest = from_latest,
                drop_last_period    = getattr(chart, "drop_last_period", False),
            )

            chart.x_labels    = t1["x_labels"]
            chart._t1_dates   = t1.get("dates", [])   # used for y2 date-alignment
            chart.chart_title = raw1["title"]
            chart._raw1       = raw1

            _sw = getattr(chart, "stack_widgets", "")

            # handle multi-widget charts (stacked bar OR multi-line)
            if _sw:
                from functools import reduce as _reduce
                widget_ids = chart.stack_widgets.split("|")
                labels     = chart.stack_labels or \
                             [f"S{i+1}" for i in range(len(widget_ids))]
                colors     = chart.stack_colors or []
                signs      = chart.stack_signs.split("|") \
                             if getattr(chart, "stack_signs", "") \
                             else ["pos"] * len(widget_ids)

                chart.stack_labels = labels
                chart.stack_colors = colors

                dfs          = []
                used_labels  = []
                # pre-compute column widths for aligned output
                _idx_w = len(str(len(widget_ids)))          # digits in largest index
                _lbl_w = max(
                    (len(labels[i]) if i < len(labels) else len(f"S{i+1}"))
                    for i in range(len(widget_ids))
                ) if widget_ids else 1
                _wid_w = max(len(str(w)) for w in widget_ids) if widget_ids else 1

                for i, wid in enumerate(widget_ids):
                    slbl = labels[i] if i < len(labels) else f"S{i+1}"
                    print(f"    [Stack-{i+1:{_idx_w}}]  "
                          f"{slbl:{_lbl_w}}  "
                          f"widget {str(wid):{_wid_w}} ...",
                          end=" ", flush=True)
                    try:
                        raw_w = fetch_and_stitch(int(wid), 0)
                        print("──────> OK")
                    except Exception as e:
                        print(f"──────> FAILED ({e})")
                        continue
                    df_w = raw_w["df"].copy()
                    vcols = [c for c in df_w.columns if c != "Relevant_Date"]
                    if not vcols:
                        continue
                    # keep only numeric columns before summing
                    num_vcols = [c for c in vcols
                                 if pd.api.types.is_numeric_dtype(df_w[c])]
                    if not num_vcols:
                        # fall back: try to coerce first column
                        num_vcols = [vcols[0]]
                    # multiple value columns → sum them into one
                    if len(num_vcols) > 1:
                        df_w["_value"] = df_w[num_vcols].sum(axis=1)
                    else:
                        df_w["_value"] = pd.to_numeric(df_w[num_vcols[0]], errors="coerce")
                    lbl = labels[i] if i < len(labels) else f"S{i+1}"
                    df_w = df_w[["Relevant_Date", "_value"]].copy()
                    df_w.columns = ["Relevant_Date", lbl]
                    if i < len(signs) and signs[i] == "neg":
                        df_w[lbl] = -df_w[lbl]

                    # ── normalise dates to month-start so all series share
                    #    the same date key — prevents duplicate rows in the
                    #    outer merge when widgets use different day-of-month
                    df_w["Relevant_Date"] = (
                        pd.to_datetime(df_w["Relevant_Date"])
                          .dt.to_period("M")
                          .dt.to_timestamp()          # 1st of each month
                    )
                    # one row per period: keep last value if duplicates exist
                    df_w = (df_w.groupby("Relevant_Date", as_index=False)[lbl]
                                .last())

                    dfs.append(df_w)
                    used_labels.append(lbl)

                if dfs:
                    merged = _reduce(
                        lambda x, y: pd.merge(x, y, on="Relevant_Date",
                                              how="outer"), dfs)
                    merged = merged.sort_values("Relevant_Date").fillna(0)

                    # reset x_labels so stacked data drives them
                    chart.x_labels = []
                    series_list = []
                    for lbl in used_labels:
                        df_col = merged[["Relevant_Date", lbl]].copy()
                        t_col  = transform(
                            df=df_col, x_grouping=chart.x_grouping,
                            x_interval=chart.x_interval,
                            last_n_months=chart.last_n_months,
                            start_year=chart.start_year,
                            end_year=chart.end_year,
                            growth_type=chart.growth_type,
                            aggregate=chart.aggregate,
                            divisor=chart.y1_divisor,
                            x_label_from_latest=from_latest,
                            drop_last_period=getattr(chart, "drop_last_period", False),
                        )
                        series_list.append(t_col["y_values"])
                        if not chart.x_labels:
                            chart.x_labels = t_col["x_labels"]
                    chart.y1_values   = series_list
                    chart.stack_labels = used_labels
                    print(f"    stacked_multi: {len(series_list)} series, "
                          f"{len(series_list[0])} points each")
                else:
                    chart.y1_values = t1["y_values"]

            # handle stacked bar — single widget multi-column
            elif chart.y1_chart_type == "stacked_bar" and \
               chart.notes.startswith("stacked:"):
                parts  = chart.notes.split("::")
                meta   = parts[0].replace("stacked:", "")
                colors = parts[1].split("|") if len(parts) > 1 else []
                labels = meta.split("|")
                chart.stack_labels = labels
                chart.stack_colors = colors

                df_raw   = raw1["df"].copy()
                val_cols = [col for col in df_raw.columns
                            if col != "Relevant_Date"]
                series_list = []
                for col in val_cols[:len(labels)]:
                    df_col = df_raw[["Relevant_Date", col]].copy()
                    t_col  = transform(
                        df=df_col, x_grouping=chart.x_grouping,
                        x_interval=chart.x_interval,
                        last_n_months=chart.last_n_months,
                        start_year=chart.start_year,
                        end_year=chart.end_year,
                        growth_type=chart.growth_type,
                        aggregate=chart.aggregate,
                        divisor=chart.y1_divisor,
                        x_label_from_latest=from_latest,
                        drop_last_period=getattr(chart, "drop_last_period", False),
                    )
                    series_list.append(t_col["y_values"])
                    if not chart.x_labels:
                        chart.x_labels = t_col["x_labels"]
                chart.y1_values = series_list
            else:
                chart.y1_values = t1["y_values"]

            # ── secondary ─────────────────────────────────────────────
            if chart.is_dual_axis:
                tag2 = f"{chart.y2_widget_id}" + (
                    f"+{chart.y2_widget_id_older}"
                    if chart.y2_widget_id_older else "")
                print(f"    [Secondary] widget {tag2} ...", end=" ", flush=True)
                try:
                    raw2 = fetch_and_stitch(chart.y2_widget_id,
                                            chart.y2_widget_id_older)
                    print("──────> OK")
                except Exception as e:
                    print(f"──────> FAILED ({e})")
                    chart.y2_values = [0] * len(chart.y1_values)
                    continue

                df2 = raw2["df"].copy()
                if chart.y2_start_date:
                    cutoff = pd.to_datetime(chart.y2_start_date)
                    df2    = df2[df2["Relevant_Date"] >= cutoff]

                t2 = transform(
                    df                  = df2,
                    x_grouping          = chart.x_grouping,
                    x_interval          = chart.x_interval,
                    last_n_months       = chart.last_n_months,
                    start_year          = chart.start_year,
                    end_year            = chart.end_year,
                    growth_type         = chart.growth_type,
                    aggregate           = chart.aggregate,
                    divisor             = chart.y2_divisor,
                    x_label_from_latest = from_latest,
                    drop_last_period    = getattr(chart, "drop_last_period", False),
                )

                # ── date-keyed left join: align y2 onto y1's date positions ──
                # Key = (year, month) — works for MY, FY (Mar), QF (Mar/Jun/Sep/Dec)
                def _pkey(dt):
                    ts = pd.Timestamp(dt)
                    return (ts.year, ts.month)

                y2_lookup = {}
                for dt, v in zip(t2.get("dates", []), t2["y_values"]):
                    y2_lookup[_pkey(dt)] = v   # later date wins on duplicates

                chart.y2_values = [
                    y2_lookup.get(_pkey(dt))        # None where y2 has no data
                    for dt in chart._t1_dates
                ]

                _auto_axis(chart)

                chart.source = _build_source(
                    raw1            = getattr(chart, "_raw1", None),
                    raw2            = raw2,
                    source_override = chart.source_override or None,
                )

            # ── AI title match ─────────────────────────────────────────
            ai_header, ai_source, ai_heading, ai_sub = \
                _match_ai_title(chart, ai_lookup)
            # always prefer runtime match over pre-matched (may be stale)
            if ai_header:
                chart.chart_header = ai_header
            if ai_source:
                chart.chart_source_ai = ai_source
            # set slide-level heading (same for all charts on same slide)
            if ai_heading and not slide.slide_heading:
                slide.slide_heading = ai_heading
            if ai_sub and not slide.slide_sub_heading:
                slide.slide_sub_heading = ai_sub
            if ai_header:
                print(f"    [Title] '{ai_header[:70]}'")

            # ── Build chart subtitle from actual plotted x_labels ──
            if getattr(chart, "subtitle_mode", "") and chart.x_labels:
                chart.chart_subheader = _build_subtitle(
                    chart.subtitle_mode,
                    getattr(chart, "subtitle_prefix", ""),
                    chart.x_labels,
                    chart.x_grouping,
                )

        # ── Slide fetch done ──────────────────────────────────────────
        _h  = getattr(slide, "slide_heading",     "") or getattr(slide, "slide_title", "") or ""
        _sh = getattr(slide, "slide_sub_heading", "") or ""
        if _h:
            print(f"  Heading    : {_h}")
        if _sh:
            print(f"  Sub-heading: {_sh}")
        print(f"  ✔  Slide {slide.slide_number} data fetched")

    # single axis charts — source and auto axis
    for slide in slides:
        for chart in slide.charts:
            if not chart.source and hasattr(chart, "_raw1"):
                chart.source = _build_source(
                    raw1            = chart._raw1,
                    source_override = chart.source_override or None,
                )
            if not chart.is_dual_axis:
                _auto_axis(chart)

    return slides


def run_pipeline(config, template, output, ai_titles, chart_theme,
                 debug_chart_data=False, debug_chart_data_path="",
                 proj_font=""):
    """
    Callable entry point for project runner scripts.
    Skips argparse — call directly with resolved paths.
    src.* modules are imported here (lazily) so sys.path is already
    configured by the caller before this runs.

    proj_font : path to a .ttf/.otf font file for this project
                (from PPT_PROJECTS_ENV_CONFIG.yml → proj_font).
                Registered with matplotlib so chart_theme's font_family
                can reference it by name (e.g. "Calibri Light").
                Any project can use a different font by setting proj_font
                and chart_theme.font_family in the YAML.
    """
    from src.config_reader import read_config
    from src.ppt_assembler import assemble_ppt
    from src.chart_theme   import theme_from_dict, ChartTheme
    from src.chart_builder import register_font
    from src.ppt_logger    import get_logger

    # ── Initialise central logger ─────────────────────────────────────
    _log      = get_logger()
    _proj_key = getattr(chart_theme, "_project_key", "") if hasattr(chart_theme, "_project_key") else ""
    _log_dir  = os.path.join(os.path.dirname(os.path.abspath(output)), "logs")
    _log.init(log_dir=_log_dir, project=_proj_key or "PPT_PIPELINE")
    _log.info("run_pipeline started", config=config, template=template, output=output)

    if not os.path.exists(config):
        raise FileNotFoundError(f"Config not found: {config}")
    if not os.path.exists(template):
        raise FileNotFoundError(f"Template not found: {template}")

    # Register the project font BEFORE any chart is built so matplotlib
    # can resolve font_family by name (works for Calibri Light, Times New
    # Roman Bold, any installed .ttf/.otf file).
    if proj_font:
        register_font(proj_font)

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    if not isinstance(chart_theme, ChartTheme):
        chart_theme = theme_from_dict(chart_theme)

    print("\nStep 1: Reading config...")
    slides = read_config(config)

    print("\nStep 1b: Loading AI titles...")
    ai_lookup = _load_ai_titles(ai_titles)

    print("\nStep 2: Fetching data...")
    slides = fetch_and_transform(slides, ai_lookup, theme=chart_theme)

    filled = sum(1 for s in slides for c in s.charts if len(c.y1_values) > 0)
    total  = sum(s.chart_count for s in slides)
    print(f"\nCharts with data: {filled}/{total}")

    if filled == 0:
        raise RuntimeError("No data fetched. Check API key in src/data_fetcher.py")

    # ── Optional: save chart data to Excel for inspection ────────────────
    if debug_chart_data and debug_chart_data_path:
        try:
            from src.debug_exporter import export_chart_data
            os.makedirs(os.path.dirname(os.path.abspath(debug_chart_data_path)),
                        exist_ok=True)
            export_chart_data(slides, debug_chart_data_path)
        except Exception as _de:
            print(f"  [debug_exporter] WARNING: could not write debug file — {_de}")

    print("\nStep 3: Building PPT...")
    assemble_ppt(slides, template, output, chart_theme)
    print(f"\nDone → {output}")

    _log.info("run_pipeline completed", output=output)
    print(f"\n📋  Log  → {_log.txt_path}")
    print(f"📋  JSONL → {_log.jsonl_path}")
    _log.close()


def main():
    from src.config_reader import read_config
    from src.ppt_assembler import assemble_ppt
    parser = argparse.ArgumentParser()
    parser.add_argument("--project",   default=None,           help="Project key from YAML (e.g. NIIF_GRIX, NCAER_MONTHLY_PPT)")
    parser.add_argument("--yaml",      default=DEFAULT_YAML,   help="Path to PPT_PROJECTS_ENV_CONFIG.yml")
    parser.add_argument("--config",    default=None,           help="Override: path to slides_config_matched.xlsx")
    parser.add_argument("--template",  default=None,           help="Override: path to PPT template")
    parser.add_argument("--output",    default=None,           help="Override: output .pptx path")
    parser.add_argument("--ai-titles", default=None,           help="Override: path to ai_titles.xlsx")
    args = parser.parse_args()

    config, template, output, ai_titles = _resolve_paths(args)

    if not os.path.exists(config):
        print(f"ERROR: config not found: {config}"); sys.exit(1)
    if not os.path.exists(template):
        print(f"ERROR: template not found: {template}"); sys.exit(1)

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)

    print("\nStep 1: Reading config...")
    slides = read_config(config)

    print("\nStep 1b: Loading AI titles...")
    ai_lookup = _load_ai_titles(ai_titles)

    print("\nStep 2: Fetching data...")
    try:
        slides = fetch_and_transform(slides, ai_lookup)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)

    filled = sum(1 for s in slides for c in s.charts if len(c.y1_values) > 0)
    total  = sum(s.chart_count for s in slides)
    print(f"\nCharts with data: {filled}/{total}")

    if filled == 0:
        print("ERROR: No data fetched. Check API key in src/data_fetcher.py")
        sys.exit(1)

    print("\nStep 3: Building PPT...")
    assemble_ppt(slides, template, output)


if __name__ == "__main__":
    main()
