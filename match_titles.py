"""
match_titles.py
---------------
Matches our slides_config.xlsx charts with their ai_titles.xlsx
using exact widget ID set matching.

Paths are resolved from PPT_PROJECTS_ENV_CONFIG.yml by project name:
  slides_config    → OUR_EXCEL    (slides config template with widget IDs)
  ai_title_dir     → THEIR_EXCEL  (AI-generated titles from pipeline)
  ai_slide_config  → OUTPUT_FILE  (merged output with matched titles)

Usage:
  python match_titles.py --project NCAER_MONTHLY_PPT
  python match_titles.py --project NIIF_GRIX --config path/to/PPT_PROJECTS_ENV_CONFIG.yml
  python match_titles.py  (falls back to hardcoded input/output dirs)
"""

import argparse
import pandas as pd
import numpy as np
import os
import yaml

# ── Default YAML config path ────────────────────────────────────────────
DEFAULT_CONFIG = r"D:\Santonu\Desktop\ADQvest\Error files\Modified(corr)\R_PPT\ppt_system\ppt_system\PPT_PROJECTS_ENV_CONFIG.yml"

# ── Column names in their Excel ─────────────────────────────────────────
THEIR_WIDGET_COL  = "widget_id"
THEIR_SLIDE_COL   = "Slide"
THEIR_CHART_COL   = "chart"
THEIR_TITLE_COL   = "chart_title"
THEIR_SOURCE_COL  = "chart_source"

# ── Column names in our Excel ────────────────────────────────────────────
OUR_WIDGET_COLS = [
    "widget_id",
    "widget_id_older",
    "y2_widget_id",
    "y2_widget_id_older",
]


# ── YAML path loader ─────────────────────────────────────────────────────
def load_project_paths(config_path, project):
    """Read the YAML and return the paths block for one project."""
    with open(config_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    projects = cfg["default"]["projects"]
    if project not in projects:
        raise KeyError(
            f"Project '{project}' not found in YAML.\n"
            f"Available: {list(projects.keys())}"
        )
    return projects[project]


def resolve_paths(args):
    """
    Return (OUR_EXCEL, THEIR_EXCEL, OUTPUT_FILE) from YAML or fallback defaults.
    YAML keys used:
      slides_config    → OUR_EXCEL
      ai_title_dir     → THEIR_EXCEL
      ai_slide_config  → OUTPUT_FILE
    """
    if args.project:
        config_path = args.config or DEFAULT_CONFIG
        paths = load_project_paths(config_path, args.project)

        our_excel = paths.get("slides_config")
        their_excel = paths.get("ai_title_dir")
        output_file = paths.get("ai_slide_config")

        missing = [k for k, v in [
            ("slides_config", our_excel),
            ("ai_title_dir",  their_excel),
            ("ai_slide_config", output_file),
        ] if not v]
        if missing:
            raise KeyError(
                f"Project '{args.project}' is missing YAML keys: {missing}\n"
                f"Add them to PPT_PROJECTS_ENV_CONFIG.yml."
            )
        return our_excel, their_excel, output_file

    # ── Fallback: old input/output folder behaviour ──────────────────────
    INPUT_DIR  = args.input_dir  or "input"
    OUTPUT_DIR = args.output_dir or "input"
    return (
        os.path.join(INPUT_DIR,  "slides_config.xlsx"),
        os.path.join(INPUT_DIR,  "ai_titles.xlsx"),
        os.path.join(OUTPUT_DIR, "slides_config_matched.xlsx"),
    )


# ── Widget helpers ───────────────────────────────────────────────────────
def clean_widget_id(val):
    """Convert widget ID to string, return None if blank/zero/invalid."""
    if val is None:
        return None
    if isinstance(val, float) and (np.isnan(val) or val == 0):
        return None
    s = str(val).strip()
    if s in ("", "0", "nan", "None", "NaN"):
        return None
    return s


def get_our_widget_set(row):
    """Get frozenset of all valid numeric widget IDs for one chart row."""
    ids = set()
    for col in OUR_WIDGET_COLS:
        val = row.get(col)
        if val is None:
            continue
        for part in str(val).split("+"):
            part = part.strip()
            try:
                cw = str(int(float(part)))
                if cw != "0":
                    ids.add(cw)
            except (ValueError, TypeError):
                continue
    sw = row.get("stack_widgets")
    if sw is not None and str(sw).strip() not in ("", "nan", "None", "0"):
        for part in str(sw).split("|"):
            part = part.strip()
            try:
                cw = str(int(float(part)))
                if cw != "0":
                    ids.add(cw)
            except (ValueError, TypeError):
                continue
    return frozenset(ids)


def build_their_groups(their_df):
    """
    Group their Excel by (Slide, chart).
    Only include groups where at least one valid numeric widget ID exists.
    """
    groups = []

    for (slide, chart), grp in their_df.groupby(
            [THEIR_SLIDE_COL, THEIR_CHART_COL], sort=False):

        widget_set = set()
        for w in grp[THEIR_WIDGET_COL].dropna():
            try:
                cw = str(int(float(str(w).strip())))
                if cw != "0":
                    widget_set.add(cw)
            except (ValueError, TypeError):
                continue

        widget_set = frozenset(widget_set)

        if not widget_set:
            continue

        # take first non-empty title
        title = ""
        for t in grp[THEIR_TITLE_COL].dropna():
            t = str(t).strip()
            if t and t not in ("NO OUTPUT",
                               "Insufficient data to generate title",
                               "nan"):
                title = t
                break

        # take first non-empty source
        source = ""
        for s in grp[THEIR_SOURCE_COL].dropna():
            s = str(s).strip()
            if s and s not in ("nan", ""):
                source = s
                break

        groups.append({
            "widget_set":        widget_set,
            "chart_title":       title,
            "chart_source":      source,
            "slide_heading":     "",
            "slide_sub_heading": "",
            "group_key":         f"Slide={slide}, chart={chart}",
        })

        if "slide_heading" in grp.columns:
            for h in grp["slide_heading"].dropna():
                h = str(h).strip()
                if h and h not in ("nan", "", "Insufficient data to generate a headline."):
                    groups[-1]["slide_heading"] = h
                    break

        if "slide_sub_heading" in grp.columns:
            for sh in grp["slide_sub_heading"].dropna():
                sh = str(sh).strip()
                if sh and sh not in ("nan", ""):
                    groups[-1]["slide_sub_heading"] = sh
                    break

    return groups


def match_chart(our_widget_set, their_groups):
    """Exact set match → return (title, source, slide_heading, slide_sub_heading)."""
    if not our_widget_set:
        return "", "", "", ""

    for group in their_groups:
        if our_widget_set == group["widget_set"]:
            return (group["chart_title"], group["chart_source"],
                    group["slide_heading"], group["slide_sub_heading"])

    return "", "", "", ""


# ── Main ─────────────────────────────────────────────────────────────────
def run(our_excel, their_excel, output_file):
    print("=" * 60)
    print("match_titles.py — Widget ID Matcher")
    print("=" * 60)

    print(f"\nslides_config  : {our_excel}")
    print(f"ai_title_dir   : {their_excel}")
    print(f"ai_slide_config: {output_file}")

    our_df = pd.read_excel(our_excel, sheet_name="slides")
    print(f"\nLoaded slides_config  → {len(our_df)} rows")

    their_df = pd.read_excel(their_excel)
    print(f"Loaded ai_title_dir   → {len(their_df)} rows")

    their_groups = build_their_groups(their_df)
    print(f"\nBuilt {len(their_groups)} chart groups from ai_titles")
    for g in their_groups:
        title_preview = g['chart_title'][:50] + "..." if len(g['chart_title']) > 50 else g['chart_title']
        print(f"  {g['group_key']} → widgets={set(g['widget_set'])} title='{title_preview}'")

    print(f"\nMatching {len(our_df)} rows...")

    matched_titles       = []
    matched_sources      = []
    matched_headings     = []
    matched_sub_headings = []
    match_count          = 0

    for idx, row in our_df.iterrows():
        our_set = get_our_widget_set(row)
        title, source, heading, sub_heading = match_chart(our_set, their_groups)

        matched_titles.append(title)
        matched_sources.append(source)
        matched_headings.append(heading)
        matched_sub_headings.append(sub_heading)

        if title:
            match_count += 1
            print(f"  Row {idx+1}: MATCH — widgets={set(our_set)} → '{title[:60]}'")
        else:
            print(f"  Row {idx+1}: no match — widgets={set(our_set)}")

    our_df["matched_chart_title"]       = matched_titles
    our_df["matched_chart_source"]      = matched_sources
    our_df["matched_slide_heading"]     = matched_headings
    our_df["matched_slide_sub_heading"] = matched_sub_headings

    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
    our_df.to_excel(output_file, index=False, sheet_name="slides")

    print(f"\n{'=' * 60}")
    print(f"Matched:  {match_count} / {len(our_df)} charts")
    print(f"Saved  →  {output_file}")
    print("=" * 60)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Match AI titles to slides config")
    parser.add_argument("--project",    default='NCAER_MONTHLY_PPT',         help="Project key from YAML (e.g. NCAER_MONTHLY_PPT)")
    parser.add_argument("--config",     default=DEFAULT_CONFIG, help="Path to PPT_PROJECTS_ENV_CONFIG.yml")
    parser.add_argument("--input_dir",  default="input",      help="Fallback input folder (used only when --project is not set)")
    parser.add_argument("--output_dir", default="input",      help="Fallback output folder (used only when --project is not set)")
    args = parser.parse_args()

    our_excel, their_excel, output_file = resolve_paths(args)
    run(our_excel, their_excel, output_file)
