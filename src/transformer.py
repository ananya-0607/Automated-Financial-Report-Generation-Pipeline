"""
transformer.py
Filters, aggregates, computes growth, scales, and formats x-axis labels.
All values plotted. Labels control x-axis tick display only.
"""

import pandas as pd
import numpy as np
from fiscalyear import FiscalYear
import fiscalyear
from .ppt_logger import get_logger, log_exceptions


def _fy_year_of(dt):
    return dt.year + 1 if dt.month >= 4 else dt.year

def _fy_label(dt):
    y = _fy_year_of(dt)
    return f"FY{str(y-1)[2:]}-{str(y)[2:]}"

def _qf_label(dt):
    q = ((dt.month - 4) % 12) // 3 + 1
    return f"Q{q} {_fy_label(dt)}"

def _my_label(dt):
    return dt.strftime("%b-%y")

def _format_label(dt, grouping):
    if grouping == "FY": return _fy_label(dt)
    if grouping == "QF": return _qf_label(dt)
    return _my_label(dt)


@log_exceptions
def transform(df: pd.DataFrame,
              x_grouping:        str   = "MY",
              x_interval:        int   = 12,
              last_n_months:     int   = 0,
              start_year:        int   = 0,
              end_year:          int   = 0,
              growth_type:       str   = "none",
              aggregate:         str   = "none",
              divisor:           float = 1.0,
              value_col:         str   = "Value",
              x_label_from_latest: bool = False,
              drop_last_period:  bool  = False) -> dict:

    df = df.copy()

    # Guard: if Relevant_Date is missing entirely, return empty
    if "Relevant_Date" not in df.columns:
        return {"x_labels": [], "y_values": [], "dates": []}

    df["Relevant_Date"] = pd.to_datetime(df["Relevant_Date"], errors="coerce")
    df = df.dropna(subset=["Relevant_Date"])

    # auto-detect value column
    if value_col not in df.columns:
        cols = [c for c in df.columns if c != "Relevant_Date"]
        if not cols:
            return {"x_labels": [], "y_values": []}
        value_col = cols[0]

    df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)
    df = df.sort_values("Relevant_Date").reset_index(drop=True)

    # filter
    if last_n_months and last_n_months > 0:
        cutoff = df["Relevant_Date"].max() - pd.DateOffset(months=last_n_months)
        df = df[df["Relevant_Date"] > cutoff]
    else:
        if start_year and start_year > 0:
            df = df[df["Relevant_Date"].apply(_fy_year_of) >= int(start_year)]
        if end_year and end_year > 0:
            df = df[df["Relevant_Date"].apply(_fy_year_of) <= int(end_year)]

    df = df.reset_index(drop=True)
    if df.empty:
        return {"x_labels": [], "y_values": []}

    # aggregate
    agg = (aggregate or "none").lower().strip()
    yoy_lag = 12
    if agg != "none":
        df = df.set_index("Relevant_Date")
        if agg == "eo_q":
            yoy_lag = 4
            df = df.resample("QE-MAR").last()
        elif agg == "eo_fy":
            yoy_lag = 1
            df = df.resample("YE-MAR").last()
        elif agg == "ytd_q":
            yoy_lag = 4
            df = df.resample("MS").sum().resample("QE-MAR").sum()
        elif agg == "ytd_fy":
            yoy_lag = 1
            # df = df.resample("MS").sum().resample("YE-MAR").sum()
            df = df.resample("MS").sum().resample("A-MAR").sum()
        df = df.reset_index()
        df.columns = ["Relevant_Date", value_col]

    # growth
    gtype = (growth_type or "none").lower().strip()
    if gtype != "none":
        series = df[value_col]
        if gtype == "mom":
            df[value_col] = series.pct_change(1) * 100
        elif gtype == "qoq":
            df[value_col] = series.pct_change(1) * 100
        elif gtype == "yoy":
            df[value_col] = series.pct_change(yoy_lag) * 100
        df = df.dropna(subset=[value_col]).reset_index(drop=True)

    # scale
    if divisor and divisor not in (1, 1.0):
        df[value_col] = df[value_col] / divisor

    # sort oldest first
    df = df.sort_values("Relevant_Date", ascending=True).reset_index(drop=True)

    # drop the most recent data point (e.g. incomplete current month)
    if drop_last_period and len(df) > 1:
        df = df.iloc[:-1].reset_index(drop=True)

    # build x labels
    grouping = (x_grouping or "MY").upper().strip()
    interval = max(1, int(x_interval or 1))

    # Pass 1: collect (row_position, formatted_label, should_label) for every row.
    # Candidates are deduplicated by their formatted label so that x_interval always
    # counts in PERIODS (months / quarters / FYs), not in raw data points.
    # This makes x_interval work correctly for daily, weekly, or any sub-period data:
    # all data points are still plotted; only which rows get a tick label changes.
    seen = set()
    row_meta = []   # list of (row_pos, label_str, should)
    for i, (_, row) in enumerate(df.iterrows()):
        dt  = row["Relevant_Date"]
        lbl = _format_label(dt, grouping)
        should = lbl not in seen
        if should:
            seen.add(lbl)
        row_meta.append((i, lbl, should))

    # Candidate positions = rows where should=True
    candidates = [(pos, lbl) for pos, lbl, s in row_meta if s]
    n_cands    = len(candidates)

    # Which candidate indices (0-based) get a visible label?
    if x_label_from_latest:
        # Count from newest end: newest always labelled, then every interval going back
        labeled_positions = {
            candidates[n_cands - 1 - k][0]
            for k in range(0, n_cands, interval)
        }
    else:
        # Count from oldest end: oldest always labelled, then every interval going forward
        labeled_positions = {
            candidates[k][0]
            for k in range(0, n_cands, interval)
        }

    # Pass 2: build final x_labels list
    x_labels = [
        lbl if (s and pos in labeled_positions) else ""
        for pos, lbl, s in row_meta
    ]

    return {
        "x_labels": x_labels,
        "y_values": df[value_col].tolist(),
        "dates":    df["Relevant_Date"].tolist(),   # raw dates, one per y_value
    }
