"""
data_fetcher.py
Fetches data from Thurro API. Supports stitching two widgets into one series.
"""
import requests
import pandas as pd
from .ppt_logger import get_logger, log_exceptions

API_URL = "https://ai.thurro.com/api/getChartData"
API_KEY = ""   # ← replace with your real API key

_cache = {}


@log_exceptions
def fetch_widget(widget_id: int) -> dict:
    if widget_id in _cache:
        return _cache[widget_id]

    headers = {"x-api-key": API_KEY}
    payload = {"widgets": [str(widget_id)]}
    r = requests.post(url=API_URL, headers=headers, data=payload, timeout=30)
    r.raise_for_status()

    result = r.json()["result"][0]
    df = pd.DataFrame.from_dict(result["data"])

    # ── Normalise date column → always "Relevant_Date" ────────────────
    # Some widgets return alternative column names (date, Date, Month, etc.)
    if "Relevant_Date" not in df.columns:
        _DATE_CANDIDATES = [
            "date", "Date", "DATE",
            "month", "Month", "MONTH",
            "period", "Period", "PERIOD",
            "time", "Time", "TIME",
            "observation_date", "ref_date", "Ref_Date",
        ]
        for _c in _DATE_CANDIDATES:
            if _c in df.columns:
                df = df.rename(columns={_c: "Relevant_Date"})
                break
        else:
            # Last resort: pick the first object/datetime-looking column
            for _c in df.columns:
                if pd.api.types.is_object_dtype(df[_c]) or \
                   pd.api.types.is_datetime64_any_dtype(df[_c]):
                    try:
                        pd.to_datetime(df[_c].dropna().iloc[:3], errors="raise")
                        df = df.rename(columns={_c: "Relevant_Date"})
                        break
                    except Exception:
                        pass

    # ── Absolute fallback: no date column found at all ────────────────
    # Synthesise a monthly sequence so the chart can still render.
    if "Relevant_Date" not in df.columns:
        print(f"  ⚠  Widget {widget_id}: no date column found "
              f"(columns: {list(df.columns)}) — synthesising monthly dates")
        df.insert(0, "Relevant_Date",
                  pd.date_range(start="2020-01-01", periods=len(df), freq="MS"))

    if "Relevant_Date" in df.columns:
        df["Relevant_Date"] = pd.to_datetime(df["Relevant_Date"], errors="coerce")
        df = df.sort_values("Relevant_Date").reset_index(drop=True)

    out = {
        "df":        df,
        "title":     result.get("title", f"Widget {widget_id}"),
        "source":    result.get("source", ""),
        "unit":      result.get("chart_unit", ""),
        "frequency": result.get("frequency", "Monthly"),
    }
    _cache[widget_id] = out
    return out


@log_exceptions
def fetch_and_stitch(widget_id: int, older_widget_id: int = 0) -> dict:
    raw = fetch_widget(widget_id)
    if not older_widget_id:
        return raw

    raw_older = fetch_widget(older_widget_id)
    df_new    = raw["df"].copy()
    df_older  = raw_older["df"].copy()

    val_new   = [c for c in df_new.columns   if c != "Relevant_Date"][0]
    val_older = [c for c in df_older.columns if c != "Relevant_Date"][0]
    df_older  = df_older.rename(columns={val_older: val_new})

    min_new  = df_new["Relevant_Date"].min()
    df_older = df_older[df_older["Relevant_Date"] < min_new]

    df_combined = pd.concat([df_older, df_new], ignore_index=True)
    df_combined = df_combined.sort_values("Relevant_Date").reset_index(drop=True)

    return {
        "df":        df_combined,
        "title":     raw["title"],
        "source":    raw["source"],
        "unit":      raw["unit"],
        "frequency": raw["frequency"],
    }


def clear_cache():
    _cache.clear()
