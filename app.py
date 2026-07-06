from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =============================================================================
# Convex Asset Trader Experience Board v17 - Streamlit MVP v2
# =============================================================================
# Purpose:
#   Web-app MVP for the PyCharm / Excel v17 trader board.
#   v16 files are NOT used.
#
# Main update in v2:
#   - v4/v5 column-name diagnostics are displayed.
#   - v4 ID-correction contribution columns are auto-detected.
#   - v5 no-correction / imbalance contribution columns are auto-detected.
#   - ID / residual imbalance graph uses uploaded v4/v5 benchmark contributions
#     when available, so the graph is no longer blank when v3 forecast-error
#     columns are missing or named differently.
#
# Required:
#   - v3_da_revenue_YYYY_MM_DD.csv
# Optional:
#   - v4_id_correction_pnl_YYYY_MM_DD.csv
#   - v5_no_correction_imbalance_pnl_YYYY_MM_DD.csv
#   - raw EPEX ID vintage CSV files
# =============================================================================

st.set_page_config(
    page_title="v17 Trader Board Web MVP",
    page_icon="⚡",
    layout="wide",
)


@dataclass(frozen=True)
class AssetConfig:
    label: str
    key: str
    display_name: str
    da_revenue_candidates: Tuple[str, ...]
    forecast_candidates: Tuple[str, ...]
    actual_candidates: Tuple[str, ...]
    error_mwh_candidates: Tuple[str, ...]
    v4_id_candidates: Tuple[str, ...]
    v5_imbalance_candidates: Tuple[str, ...]


ASSETS: Dict[str, AssetConfig] = {
    "Portfolio / DK1 Wind + Solar": AssetConfig(
        label="Portfolio / DK1 Wind + Solar",
        key="portfolio",
        display_name="Portfolio",
        da_revenue_candidates=(
            "portfolio_da_revenue_eur",
            "portfolio_da_revenue_5min_eur",
            "portfolio_da_revenue_eur_5min",
            "da_revenue_eur",
        ),
        forecast_candidates=(
            "portfolio_forecast_mw",
            "portfolio_da_forecast_mw",
            "da_portfolio_forecast_mw",
            "portfolio_forecast_day_ahead_mw",
            "portfolio_da_forecast",
            "da_forecast_mw",
            "forecast_mw",
        ),
        actual_candidates=(
            "portfolio_actual_mw",
            "portfolio_generation_actual_mw",
            "actual_portfolio_mw",
            "portfolio_generation_mw",
            "total_actual_mw",
            "actual_mw",
        ),
        error_mwh_candidates=(
            "portfolio_forecast_error_mwh",
            "portfolio_error_mwh",
            "portfolio_error_5min_mwh",
            "forecast_error_mwh_5min",
            "forecast_error_mwh",
            "error_mwh",
        ),
        v4_id_candidates=(
            "portfolio_id_correction_pnl_eur",
            "portfolio_id_correction_revenue_eur",
            "portfolio_id_correction_value_eur",
            "portfolio_id_revenue_eur",
            "portfolio_id_pnl_eur",
            "id_correction_pnl_eur",
            "id_revenue_eur",
        ),
        v5_imbalance_candidates=(
            "portfolio_no_correction_imbalance_pnl_eur",
            "portfolio_imbalance_pnl_eur",
            "portfolio_imbalance_settlement_eur",
            "portfolio_no_correction_pnl_eur",
            "no_correction_imbalance_pnl_eur",
            "imbalance_settlement_eur",
            "imbalance_pnl_eur",
        ),
    ),
    "Wind-only / DK1 Wind Only": AssetConfig(
        label="Wind-only / DK1 Wind Only",
        key="wind",
        display_name="Wind",
        da_revenue_candidates=(
            "wind_da_revenue_eur",
            "wind_da_revenue_5min_eur",
            "wind_da_revenue_eur_5min",
        ),
        forecast_candidates=(
            "wind_forecast_mw",
            "wind_da_forecast_mw",
            "da_wind_forecast_mw",
            "wind_forecast_day_ahead_mw",
            "wind_da_forecast",
        ),
        actual_candidates=(
            "wind_actual_mw",
            "wind_generation_actual_mw",
            "actual_wind_mw",
            "wind_generation_mw",
        ),
        error_mwh_candidates=(
            "wind_forecast_error_mwh",
            "wind_error_mwh",
            "wind_error_5min_mwh",
            "forecast_error_mwh_5min",
        ),
        v4_id_candidates=(
            "wind_id_correction_pnl_eur",
            "wind_id_correction_revenue_eur",
            "wind_id_correction_value_eur",
            "wind_id_revenue_eur",
            "wind_id_pnl_eur",
        ),
        v5_imbalance_candidates=(
            "wind_no_correction_imbalance_pnl_eur",
            "wind_imbalance_pnl_eur",
            "wind_imbalance_settlement_eur",
            "wind_no_correction_pnl_eur",
        ),
    ),
    "Solar-only / DK1 Solar Only": AssetConfig(
        label="Solar-only / DK1 Solar Only",
        key="solar",
        display_name="Solar",
        da_revenue_candidates=(
            "solar_da_revenue_eur",
            "solar_da_revenue_5min_eur",
            "solar_da_revenue_eur_5min",
        ),
        forecast_candidates=(
            "solar_forecast_mw",
            "solar_da_forecast_mw",
            "da_solar_forecast_mw",
            "solar_forecast_day_ahead_mw",
            "solar_da_forecast",
        ),
        actual_candidates=(
            "solar_actual_mw",
            "solar_generation_actual_mw",
            "actual_solar_mw",
            "solar_generation_mw",
        ),
        error_mwh_candidates=(
            "solar_forecast_error_mwh",
            "solar_error_mwh",
            "solar_error_5min_mwh",
            "forecast_error_mwh_5min",
        ),
        v4_id_candidates=(
            "solar_id_correction_pnl_eur",
            "solar_id_correction_revenue_eur",
            "solar_id_correction_value_eur",
            "solar_id_revenue_eur",
            "solar_id_pnl_eur",
        ),
        v5_imbalance_candidates=(
            "solar_no_correction_imbalance_pnl_eur",
            "solar_imbalance_pnl_eur",
            "solar_imbalance_settlement_eur",
            "solar_no_correction_pnl_eur",
        ),
    ),
}

PRICE_CANDIDATES = {
    "da_price": (
        "da_price_eur_per_mwh",
        "day_ahead_price_eur_per_mwh",
        "dayahead_price_eur_per_mwh",
        "da_price",
        "price_eur_per_mwh",
    ),
    "id3": (
        "id3_benchmark_price_eur_per_mwh",
        "id3_price_eur_per_mwh",
        "id3_benchmark_price",
        "id3_price",
    ),
    "imbalance": (
        "imbalance_price_eur_per_mwh",
        "imbalance_price",
        "balancing_price_eur_per_mwh",
        "regulating_price_eur_per_mwh",
    ),
}

TIME_CANDIDATES = (
    "delivery_time",
    "delivery_timestamp",
    "timestamp",
    "datetime",
    "time",
    "delivery_start",
    "delivery_start_time",
    "delivery_period_start",
    "time_slot",
    "mtu",
)

WINDOWS = {
    "PRE_IDA": ("00:00", "09:55"),
    "IDA3": ("09:55", "14:55"),
    "IDA1": ("14:55", "21:55"),
    "IDA2": ("21:55", "24:00"),
}


# =============================================================================
# Basic helpers
# =============================================================================

def normalize_col(name: object) -> str:
    text = str(name).strip().lower()
    text = text.replace("/", "_").replace("-", "_")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def read_csv_bytes(raw: bytes) -> pd.DataFrame:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=None, engine="python", encoding=encoding)
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise ValueError(f"Could not read CSV: {last_error}")


def read_csv_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        raise ValueError("No file uploaded")
    return read_csv_bytes(uploaded_file.getvalue())


def with_normalized_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [normalize_col(c) for c in out.columns]
    return out


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    cols = set(df.columns)
    for cand in candidates:
        norm = normalize_col(cand)
        if norm in cols:
            return norm
    return None


def numeric_series(df: pd.DataFrame, candidates: Iterable[str], default: float = 0.0) -> Tuple[pd.Series, str]:
    col = first_existing_column(df, candidates)
    if col is None:
        return pd.Series(default, index=df.index, dtype=float), "default"
    return pd.to_numeric(df[col], errors="coerce").fillna(default).astype(float), col


def parse_time_column(df: pd.DataFrame) -> pd.Series:
    col = first_existing_column(df, TIME_CANDIDATES)
    n = len(df)

    if col is None:
        return pd.Series(pd.date_range("2026-01-01", periods=n, freq="5min"), index=df.index)

    raw = df[col]
    parsed = pd.to_datetime(raw, errors="coerce")

    if parsed.isna().mean() > 0.50:
        parsed = pd.to_datetime("2026-01-01 " + raw.astype(str), errors="coerce")

    if parsed.isna().all():
        return pd.Series(pd.date_range("2026-01-01", periods=n, freq="5min"), index=df.index)

    first_valid = parsed.dropna().iloc[0]
    fallback = pd.Series(pd.date_range(first_valid.normalize(), periods=n, freq="5min"), index=df.index)
    return pd.Series(parsed, index=df.index).fillna(fallback)


def hhmm_to_minutes(text: str) -> int:
    if text == "24:00":
        return 24 * 60
    hour, minute = text.split(":")
    return int(hour) * 60 + int(minute)


def time_minutes(ts: pd.Series) -> pd.Series:
    dt = pd.to_datetime(ts)
    return dt.dt.hour * 60 + dt.dt.minute


def window_mask(ts: pd.Series, start_hhmm: str, end_hhmm: str) -> pd.Series:
    mins = time_minutes(ts)
    return (mins >= hhmm_to_minutes(start_hhmm)) & (mins < hhmm_to_minutes(end_hhmm))


def build_capture_factor(df: pd.DataFrame, use_flags: Dict[str, bool], capture_pcts: Dict[str, float], allow_id: bool) -> pd.Series:
    factor = pd.Series(0.0, index=df.index, dtype=float)
    if not allow_id:
        return factor
    for key, (start, end) in WINDOWS.items():
        if use_flags.get(key, False):
            mask = window_mask(df["timestamp"], start, end)
            factor.loc[mask] = capture_pcts.get(key, 0.0) / 100.0
    return factor.clip(lower=0.0, upper=1.0)


def nice_axis_limits(values: pd.Series, pad: float = 1.15) -> Tuple[float, float]:
    clean = pd.to_numeric(values, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if clean.empty:
        return -1.0, 1.0
    vmin = float(clean.min())
    vmax = float(clean.max())
    if vmin == 0 and vmax == 0:
        return -1.0, 1.0
    max_abs = max(abs(vmin), abs(vmax)) * pad
    if vmin >= 0:
        return 0.0, max_abs
    if vmax <= 0:
        return -max_abs, 0.0
    return -max_abs, max_abs


def detect_numeric_column_by_score(
    df: pd.DataFrame,
    asset: AssetConfig,
    exact_candidates: Iterable[str],
    purpose: str,
) -> Tuple[Optional[str], pd.DataFrame]:
    """Return best matching numeric column and a score table for diagnostics."""
    exact = first_existing_column(df, exact_candidates)
    rows: List[Dict[str, object]] = []

    for col in df.columns:
        s = pd.to_numeric(df[col], errors="coerce")
        numeric_count = int(s.notna().sum())
        if numeric_count == 0:
            continue

        score = 0
        reasons: List[str] = []

        if col == exact:
            score += 100
            reasons.append("exact candidate")

        tokens = set(col.split("_"))
        name = col

        if asset.key in tokens or asset.key in name:
            score += 20
            reasons.append(f"asset={asset.key}")

        if purpose == "id":
            if "id" in tokens or "intraday" in name:
                score += 15
                reasons.append("id")
            if "correction" in tokens or "correction" in name:
                score += 12
                reasons.append("correction")
            if any(w in tokens or w in name for w in ["pnl", "revenue", "value", "eur"]):
                score += 8
                reasons.append("eur/pnl/revenue/value")
            if "imbalance" in name:
                score -= 20
                reasons.append("exclude imbalance")
        elif purpose == "imbalance":
            if "imbalance" in tokens or "imbalance" in name:
                score += 18
                reasons.append("imbalance")
            if "no" in tokens or "correction" in tokens or "no_correction" in name:
                score += 6
                reasons.append("no/correction")
            if any(w in tokens or w in name for w in ["pnl", "settlement", "revenue", "value", "eur"]):
                score += 8
                reasons.append("eur/pnl/settlement/value")

        if any(w in tokens or w in name for w in ["price", "mwh", "mw", "volume", "quantity", "forecast", "actual"]):
            score -= 12
            reasons.append("not contribution-like")
        if any(w in tokens or w in name for w in ["check", "gap", "summary"]):
            score -= 20
            reasons.append("check/summary")

        abs_sum = float(s.fillna(0.0).abs().sum())
        total_sum = float(s.fillna(0.0).sum())
        if abs_sum > 0:
            score += 2

        rows.append(
            {
                "column": col,
                "score": score,
                "numeric_count": numeric_count,
                "sum": total_sum,
                "abs_sum": abs_sum,
                "reason": ", ".join(reasons),
            }
        )

    score_df = pd.DataFrame(rows).sort_values(["score", "abs_sum"], ascending=[False, False]) if rows else pd.DataFrame()
    if score_df.empty:
        return None, score_df

    best = str(score_df.iloc[0]["column"])
    if float(score_df.iloc[0]["score"]) <= 0:
        return None, score_df
    return best, score_df


def align_series_to_input(source_df: pd.DataFrame, input_df: pd.DataFrame, value_col: str) -> pd.Series:
    value = pd.to_numeric(source_df[value_col], errors="coerce").fillna(0.0)

    time_col = first_existing_column(source_df, TIME_CANDIDATES)
    if time_col is not None and "timestamp" in input_df.columns:
        ts = parse_time_column(source_df)
        temp = pd.DataFrame({"timestamp": pd.to_datetime(ts), "value": value})
        temp = temp.drop_duplicates("timestamp", keep="last").sort_values("timestamp")
        base = pd.DataFrame({"timestamp": pd.to_datetime(input_df["timestamp"])})
        merged = base.merge(temp, on="timestamp", how="left")
        if merged["value"].notna().sum() > 0:
            return merged["value"].fillna(0.0).astype(float).reset_index(drop=True)

    # Fallback: row order. This is usually correct for v3/v4/v5 288-row daily files.
    aligned = value.reset_index(drop=True).reindex(range(len(input_df))).fillna(0.0)
    return aligned.astype(float)


def load_contribution_file(uploaded_file, input_df: pd.DataFrame, asset: AssetConfig, purpose: str) -> Tuple[pd.Series, Dict[str, object], pd.DataFrame, pd.DataFrame]:
    if uploaded_file is None:
        empty = pd.Series(0.0, index=input_df.index, dtype=float)
        return empty, {"file": None, "selected_column": "not uploaded", "sum": 0.0}, pd.DataFrame(), pd.DataFrame()

    raw_df = read_csv_upload(uploaded_file)
    df = with_normalized_columns(raw_df)

    candidates = asset.v4_id_candidates if purpose == "id" else asset.v5_imbalance_candidates
    selected_col, score_df = detect_numeric_column_by_score(df, asset, candidates, purpose=purpose)

    if selected_col is None:
        empty = pd.Series(0.0, index=input_df.index, dtype=float)
        info = {
            "file": uploaded_file.name,
            "selected_column": "not detected",
            "sum": 0.0,
        }
        return empty, info, df, score_df

    aligned = align_series_to_input(df, input_df, selected_col)
    info = {
        "file": uploaded_file.name,
        "selected_column": selected_col,
        "sum": float(aligned.sum()),
        "abs_sum": float(aligned.abs().sum()),
        "rows": int(len(aligned)),
    }
    return aligned, info, df, score_df


# =============================================================================
# Raw EPEX helpers
# =============================================================================

def find_price_column(df: pd.DataFrame) -> Optional[str]:
    explicit = first_existing_column(
        df,
        (
            "price_eur_per_mwh",
            "price",
            "vwap",
            "weighted_average_price",
            "id_price",
            "id3_price",
        ),
    )
    if explicit:
        return explicit

    for col in df.columns:
        if "price" in col and pd.to_numeric(df[col], errors="coerce").notna().sum() > 0:
            return col
    return None


def align_price_to_input(raw_df: pd.DataFrame, input_df: pd.DataFrame) -> Optional[pd.Series]:
    if raw_df.empty:
        return None

    raw = with_normalized_columns(raw_df)
    price_col = find_price_column(raw)
    if price_col is None:
        return None

    price = pd.to_numeric(raw[price_col], errors="coerce")
    time_col = first_existing_column(raw, TIME_CANDIDATES)

    if time_col is not None:
        ts = parse_time_column(raw)
        temp = pd.DataFrame({"timestamp": pd.to_datetime(ts), "price": price}).dropna(subset=["price"])
        if not temp.empty:
            temp = temp.drop_duplicates("timestamp", keep="last").sort_values("timestamp")
            base = pd.DataFrame({"timestamp": pd.to_datetime(input_df["timestamp"])})
            merged = base.merge(temp, on="timestamp", how="left")
            return merged["price"].ffill().bfill().fillna(0.0)

    if len(price) >= len(input_df):
        return price.iloc[: len(input_df)].reset_index(drop=True).ffill().bfill().fillna(0.0)
    if len(price) > 0:
        return price.reindex(range(len(input_df))).ffill().bfill().fillna(float(price.dropna().iloc[0]) if price.dropna().size else 0.0)
    return None


def raw_price_key_from_filename(name: str) -> Optional[str]:
    n = name.lower()
    if "ida3" in n:
        return "IDA3 vintage"
    if "ida1" in n:
        return "IDA1 vintage"
    if "ida2" in n:
        return "IDA2 vintage"
    if "0300" in n or "03_00" in n:
        return "03:00 ID vintage"
    if "0600" in n or "06_00" in n:
        return "06:00 ID vintage"
    if "0900" in n or "09_00" in n:
        return "09:00 ID vintage"
    return None


def load_raw_vintage_prices(files, input_df: pd.DataFrame, fallback: pd.Series) -> Tuple[Dict[str, pd.Series], pd.DataFrame]:
    result: Dict[str, pd.Series] = {
        "ID3 Benchmark Price": fallback.copy(),
        "03:00 ID vintage": fallback.copy(),
        "06:00 ID vintage": fallback.copy(),
        "09:00 ID vintage": fallback.copy(),
        "IDA3 vintage": fallback.copy(),
        "IDA1 vintage": fallback.copy(),
        "IDA2 vintage": fallback.copy(),
    }
    rows: List[Dict[str, object]] = []

    if not files:
        return result, pd.DataFrame(rows)

    for uploaded in files:
        key = raw_price_key_from_filename(uploaded.name)
        if key is None:
            rows.append({"file": uploaded.name, "mapped_to": "ignored", "status": "filename not recognized"})
            continue
        try:
            df = read_csv_upload(uploaded)
            aligned = align_price_to_input(df, input_df)
            if aligned is not None:
                result[key] = aligned.astype(float).reset_index(drop=True)
                rows.append({"file": uploaded.name, "mapped_to": key, "status": "ok", "avg_price": float(result[key].mean())})
            else:
                rows.append({"file": uploaded.name, "mapped_to": key, "status": "no price column detected"})
        except Exception as exc:  # noqa: BLE001
            rows.append({"file": uploaded.name, "mapped_to": key, "status": f"error: {exc}"})

    return result, pd.DataFrame(rows)


# =============================================================================
# Build input and simulation
# =============================================================================

@st.cache_data(show_spinner=False)
def build_input_data(v3_bytes: bytes, asset_label: str) -> Tuple[pd.DataFrame, Dict[str, str], pd.DataFrame]:
    asset = ASSETS[asset_label]
    raw_df = read_csv_bytes(v3_bytes)
    df = with_normalized_columns(raw_df)

    timestamp = parse_time_column(df)

    da_revenue, da_rev_col = numeric_series(df, asset.da_revenue_candidates, default=0.0)
    forecast_mw, forecast_col = numeric_series(df, asset.forecast_candidates, default=np.nan)
    actual_mw, actual_col = numeric_series(df, asset.actual_candidates, default=np.nan)
    error_mwh, error_col = numeric_series(df, asset.error_mwh_candidates, default=np.nan)

    da_price, da_price_col = numeric_series(df, PRICE_CANDIDATES["da_price"], default=np.nan)
    id3_price, id3_col = numeric_series(df, PRICE_CANDIDATES["id3"], default=np.nan)
    imbalance_price, imb_col = numeric_series(df, PRICE_CANDIDATES["imbalance"], default=np.nan)

    # Defensive reconstruction where possible.
    if forecast_mw.isna().all() and da_price.notna().any() and da_revenue.notna().any():
        safe_price = da_price.replace(0, np.nan)
        forecast_mw = (da_revenue * 12.0 / safe_price).replace([np.inf, -np.inf], np.nan)
        forecast_col = "reconstructed_from_da_revenue_and_da_price"

    if error_mwh.isna().all() and forecast_mw.notna().any() and actual_mw.notna().any():
        error_mwh = (actual_mw - forecast_mw) / 12.0
        error_col = "reconstructed_actual_minus_forecast"

    if actual_mw.isna().all() and forecast_mw.notna().any() and error_mwh.notna().any():
        actual_mw = forecast_mw + error_mwh * 12.0
        actual_col = "reconstructed_from_forecast_and_error"

    da_price = da_price.ffill().bfill().fillna(0.0)
    id3_price = id3_price.ffill().bfill().fillna(0.0)
    imbalance_price = imbalance_price.ffill().bfill().fillna(id3_price).fillna(0.0)

    out = pd.DataFrame(
        {
            "timestamp": pd.to_datetime(timestamp),
            "time_label": pd.to_datetime(timestamp).dt.strftime("%H:%M"),
            "da_forecast_mw": forecast_mw.astype(float),
            "actual_mw": actual_mw.astype(float),
            "forecast_error_mwh": error_mwh.astype(float),
            "da_price_eur_per_mwh": da_price.astype(float),
            "id3_benchmark_price_eur_per_mwh": id3_price.astype(float),
            "imbalance_price_eur_per_mwh": imbalance_price.astype(float),
            "da_revenue_eur_5min": da_revenue.astype(float),
        }
    )

    out = out.dropna(subset=["timestamp"]).sort_values("timestamp").drop_duplicates("timestamp", keep="last")
    if len(out) > 288:
        out = out.iloc[:288].copy()
    out = out.reset_index(drop=True)

    out["forecast_error_mwh"] = out["forecast_error_mwh"].fillna(0.0)
    out["da_forecast_mw"] = out["da_forecast_mw"].fillna(0.0)
    out["actual_mw"] = out["actual_mw"].fillna(out["da_forecast_mw"])

    sources = {
        "DA revenue column": da_rev_col,
        "Forecast column": forecast_col,
        "Actual column": actual_col,
        "Forecast error column": error_col,
        "DA price column": da_price_col,
        "ID3 price column": id3_col,
        "Imbalance price column": imb_col,
    }
    return out, sources, df


def run_simulation(
    input_df: pd.DataFrame,
    vintage_prices: Dict[str, pd.Series],
    strategy_mode: str,
    da_sold_pct: float,
    use_flags: Dict[str, bool],
    capture_pcts: Dict[str, float],
    price_sources: Dict[str, str],
    contribution_source_mode: str,
    v4_id_base: pd.Series,
    v5_imbalance_base: pd.Series,
) -> Tuple[pd.DataFrame, Dict[str, float]]:
    df = input_df.copy()

    da_multiplier = 0.0 if strategy_mode == "ID Only" else da_sold_pct / 100.0
    allow_id = strategy_mode != "DA Only"
    capture_factor = build_capture_factor(df, use_flags, capture_pcts, allow_id=allow_id)

    df["da_revenue_eur"] = df["da_revenue_eur_5min"] * da_multiplier
    df["capture_factor"] = capture_factor

    use_uploaded_v4 = contribution_source_mode in ("Auto: prefer uploaded v4/v5", "Use uploaded v4/v5 benchmark columns") and v4_id_base.abs().sum() > 0
    use_uploaded_v5 = contribution_source_mode in ("Auto: prefer uploaded v4/v5", "Use uploaded v4/v5 benchmark columns") and v5_imbalance_base.abs().sum() > 0

    if use_uploaded_v4:
        df["id_revenue_eur"] = v4_id_base.reindex(df.index).fillna(0.0).astype(float) * capture_factor
        df["id_revenue_source"] = "uploaded_v4_scaled_by_capture"
        # Approximate ID trade MWh only for KPI display when price is available.
        safe_price = df["id3_benchmark_price_eur_per_mwh"].replace(0, np.nan)
        df["total_id_trade_mwh"] = (df["id_revenue_eur"] / safe_price).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    else:
        total_trade = pd.Series(0.0, index=df.index)
        total_id_revenue = pd.Series(0.0, index=df.index)
        for key, (start, end) in WINDOWS.items():
            mask = window_mask(df["timestamp"], start, end)
            capture = (capture_pcts.get(key, 0.0) / 100.0) if (allow_id and use_flags.get(key, False)) else 0.0
            trade_col = f"{key.lower()}_trade_mwh"
            price_col = f"{key.lower()}_price_eur_per_mwh"
            rev_col = f"{key.lower()}_id_revenue_eur"

            source_name = price_sources.get(key, "ID3 Benchmark Price")
            price = vintage_prices.get(source_name, df["id3_benchmark_price_eur_per_mwh"])
            price = pd.Series(price).reset_index(drop=True).reindex(df.index).ffill().bfill().fillna(0.0)

            df[trade_col] = np.where(mask, df["forecast_error_mwh"] * capture, 0.0)
            df[price_col] = price
            df[rev_col] = df[trade_col] * df[price_col]

            total_trade = total_trade + df[trade_col]
            total_id_revenue = total_id_revenue + df[rev_col]

        df["total_id_trade_mwh"] = total_trade
        df["id_revenue_eur"] = total_id_revenue
        df["id_revenue_source"] = "model_from_forecast_error_and_price"

    df["remaining_imbalance_mwh"] = df["forecast_error_mwh"] - df["total_id_trade_mwh"]

    if use_uploaded_v5:
        # v5 is the no-correction imbalance benchmark. For MVP simulation, residual imbalance
        # is reduced by the capture factor. This is an approximation, but it gives an intuitive
        # web control before the full Excel formula engine is migrated.
        df["imbalance_settlement_eur"] = v5_imbalance_base.reindex(df.index).fillna(0.0).astype(float) * (1.0 - capture_factor)
        df["imbalance_source"] = "uploaded_v5_scaled_by_remaining_capture"
    else:
        df["imbalance_settlement_eur"] = df["remaining_imbalance_mwh"] * df["imbalance_price_eur_per_mwh"]
        df["imbalance_source"] = "model_remaining_mwh_times_imbalance_price"

    df["total_revenue_eur"] = df["da_revenue_eur"] + df["id_revenue_eur"] + df["imbalance_settlement_eur"]

    if v5_imbalance_base.abs().sum() > 0:
        df["no_id_benchmark_eur"] = df["da_revenue_eur"] + v5_imbalance_base.reindex(df.index).fillna(0.0).astype(float)
    else:
        df["no_id_benchmark_eur"] = df["da_revenue_eur"] + df["forecast_error_mwh"] * df["imbalance_price_eur_per_mwh"]

    df["id_strategy_value_eur"] = df["total_revenue_eur"] - df["no_id_benchmark_eur"]
    df["id_plus_imbalance_eur"] = df["id_revenue_eur"] + df["imbalance_settlement_eur"]

    kpi = {
        "DA revenue EUR": float(df["da_revenue_eur"].sum()),
        "ID revenue EUR": float(df["id_revenue_eur"].sum()),
        "Imbalance settlement EUR": float(df["imbalance_settlement_eur"].sum()),
        "Total revenue EUR": float(df["total_revenue_eur"].sum()),
        "No-ID benchmark EUR": float(df["no_id_benchmark_eur"].sum()),
        "ID strategy value EUR": float(df["id_strategy_value_eur"].sum()),
        "Total forecast error MWh": float(df["forecast_error_mwh"].sum()),
        "Total executed ID trade MWh": float(df["total_id_trade_mwh"].sum()),
        "Remaining imbalance MWh": float(df["remaining_imbalance_mwh"].sum()),
    }
    return df, kpi


# =============================================================================
# Charts and downloads
# =============================================================================

def make_da_chart(df: pd.DataFrame) -> go.Figure:
    pos = df["da_revenue_eur"].clip(lower=0)
    neg = df["da_revenue_eur"].clip(upper=0)
    y_min, y_max = nice_axis_limits(df["da_revenue_eur"])

    fig = go.Figure()
    fig.add_bar(x=df["time_label"], y=pos, name="Positive DA Revenue EUR")
    fig.add_bar(x=df["time_label"], y=neg, name="Negative DA Revenue EUR")
    fig.update_layout(
        title="DA Revenue EUR - 5-min contribution",
        barmode="relative",
        height=430,
        margin=dict(l=40, r=20, t=60, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(tickmode="array", tickvals=df["time_label"].iloc[::12], tickangle=-45),
        yaxis=dict(range=[y_min, y_max], title="EUR per 5-min interval"),
    )
    return fig


def make_id_imbalance_chart(df: pd.DataFrame) -> go.Figure:
    id_pos = df["id_revenue_eur"].clip(lower=0)
    id_neg = df["id_revenue_eur"].clip(upper=0)
    imb_pos = df["imbalance_settlement_eur"].clip(lower=0)
    imb_neg = df["imbalance_settlement_eur"].clip(upper=0)
    total = df["id_revenue_eur"] + df["imbalance_settlement_eur"]
    y_min, y_max = nice_axis_limits(pd.concat([id_pos, id_neg, imb_pos, imb_neg, total]))

    fig = go.Figure()
    fig.add_bar(x=df["time_label"], y=id_pos, name="Positive ID contribution")
    fig.add_bar(x=df["time_label"], y=id_neg, name="Negative ID contribution")
    fig.add_bar(x=df["time_label"], y=imb_pos, name="Positive residual imbalance settlement")
    fig.add_bar(x=df["time_label"], y=imb_neg, name="Negative residual imbalance settlement")
    fig.update_layout(
        title="ID Revenue / Residual Imbalance Contribution - 5-min",
        barmode="relative",
        height=500,
        margin=dict(l=40, r=20, t=90, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(tickmode="array", tickvals=df["time_label"].iloc[::12], tickangle=-45),
        yaxis=dict(range=[y_min, y_max], title="EUR per 5-min interval"),
    )
    return fig


def make_excel_download(input_df: pd.DataFrame, settlement_df: pd.DataFrame, kpi: Dict[str, float], diagnostics: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([kpi]).T.rename(columns={0: "value"}).to_excel(writer, sheet_name="KPI")
        input_df.to_excel(writer, sheet_name="Input_Data", index=False)
        settlement_df.to_excel(writer, sheet_name="Settlement", index=False)
        diagnostics.to_excel(writer, sheet_name="Diagnostics", index=False)
    return output.getvalue()


def fmt_eur(value: float) -> str:
    return f"{value:,.0f} EUR"


def fmt_mwh(value: float) -> str:
    return f"{value:,.1f} MWh"


# =============================================================================
# UI
# =============================================================================

st.title("⚡ Convex Asset Trader Experience Board v17 - Web MVP")
st.caption("Python / Streamlit version. Uses v3/v4/v5 and raw EPEX vintage files. v16 files are not used.")

with st.sidebar:
    st.header("1. Upload CSV files")
    v3_file = st.file_uploader("Required: v3_da_revenue_YYYY_MM_DD.csv", type=["csv"])
    v4_file = st.file_uploader("Optional: v4_id_correction_pnl_YYYY_MM_DD.csv", type=["csv"])
    v5_file = st.file_uploader("Optional: v5_no_correction_imbalance_pnl_YYYY_MM_DD.csv", type=["csv"])
    epex_files = st.file_uploader(
        "Optional: raw EPEX ID vintage CSVs",
        type=["csv"],
        accept_multiple_files=True,
    )

    st.header("2. Asset")
    asset_label = st.selectbox("Asset", list(ASSETS.keys()))
    asset = ASSETS[asset_label]

    st.header("3. Strategy")
    strategy_mode = st.selectbox("Strategy Mode", ["DA + ID Correction", "DA Only", "ID Only"])
    da_sold_pct = st.slider("DA Sold % of Forecast", 0.0, 100.0, 100.0, 1.0)

    st.header("4. Contribution source")
    contribution_source_mode = st.selectbox(
        "ID / imbalance graph source",
        [
            "Auto: prefer uploaded v4/v5",
            "Use uploaded v4/v5 benchmark columns",
            "Model from forecast error and prices only",
        ],
        index=0,
    )

    st.header("5. ID capture")
    price_options = [
        "ID3 Benchmark Price",
        "03:00 ID vintage",
        "06:00 ID vintage",
        "09:00 ID vintage",
        "IDA3 vintage",
        "IDA1 vintage",
        "IDA2 vintage",
    ]

    use_flags: Dict[str, bool] = {}
    capture_pcts: Dict[str, float] = {}
    price_sources: Dict[str, str] = {}

    defaults = {
        "PRE_IDA": (False, 0.0, "09:00 ID vintage"),
        "IDA3": (True, 100.0, "IDA3 vintage"),
        "IDA1": (True, 100.0, "IDA1 vintage"),
        "IDA2": (True, 100.0, "IDA2 vintage"),
    }

    for key in ["PRE_IDA", "IDA3", "IDA1", "IDA2"]:
        default_use, default_capture, default_price = defaults[key]
        with st.expander(key, expanded=(key != "PRE_IDA")):
            use_flags[key] = st.checkbox(f"Use {key}", value=default_use, key=f"use_{key}")
            capture_pcts[key] = st.slider(
                f"{key} Target Capture %",
                0.0,
                100.0,
                default_capture,
                1.0,
                key=f"capture_{key}",
            )
            price_sources[key] = st.selectbox(
                f"{key} price source",
                price_options,
                index=price_options.index(default_price),
                key=f"price_{key}",
            )

if v3_file is None:
    st.info("Upload v3_da_revenue CSV to start. v4/v5 and raw EPEX files can be added after that.")
    st.stop()

try:
    input_df, sources, v3_norm_df = build_input_data(v3_file.getvalue(), asset_label)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not build input data from v3 file: {exc}")
    st.stop()

v4_id_base, v4_info, v4_norm_df, v4_score_df = load_contribution_file(v4_file, input_df, asset, purpose="id")
v5_imbalance_base, v5_info, v5_norm_df, v5_score_df = load_contribution_file(v5_file, input_df, asset, purpose="imbalance")

vintage_prices, epex_info_df = load_raw_vintage_prices(
    epex_files,
    input_df,
    fallback=input_df["id3_benchmark_price_eur_per_mwh"],
)

settlement_df, kpi = run_simulation(
    input_df=input_df,
    vintage_prices=vintage_prices,
    strategy_mode=strategy_mode,
    da_sold_pct=da_sold_pct,
    use_flags=use_flags,
    capture_pcts=capture_pcts,
    price_sources=price_sources,
    contribution_source_mode=contribution_source_mode,
    v4_id_base=v4_id_base,
    v5_imbalance_base=v5_imbalance_base,
)

st.subheader(f"Trader Board Summary - {asset_label}")

kpi_cols = st.columns(4)
kpi_cols[0].metric("DA Revenue", fmt_eur(kpi["DA revenue EUR"]))
kpi_cols[1].metric("ID Revenue", fmt_eur(kpi["ID revenue EUR"]))
kpi_cols[2].metric("Imbalance Settlement", fmt_eur(kpi["Imbalance settlement EUR"]))
kpi_cols[3].metric("Total Revenue", fmt_eur(kpi["Total revenue EUR"]))

kpi_cols2 = st.columns(4)
kpi_cols2[0].metric("No-ID Benchmark", fmt_eur(kpi["No-ID benchmark EUR"]))
kpi_cols2[1].metric("ID Strategy Value", fmt_eur(kpi["ID strategy value EUR"]))
kpi_cols2[2].metric("Executed ID Trade", fmt_mwh(kpi["Total executed ID trade MWh"]))
kpi_cols2[3].metric("Remaining Imbalance", fmt_mwh(kpi["Remaining imbalance MWh"]))

with st.expander("Current data-source status", expanded=True):
    source_rows = [
        {"item": "v3 DA revenue", "status": "ok", "selected_column": sources.get("DA revenue column"), "sum": float(input_df["da_revenue_eur_5min"].sum())},
        {"item": "v4 ID contribution", "status": "ok" if v4_info.get("selected_column") not in ["not uploaded", "not detected"] else str(v4_info.get("selected_column")), "selected_column": v4_info.get("selected_column"), "sum": v4_info.get("sum", 0.0)},
        {"item": "v5 imbalance contribution", "status": "ok" if v5_info.get("selected_column") not in ["not uploaded", "not detected"] else str(v5_info.get("selected_column")), "selected_column": v5_info.get("selected_column"), "sum": v5_info.get("sum", 0.0)},
        {"item": "ID revenue source used", "status": str(settlement_df["id_revenue_source"].iloc[0]), "selected_column": "", "sum": float(settlement_df["id_revenue_eur"].sum())},
        {"item": "Imbalance source used", "status": str(settlement_df["imbalance_source"].iloc[0]), "selected_column": "", "sum": float(settlement_df["imbalance_settlement_eur"].sum())},
    ]
    st.dataframe(pd.DataFrame(source_rows), use_container_width=True)

st.divider()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Graphs",
    "Settlement table",
    "Input check",
    "Column diagnostics",
    "Download",
])

with tab1:
    st.plotly_chart(make_da_chart(settlement_df), use_container_width=True)
    st.plotly_chart(make_id_imbalance_chart(settlement_df), use_container_width=True)

with tab2:
    display_cols = [
        "time_label",
        "da_revenue_eur",
        "id_revenue_eur",
        "imbalance_settlement_eur",
        "id_plus_imbalance_eur",
        "total_revenue_eur",
        "forecast_error_mwh",
        "capture_factor",
        "total_id_trade_mwh",
        "remaining_imbalance_mwh",
    ]
    st.dataframe(settlement_df[display_cols], use_container_width=True, height=500)

with tab3:
    st.write("Detected v3 source columns")
    st.dataframe(pd.DataFrame([sources]).T.rename(columns={0: "detected_column"}), use_container_width=True)
    st.write("Input data preview")
    st.dataframe(input_df.head(20), use_container_width=True)

with tab4:
    st.write("v3 normalized columns")
    st.dataframe(pd.DataFrame({"v3_columns": list(v3_norm_df.columns)}), use_container_width=True, height=220)

    st.write("v4 ID contribution detection")
    st.json(v4_info)
    if not v4_score_df.empty:
        st.dataframe(v4_score_df.head(30), use_container_width=True, height=300)
    elif v4_file is None:
        st.info("v4 file was not uploaded.")
    else:
        st.warning("No numeric v4 columns were detected.")

    st.write("v5 imbalance contribution detection")
    st.json(v5_info)
    if not v5_score_df.empty:
        st.dataframe(v5_score_df.head(30), use_container_width=True, height=300)
    elif v5_file is None:
        st.info("v5 file was not uploaded.")
    else:
        st.warning("No numeric v5 columns were detected.")

    st.write("raw EPEX file mapping")
    if epex_info_df.empty:
        st.info("No raw EPEX files uploaded.")
    else:
        st.dataframe(epex_info_df, use_container_width=True)

with tab5:
    diagnostics = pd.DataFrame(source_rows)
    excel_bytes = make_excel_download(input_df, settlement_df, kpi, diagnostics)
    st.download_button(
        label="Download calculated result as Excel",
        data=excel_bytes,
        file_name="v17_streamlit_mvp_v2_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        label="Download settlement table as CSV",
        data=settlement_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="v17_streamlit_mvp_v2_settlement.csv",
        mime="text/csv",
    )
