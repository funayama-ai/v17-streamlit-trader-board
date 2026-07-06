from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, Iterable, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =============================================================================
# Convex Asset Trader Experience Board v17 - Streamlit MVP
# =============================================================================
# Purpose:
#   This is the first web-app version of the PyCharm / Excel v17 trader board.
#   It does not try to reproduce every Excel cell/formula yet. Instead, it gives
#   a clean MVP:
#       1) upload v3/v4/v5 CSVs and raw EPEX vintage CSVs
#       2) choose Portfolio / Wind / Solar
#       3) change DA / IDA3 / IDA1 / IDA2 simulation parameters
#       4) display KPI, DA graph, and ID / residual imbalance graph
#       5) download the calculated settlement table as Excel
#
# Required input:
#   - v3_da_revenue_YYYY_MM_DD.csv
# Optional but recommended:
#   - v4_id_correction_pnl_YYYY_MM_DD.csv
#   - v5_no_correction_imbalance_pnl_YYYY_MM_DD.csv
#   - raw EPEX ID vintage CSV files
#
# Notes:
#   - v16 files are not used.
#   - This app is intentionally robust to small CSV column-name differences.
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
    da_revenue_candidates: Tuple[str, ...]
    forecast_candidates: Tuple[str, ...]
    actual_candidates: Tuple[str, ...]
    error_mwh_candidates: Tuple[str, ...]


ASSETS: Dict[str, AssetConfig] = {
    "Portfolio / DK1 Wind + Solar": AssetConfig(
        label="Portfolio / DK1 Wind + Solar",
        key="portfolio",
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
            "forecast_error_mwh_5min",
            "forecast_error_mwh",
            "error_mwh",
        ),
    ),
    "Wind-only / DK1 Wind Only": AssetConfig(
        label="Wind-only / DK1 Wind Only",
        key="wind",
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
            "forecast_error_mwh_5min",
        ),
    ),
    "Solar-only / DK1 Solar Only": AssetConfig(
        label="Solar-only / DK1 Solar Only",
        key="solar",
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
            "forecast_error_mwh_5min",
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
    "timestamp",
    "datetime",
    "time",
    "delivery_start",
    "delivery_start_time",
    "delivery period start",
    "time_slot",
)

WINDOWS = {
    "PRE_IDA": ("00:00", "09:55"),
    "IDA3": ("09:55", "14:55"),
    "IDA1": ("14:55", "21:55"),
    "IDA2": ("21:55", "24:00"),
}


# =============================================================================
# Helpers
# =============================================================================

def normalize_col(name: object) -> str:
    text = str(name).strip().lower()
    text = text.replace("/", "_").replace("-", "_")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def read_csv_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        raise ValueError("No file uploaded")

    raw = uploaded_file.getvalue()
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=None, engine="python", encoding=encoding)
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise ValueError(f"Could not read CSV: {last_error}")


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


def first_column_containing(df: pd.DataFrame, required_words: Iterable[str]) -> Optional[str]:
    words = [normalize_col(w) for w in required_words]
    for col in df.columns:
        if all(w in col for w in words):
            return col
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

    # If only HH:MM was given, pd.to_datetime may attach today's date or fail.
    if parsed.isna().mean() > 0.50:
        parsed = pd.to_datetime("2026-01-01 " + raw.astype(str), errors="coerce")

    if parsed.isna().all():
        return pd.Series(pd.date_range("2026-01-01", periods=n, freq="5min"), index=df.index)

    # Fill missing with a regular 5-min sequence to keep the chart stable.
    fallback = pd.Series(pd.date_range(parsed.dropna().iloc[0].normalize(), periods=n, freq="5min"), index=df.index)
    return pd.Series(parsed, index=df.index).fillna(fallback)


def hhmm_to_minutes(text: str) -> int:
    if text == "24:00":
        return 24 * 60
    hour, minute = text.split(":")
    return int(hour) * 60 + int(minute)


def time_minutes(ts: pd.Series) -> pd.Series:
    return pd.to_datetime(ts).dt.hour * 60 + pd.to_datetime(ts).dt.minute


def window_mask(ts: pd.Series, start_hhmm: str, end_hhmm: str) -> pd.Series:
    mins = time_minutes(ts)
    return (mins >= hhmm_to_minutes(start_hhmm)) & (mins < hhmm_to_minutes(end_hhmm))


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
        temp = pd.DataFrame({"timestamp": ts, "price": price}).dropna(subset=["price"])
        if not temp.empty:
            temp = temp.drop_duplicates("timestamp", keep="last").sort_values("timestamp")
            base = pd.DataFrame({"timestamp": input_df["timestamp"]})
            merged = base.merge(temp, on="timestamp", how="left")
            return merged["price"].ffill().bfill()

    # Fallback: if the file has 288 or similar rows, align by row order.
    if len(price) >= len(input_df):
        return price.iloc[: len(input_df)].reset_index(drop=True).ffill().bfill()
    if len(price) > 0:
        return price.reindex(range(len(input_df))).ffill().bfill().fillna(float(price.dropna().iloc[0]))
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


def load_raw_vintage_prices(files, input_df: pd.DataFrame, fallback: pd.Series) -> Dict[str, pd.Series]:
    result: Dict[str, pd.Series] = {
        "ID3 Benchmark Price": fallback.copy(),
        "03:00 ID vintage": fallback.copy(),
        "06:00 ID vintage": fallback.copy(),
        "09:00 ID vintage": fallback.copy(),
        "IDA3 vintage": fallback.copy(),
        "IDA1 vintage": fallback.copy(),
        "IDA2 vintage": fallback.copy(),
    }

    if not files:
        return result

    for uploaded in files:
        key = raw_price_key_from_filename(uploaded.name)
        if key is None:
            continue
        try:
            df = read_csv_upload(uploaded)
            aligned = align_price_to_input(df, input_df)
            if aligned is not None:
                result[key] = aligned.astype(float).reset_index(drop=True)
        except Exception as exc:  # noqa: BLE001
            st.warning(f"Could not read raw EPEX file {uploaded.name}: {exc}")

    return result


@st.cache_data(show_spinner=False)
def build_input_data(v3_bytes: bytes, asset_label: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    asset = ASSETS[asset_label]
    raw_df = pd.read_csv(io.BytesIO(v3_bytes), sep=None, engine="python")
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

    # Fill prices with safe defaults if missing, so the app remains usable.
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

    # Keep the first 288 rows if a full-day file has extra rows.
    out = out.dropna(subset=["timestamp"]).sort_values("timestamp").drop_duplicates("timestamp", keep="last")
    if len(out) > 288:
        out = out.iloc[:288].copy()
    out = out.reset_index(drop=True)

    # If error is still missing, set 0 only at the final fallback stage.
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
    return out, sources


def run_simulation(
    input_df: pd.DataFrame,
    vintage_prices: Dict[str, pd.Series],
    strategy_mode: str,
    da_sold_pct: float,
    use_flags: Dict[str, bool],
    capture_pcts: Dict[str, float],
    price_sources: Dict[str, str],
) -> Tuple[pd.DataFrame, Dict[str, float]]:
    df = input_df.copy()

    da_multiplier = 0.0 if strategy_mode == "ID Only" else da_sold_pct / 100.0
    allow_id = strategy_mode != "DA Only"

    df["da_revenue_eur"] = df["da_revenue_eur_5min"] * da_multiplier

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
    df["remaining_imbalance_mwh"] = df["forecast_error_mwh"] - df["total_id_trade_mwh"]
    df["id_revenue_eur"] = total_id_revenue
    df["imbalance_settlement_eur"] = df["remaining_imbalance_mwh"] * df["imbalance_price_eur_per_mwh"]
    df["total_revenue_eur"] = df["da_revenue_eur"] + df["id_revenue_eur"] + df["imbalance_settlement_eur"]
    df["no_id_benchmark_eur"] = df["da_revenue_eur"] + df["forecast_error_mwh"] * df["imbalance_price_eur_per_mwh"]
    df["id_strategy_value_eur"] = df["total_revenue_eur"] - df["no_id_benchmark_eur"]

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
        height=470,
        margin=dict(l=40, r=20, t=70, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(tickmode="array", tickvals=df["time_label"].iloc[::12], tickangle=-45),
        yaxis=dict(range=[y_min, y_max], title="EUR per 5-min interval"),
    )
    return fig


def make_excel_download(input_df: pd.DataFrame, settlement_df: pd.DataFrame, kpi: Dict[str, float]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([kpi]).T.rename(columns={0: "value"}).to_excel(writer, sheet_name="KPI")
        input_df.to_excel(writer, sheet_name="Input_Data", index=False)
        settlement_df.to_excel(writer, sheet_name="Settlement", index=False)
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

    st.header("3. Strategy")
    strategy_mode = st.selectbox("Strategy Mode", ["DA + ID Correction", "DA Only", "ID Only"])
    da_sold_pct = st.slider("DA Sold % of Forecast", 0.0, 100.0, 100.0, 1.0)

    st.header("4. ID capture")
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
    input_df, sources = build_input_data(v3_file.getvalue(), asset_label)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not build input data from v3 file: {exc}")
    st.stop()

vintage_prices = load_raw_vintage_prices(
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

st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "Graphs",
    "Settlement table",
    "Input check",
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
        "total_revenue_eur",
        "forecast_error_mwh",
        "total_id_trade_mwh",
        "remaining_imbalance_mwh",
    ]
    st.dataframe(settlement_df[display_cols], use_container_width=True, height=500)

with tab3:
    st.write("Detected source columns")
    st.dataframe(pd.DataFrame([sources]).T.rename(columns={0: "detected_column"}), use_container_width=True)
    st.write("Input data preview")
    st.dataframe(input_df.head(20), use_container_width=True)

    if v4_file is not None:
        st.success(f"v4 uploaded: {v4_file.name}")
    if v5_file is not None:
        st.success(f"v5 uploaded: {v5_file.name}")
    if epex_files:
        st.success(f"raw EPEX files uploaded: {len(epex_files)}")

with tab4:
    excel_bytes = make_excel_download(input_df, settlement_df, kpi)
    st.download_button(
        label="Download calculated result as Excel",
        data=excel_bytes,
        file_name="v17_streamlit_mvp_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        label="Download settlement table as CSV",
        data=settlement_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="v17_streamlit_mvp_settlement.csv",
        mime="text/csv",
    )
