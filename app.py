from __future__ import annotations

import io
import re
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


# =============================================================================
# Convex Asset Trader Experience Board v17 - Streamlit Web MVP v3.11
# =============================================================================
# Purpose:
#   Web-app MVP for the PyCharm / Excel v17 trader board.
#   This version is designed as the next step after Streamlit MVP v2.
#
# Important scope:
#   - v16 files are NOT used.
#   - Main inputs are v3 / v4 / v5 CSV files plus optional raw EPEX ID vintage CSVs.
#   - This is still a Web MVP, not a complete Excel formula-engine reproduction.
#
# Main updates in v3:
#   - Cleaner Trader Decision Board style layout.
#   - Strategy Mode: DA + ID Correction / DA Only.
#   - DA Sold % control.
#   - Separate IDA3 / IDA1 / IDA2 Use Auction and Target Capture % controls.
#   - Raw EPEX vintage price mapping and price-source status display.
#   - ID auction breakdown table and chart.
#   - Excel / CSV download of calculated web results.
#
# Main fixes in v3.1:
#   - Executed ID Trade MWh no longer becomes 0 just because raw EPEX prices are missing.
#   - Uploaded v4 contribution mode uses forecast-error capture as an approximate MWh basis.
#   - Auction price status shows not loaded when the selected raw EPEX price is unavailable.
#   - Imbalance wording is shown as Imbalance Settlement EUR in the UI.
#   - Cumulative chart spacing is adjusted to avoid title / legend overlap.
#
# Main fixes in v3.3:
#   - Plotly accidental box-zoom / drag-zoom is disabled for all charts.
#   - Plotly modebar removes zoom / pan / select / lasso controls.
#   - DA positive revenue uses a light orange color.
#   - DA negative revenue uses a light blue color.
#
# Main fixes in v3.7:
#   - UI wording remains Imbalance Settlement EUR.
#   - Corrects v5 no-correction total PnL handling:
#       Imbalance Settlement EUR = no-correction total PnL - DA revenue.
#   - DA Only mode shows positive / negative Imbalance Settlement EUR in the second graph.
#   - The third graph is restored to ID Revenue by Auction Window.
#
# Main fixes in v3.8:
#   - DA + ID Correction logic now keeps explicit no-ID imbalance settlement.
#   - Residual imbalance settlement is calculated from the residual imbalance ratio.
#   - DA-only benchmark and ID-correction case are compared explicitly.
#
# Main fixes in v3.9:
#   - Positive residual imbalance settlement EUR is green.
#   - Negative residual imbalance settlement EUR is light green.
#
# Main fixes in v3.10:
#   - Positive residual imbalance settlement EUR uses green.
#   - Negative residual imbalance settlement EUR uses pale green.
#   - DA-only imbalance settlement colors are preserved.
#
# Main fixes in v3.11:
#   - Adds built-in demo data mode for sharing the Streamlit URL with graphs already visible.
#   - Demo files are read from data/demo_2026_06_22/ in the GitHub repository.
#   - Upload custom CSV mode is still available for user-provided data.
#
# Required:
#   - v3_da_revenue_YYYY_MM_DD.csv
# Optional:
#   - v4_id_correction_pnl_YYYY_MM_DD.csv
#   - v5_no_correction_imbalance_pnl_YYYY_MM_DD.csv
#   - raw EPEX ID vintage CSV files
# =============================================================================

st.set_page_config(
    page_title="v17 Trader Board Web MVP v3.11",
    page_icon="⚡",
    layout="wide",
)

# =============================================================================
# Chart interaction and color settings
# =============================================================================
# Prevent accidental click-drag zoom behavior in Plotly charts.
PLOTLY_CONFIG = {
    "displayModeBar": True,
    "scrollZoom": False,
    "displaylogo": False,
    "modeBarButtonsToRemove": [
        "zoom2d",
        "pan2d",
        "select2d",
        "lasso2d",
        "zoomIn2d",
        "zoomOut2d",
        "autoScale2d",
    ],
}

# Requested DA revenue colors.
DA_POSITIVE_COLOR = "rgba(255, 183, 77, 0.75)"   # light orange
DA_NEGATIVE_COLOR = "rgba(144, 202, 249, 0.75)"  # light blue

# DA-only mode: show the no-correction imbalance settlement in light green.
DA_ONLY_IMBALANCE_POSITIVE_COLOR = "rgba(129, 199, 132, 0.72)"  # light green
DA_ONLY_IMBALANCE_NEGATIVE_COLOR = "rgba(165, 214, 167, 0.72)"  # pale green
DEFAULT_IMBALANCE_POSITIVE_COLOR = "rgba(67, 160, 71, 0.85)"    # green
DEFAULT_IMBALANCE_NEGATIVE_COLOR = "rgba(165, 214, 167, 0.72)"  # pale green
ID_POSITIVE_COLOR = "rgba(0, 102, 204, 0.95)"                  # blue
ID_NEGATIVE_COLOR = "rgba(144, 202, 249, 0.85)"                # light blue

# =============================================================================
# Built-in demo data settings
# =============================================================================

APP_DIR = Path(__file__).resolve().parent
DEMO_DATA_DIR = APP_DIR / "data" / "demo_2026_06_22"
DEMO_FILES = {
    "v3": DEMO_DATA_DIR / "v3_da_revenue_2026_06_22.csv",
    "v4": DEMO_DATA_DIR / "v4_id_correction_pnl_2026_06_22.csv",
    "v5": DEMO_DATA_DIR / "v5_no_correction_imbalance_pnl_2026_06_22.csv",
}


class LocalDemoFile:
    """Small wrapper so local demo CSVs behave like Streamlit UploadedFile objects."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.name = self.path.name

    def getvalue(self) -> bytes:
        return self.path.read_bytes()


def demo_files_available() -> bool:
    return all(path.exists() and path.is_file() for path in DEMO_FILES.values())


def missing_demo_files() -> List[str]:
    return [str(path.as_posix()) for path in DEMO_FILES.values() if not path.exists()]


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
            "portfolio_da_only_revenue_eur",
            "da_revenue_eur",
            "da_revenue_5min_eur",
            "da_revenue_eur_5min",
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
            "portfolio_forecast_error_5min_mwh",
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
            "id_correction_revenue_eur",
            "id_correction_value_eur",
            "id_revenue_eur",
        ),
        v5_imbalance_candidates=(
            "portfolio_no_correction_imbalance_pnl_eur",
            "portfolio_imbalance_pnl_eur",
            "portfolio_imbalance_settlement_eur",
            "portfolio_no_correction_pnl_eur",
            "no_correction_imbalance_pnl_eur",
            "no_correction_pnl_eur",
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
            "wind_da_only_revenue_eur",
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
            "wind_forecast_error_5min_mwh",
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
            "solar_da_only_revenue_eur",
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
            "solar_forecast_error_5min_mwh",
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
        "id3_benchmark_price_eur_mwh",
        "expost_id3_price_eur_per_mwh",
        "expost_id3_price_eur_mwh",
        "final_id3_price_eur_per_mwh",
        "final_id3_price_eur_mwh",
        "id3_price_eur_per_mwh",
        "id3_price_eur_mwh",
        "id3_benchmark_price",
        "id3_price",
        "id_price_eur_per_mwh",
        "id_price_eur_mwh",
        "intraday_price_eur_per_mwh",
        "intraday_price_eur_mwh",
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

AUCTIONS = {
    "IDA3": ("09:55", "14:55"),
    "IDA1": ("14:55", "21:55"),
    "IDA2": ("21:55", "24:00"),
}

PRICE_OPTIONS = [
    "ID3 Benchmark Price",
    "03:00 ID vintage",
    "06:00 ID vintage",
    "09:00 ID vintage",
    "IDA3 vintage",
    "IDA1 vintage",
    "IDA2 vintage",
]

DEFAULT_AUCTION_SETTINGS = {
    "IDA3": (True, 100.0, "ID3 Benchmark Price"),
    "IDA1": (True, 100.0, "ID3 Benchmark Price"),
    "IDA2": (True, 100.0, "ID3 Benchmark Price"),
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

    for key, (start, end) in AUCTIONS.items():
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


def safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    denom = pd.to_numeric(denominator, errors="coerce").replace(0, np.nan)
    out = pd.to_numeric(numerator, errors="coerce") / denom
    return out.replace([np.inf, -np.inf], np.nan).fillna(0.0)


def price_series_is_loaded(price: pd.Series) -> bool:
    """Return True when a price series looks usable for display / MWh conversion.

    In this MVP, missing raw EPEX files often become an all-zero fallback series.
    Showing that as 0 EUR/MWh is misleading, so the UI marks it as not loaded.
    """
    clean = pd.to_numeric(price, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if clean.empty:
        return False
    return float(clean.abs().sum()) > 1e-9


def fmt_price_or_not_loaded(value: object) -> str:
    if value is None or pd.isna(value):
        return "not loaded"
    try:
        return f"{float(value):,.2f}"
    except Exception:  # noqa: BLE001
        return "not loaded"


def aligned_price_from_source_df(
    source_df: pd.DataFrame,
    input_df: pd.DataFrame,
    candidates: Iterable[str],
) -> Tuple[Optional[pd.Series], str]:
    """Find and align a price column from a normalized source dataframe."""
    if source_df is None or source_df.empty:
        return None, "not available"

    price_col = first_existing_column(source_df, candidates)
    if price_col is None:
        return None, "not detected"

    aligned = align_series_to_input(source_df, input_df, price_col)
    if not price_series_is_loaded(aligned):
        return None, f"{price_col} detected but empty/zero"

    return aligned.astype(float).reset_index(drop=True), price_col


def derive_id3_price_from_v4_contribution(input_df: pd.DataFrame, v4_id_base: pd.Series) -> Optional[pd.Series]:
    """Derive an implied ID3 benchmark price when no explicit price column is available.

    This is a fallback for the web MVP only. It assumes the uploaded v4 ID contribution
    is approximately forecast_error_mwh × ID3 benchmark price. Rows with near-zero
    exposure are ignored, and gaps are interpolated.
    """
    if v4_id_base is None or v4_id_base.abs().sum() <= 0:
        return None

    error = pd.to_numeric(input_df.get("forecast_error_mwh", pd.Series(0.0, index=input_df.index)), errors="coerce")
    contrib = pd.to_numeric(v4_id_base, errors="coerce").reindex(input_df.index).fillna(0.0)
    denom = error.where(error.abs() > 1e-6)
    implied = (contrib / denom).replace([np.inf, -np.inf], np.nan)

    # Keep a broad but sane electricity-price range. Negative ID prices are allowed.
    implied = implied.where(implied.between(-1000.0, 2000.0))
    if implied.notna().sum() < 10:
        return None

    implied = implied.interpolate(limit_direction="both").ffill().bfill().fillna(0.0)
    if not price_series_is_loaded(implied):
        return None
    return implied.astype(float).reset_index(drop=True)


def build_id3_benchmark_price(
    input_df: pd.DataFrame,
    v4_norm_df: pd.DataFrame,
    v5_norm_df: pd.DataFrame,
    v4_id_base: pd.Series,
) -> Tuple[pd.Series, Dict[str, object]]:
    """Return the best available ID3 benchmark price series and source info."""
    zero = pd.Series(0.0, index=input_df.index, dtype=float).reset_index(drop=True)

    direct = pd.to_numeric(input_df.get("id3_benchmark_price_eur_per_mwh", zero), errors="coerce").fillna(0.0)
    direct = direct.reset_index(drop=True)
    if price_series_is_loaded(direct):
        return direct, {
            "status": "loaded from v3",
            "selected_column": "id3_benchmark_price_eur_per_mwh",
            "avg_price": float(direct.mean()),
            "min_price": float(direct.min()),
            "max_price": float(direct.max()),
        }

    for label, source_df in (("v4", v4_norm_df), ("v5", v5_norm_df)):
        aligned, col = aligned_price_from_source_df(source_df, input_df, PRICE_CANDIDATES["id3"])
        if aligned is not None:
            return aligned, {
                "status": f"loaded from {label}",
                "selected_column": col,
                "avg_price": float(aligned.mean()),
                "min_price": float(aligned.min()),
                "max_price": float(aligned.max()),
            }

    implied = derive_id3_price_from_v4_contribution(input_df, v4_id_base)
    if implied is not None:
        return implied, {
            "status": "implied from v4 ID contribution / forecast error",
            "selected_column": "v4_id_contribution / forecast_error_mwh",
            "avg_price": float(implied.mean()),
            "min_price": float(implied.min()),
            "max_price": float(implied.max()),
        }

    return zero, {
        "status": "not loaded",
        "selected_column": "not detected",
        "avg_price": np.nan,
        "min_price": np.nan,
        "max_price": np.nan,
    }


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
        tokens = set(col.split("_"))
        name = col

        if col == exact:
            score += 100
            reasons.append("exact candidate")

        if asset.key in tokens or asset.key in name:
            score += 20
            reasons.append(f"asset={asset.key}")

        if purpose == "id":
            if "id" in tokens or "intraday" in name:
                score += 15
                reasons.append("id/intraday")
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

    # Fallback: row order. This is usually correct for 288-row daily files.
    aligned = value.reset_index(drop=True).reindex(range(len(input_df))).fillna(0.0)
    return aligned.astype(float)


def is_total_no_correction_pnl_column(column_name: str) -> bool:
    """Detect v5 columns that are total no-correction PnL, not imbalance-only settlement.

    Example:
      portfolio_no_correction_pnl_eur

    This column is DA revenue + imbalance settlement. For DA-only imbalance
    visualization we must subtract the DA revenue component.
    """
    name = normalize_col(column_name)
    has_no_correction = "no_correction" in name or ("no" in name and "correction" in name)
    looks_like_total_pnl = any(token in name for token in ["pnl", "revenue", "value", "eur"])
    is_imbalance_specific = "imbalance" in name or "settlement" in name
    return bool(has_no_correction and looks_like_total_pnl and not is_imbalance_specific)


def convert_v5_to_imbalance_settlement_if_needed(
    aligned: pd.Series,
    input_df: pd.DataFrame,
    selected_col: str,
    purpose: str,
) -> Tuple[pd.Series, Dict[str, object]]:
    """Return imbalance-settlement-only series for v5.

    Some v5 files expose `portfolio_no_correction_pnl_eur`, which is not the
    imbalance settlement alone. It is the total DA-only / no-correction PnL:

        no_correction_total_pnl = DA revenue + imbalance settlement

    Therefore:

        imbalance settlement = no_correction_total_pnl - DA revenue

    This is why the 2026-06-22 Portfolio value should be around:

        4,371,252 EUR - 4,328,456 EUR = 42,796 EUR
    """
    aligned = pd.to_numeric(aligned, errors="coerce").fillna(0.0).astype(float).reset_index(drop=True)

    if purpose != "imbalance" or not is_total_no_correction_pnl_column(selected_col):
        return aligned, {
            "source_type": "imbalance_settlement_column",
            "conversion": "none",
            "original_sum": float(aligned.sum()),
            "da_revenue_subtracted_sum": 0.0,
        }

    da_revenue = pd.to_numeric(input_df.get("da_revenue_eur_5min", 0.0), errors="coerce").fillna(0.0)
    da_revenue = da_revenue.reset_index(drop=True).reindex(range(len(aligned))).fillna(0.0).astype(float)
    converted = aligned - da_revenue

    return converted.astype(float), {
        "source_type": "converted_from_no_correction_total_pnl",
        "conversion": "imbalance_settlement_eur = selected_v5_total_pnl - v3_da_revenue_eur_5min",
        "original_sum": float(aligned.sum()),
        "da_revenue_subtracted_sum": float(da_revenue.sum()),
    }


def load_contribution_file(
    uploaded_file,
    input_df: pd.DataFrame,
    asset: AssetConfig,
    purpose: str,
) -> Tuple[pd.Series, Dict[str, object], pd.DataFrame, pd.DataFrame]:
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

    aligned_raw = align_series_to_input(df, input_df, selected_col)
    aligned, conversion_info = convert_v5_to_imbalance_settlement_if_needed(
        aligned=aligned_raw,
        input_df=input_df,
        selected_col=selected_col,
        purpose=purpose,
    )
    info = {
        "file": uploaded_file.name,
        "selected_column": selected_col,
        "sum": float(aligned.sum()),
        "abs_sum": float(aligned.abs().sum()),
        "rows": int(len(aligned)),
        **conversion_info,
    }
    return aligned, info, df, score_df


# =============================================================================
# Raw EPEX vintage helpers
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
            "intraday_price",
            "intraday_price_eur_per_mwh",
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
        default_price = float(price.dropna().iloc[0]) if price.dropna().size else 0.0
        return price.reindex(range(len(input_df))).ffill().bfill().fillna(default_price)
    return None


def raw_price_key_from_filename(name: str) -> Optional[str]:
    n = name.lower()
    normalized = re.sub(r"[^a-z0-9]+", "_", n)

    if "ida3" in normalized or "09_55" in normalized or "0955" in normalized or "10_05" in normalized or "1005" in normalized:
        return "IDA3 vintage"
    if "ida1" in normalized or "14_55" in normalized or "1455" in normalized or "15_05" in normalized or "1505" in normalized:
        return "IDA1 vintage"
    if "ida2" in normalized or "21_55" in normalized or "2155" in normalized or "22_05" in normalized or "2205" in normalized:
        return "IDA2 vintage"
    if "03_00" in normalized or "0300" in normalized:
        return "03:00 ID vintage"
    if "06_00" in normalized or "0600" in normalized:
        return "06:00 ID vintage"
    if "09_00" in normalized or "0900" in normalized:
        return "09:00 ID vintage"
    return None


def load_raw_vintage_prices(
    files,
    input_df: pd.DataFrame,
    id3_benchmark_price: pd.Series,
    id3_benchmark_info: Dict[str, object],
) -> Tuple[Dict[str, pd.Series], pd.DataFrame]:
    zero = pd.Series(0.0, index=input_df.index, dtype=float).reset_index(drop=True)
    id3 = pd.to_numeric(id3_benchmark_price, errors="coerce").reset_index(drop=True).reindex(range(len(input_df))).ffill().bfill().fillna(0.0)

    # Important v3.4 behavior:
    #   - ID3 Benchmark Price uses the best available ID3 benchmark series.
    #   - Raw vintage price sources remain "not loaded" unless the user uploads matching raw EPEX CSVs.
    # This avoids accidentally treating IDA1/IDA2/IDA3 vintage prices as if they were ID3 benchmark prices.
    result: Dict[str, pd.Series] = {
        "ID3 Benchmark Price": id3.astype(float),
        "03:00 ID vintage": zero.copy(),
        "06:00 ID vintage": zero.copy(),
        "09:00 ID vintage": zero.copy(),
        "IDA3 vintage": zero.copy(),
        "IDA1 vintage": zero.copy(),
        "IDA2 vintage": zero.copy(),
    }

    rows: List[Dict[str, object]] = [
        {
            "file": "v3/v4/v5 detected source",
            "mapped_to": "ID3 Benchmark Price",
            "status": id3_benchmark_info.get("status", "not loaded"),
            "selected_column": id3_benchmark_info.get("selected_column", ""),
            "avg_price": id3_benchmark_info.get("avg_price", np.nan),
            "min_price": id3_benchmark_info.get("min_price", np.nan),
            "max_price": id3_benchmark_info.get("max_price", np.nan),
        }
    ]

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
            if aligned is not None and price_series_is_loaded(aligned):
                result[key] = aligned.astype(float).reset_index(drop=True)
                rows.append(
                    {
                        "file": uploaded.name,
                        "mapped_to": key,
                        "status": "ok",
                        "selected_column": "auto-detected price column",
                        "avg_price": float(result[key].mean()),
                        "min_price": float(result[key].min()),
                        "max_price": float(result[key].max()),
                    }
                )
            else:
                rows.append({"file": uploaded.name, "mapped_to": key, "status": "no usable price column detected"})
        except Exception as exc:  # noqa: BLE001
            rows.append({"file": uploaded.name, "mapped_to": key, "status": f"error: {exc}"})

    return result, pd.DataFrame(rows)


# =============================================================================
# Build input and run simulation
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
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, float]]:
    df = input_df.copy()
    df["strategy_mode"] = strategy_mode

    da_factor = da_sold_pct / 100.0
    allow_id = strategy_mode == "DA + ID Correction"
    capture_factor = build_capture_factor(df, use_flags, capture_pcts, allow_id=allow_id)

    # MVP assumption:
    # DA Sold % scales DA revenue and the original v17 forecast-error exposure.
    # Therefore DA Sold % = 0% suppresses both DA revenue and ID correction value.
    df["da_position_factor"] = da_factor
    df["da_revenue_eur"] = df["da_revenue_eur_5min"] * da_factor
    df["strategy_error_mwh"] = df["forecast_error_mwh"] * da_factor
    df["capture_factor"] = capture_factor

    use_uploaded_v4 = (
        contribution_source_mode in ("Auto: prefer uploaded v4/v5", "Use uploaded v4/v5 benchmark columns")
        and v4_id_base.abs().sum() > 0
    )
    use_uploaded_v5 = (
        contribution_source_mode in ("Auto: prefer uploaded v4/v5", "Use uploaded v4/v5 benchmark columns")
        and v5_imbalance_base.abs().sum() > 0
    )

    total_trade_mwh = pd.Series(0.0, index=df.index, dtype=float)
    total_id_revenue = pd.Series(0.0, index=df.index, dtype=float)
    auction_rows: List[Dict[str, object]] = []

    for key, (start, end) in AUCTIONS.items():
        mask = window_mask(df["timestamp"], start, end)
        is_used = bool(use_flags.get(key, False) and allow_id)
        capture = (capture_pcts.get(key, 0.0) / 100.0) if is_used else 0.0
        source_name = price_sources.get(key, "ID3 Benchmark Price")
        price = vintage_prices.get(source_name, df["id3_benchmark_price_eur_per_mwh"])
        price = pd.Series(price).reset_index(drop=True).reindex(df.index).ffill().bfill().fillna(0.0)

        trade_col = f"{key.lower()}_trade_mwh"
        price_col = f"{key.lower()}_price_eur_per_mwh"
        revenue_col = f"{key.lower()}_id_revenue_eur"

        df[price_col] = price
        price_available = price_series_is_loaded(price)
        avg_price = float(price.mean()) if price_available else np.nan

        if use_uploaded_v4:
            # Use v4 as the benchmark ID-correction EUR contribution, then scale by
            # DA position and auction capture. For MWh display, do NOT divide by
            # price when raw EPEX prices are missing. Instead, show an approximate
            # executed MWh from forecast-error exposure x auction capture.
            df[revenue_col] = np.where(mask, v4_id_base.reindex(df.index).fillna(0.0) * da_factor * capture, 0.0)
            df[trade_col] = np.where(mask, df["strategy_error_mwh"] * capture, 0.0)
            mwh_basis = "approx: forecast_error_mwh × capture"
        else:
            # Model-only fallback: correction volume equals exposed forecast error
            # multiplied by auction capture, valued at selected vintage price.
            df[trade_col] = np.where(mask, df["strategy_error_mwh"] * capture, 0.0)
            df[revenue_col] = df[trade_col] * price
            mwh_basis = "forecast_error_mwh × capture"

        total_trade_mwh = total_trade_mwh + df[trade_col]
        total_id_revenue = total_id_revenue + df[revenue_col]

        auction_rows.append(
            {
                "auction": key,
                "window": f"{start}-{end}",
                "use_auction": is_used,
                "target_capture_pct": capture_pcts.get(key, 0.0),
                "price_source": source_name,
                "price_status": "loaded" if price_available else "not loaded",
                "avg_price_eur_per_mwh": avg_price,
                "avg_price_display": fmt_price_or_not_loaded(avg_price),
                "net_executed_trade_mwh": float(df[trade_col].sum()),
                "abs_executed_trade_mwh": float(df[trade_col].abs().sum()),
                "id_revenue_eur": float(df[revenue_col].sum()),
                "mwh_basis": mwh_basis,
            }
        )

    df["total_id_trade_mwh"] = total_trade_mwh
    df["total_id_trade_abs_mwh"] = df["total_id_trade_mwh"].abs()
    df["id_revenue_eur"] = total_id_revenue
    df["id_revenue_source"] = "uploaded_v4_scaled_by_da_position_and_auction_capture_mwh_approx" if use_uploaded_v4 else "model_from_strategy_error_and_selected_vintage_prices"

    df["remaining_imbalance_mwh"] = df["strategy_error_mwh"] - df["total_id_trade_mwh"]
    df["remaining_imbalance_abs_mwh"] = df["remaining_imbalance_mwh"].abs()

    # Explicit no-ID imbalance settlement benchmark.
    # This is important for explaining the Excel v17 logic:
    #   DA-only / no-correction revenue = DA revenue + no-ID imbalance settlement.
    if use_uploaded_v5:
        df["no_id_imbalance_settlement_eur"] = (
            v5_imbalance_base.reindex(df.index).fillna(0.0).astype(float) * da_factor
        )
    else:
        df["no_id_imbalance_settlement_eur"] = df["strategy_error_mwh"] * df["imbalance_price_eur_per_mwh"]

    # Residual imbalance ratio after ID correction.
    # If no ID correction is active, this ratio is 1.0.
    # If an interval is fully captured, this ratio approaches 0.0.
    safe_strategy_error = df["strategy_error_mwh"].where(df["strategy_error_mwh"].abs() > 1e-9)
    residual_ratio = (df["remaining_imbalance_mwh"] / safe_strategy_error).replace([np.inf, -np.inf], np.nan)
    residual_ratio = residual_ratio.fillna(1.0 if not allow_id else 0.0).clip(lower=-2.0, upper=2.0)
    df["residual_imbalance_ratio"] = residual_ratio

    if use_uploaded_v5:
        # v5 is the no-correction imbalance benchmark after conversion.
        # Residual imbalance settlement is scaled by the residual volume ratio,
        # not by total no-correction PnL. This keeps DA-only and DA+ID cases separate.
        df["imbalance_settlement_eur"] = df["no_id_imbalance_settlement_eur"] * df["residual_imbalance_ratio"]
        df["imbalance_source"] = "uploaded_v5_scaled_by_residual_imbalance_ratio"
    else:
        df["imbalance_settlement_eur"] = df["remaining_imbalance_mwh"] * df["imbalance_price_eur_per_mwh"]
        df["imbalance_source"] = "model_remaining_mwh_times_imbalance_price"

    df["total_revenue_eur"] = df["da_revenue_eur"] + df["id_revenue_eur"] + df["imbalance_settlement_eur"]
    df["no_id_benchmark_eur"] = df["da_revenue_eur"] + df["no_id_imbalance_settlement_eur"]

    df["id_strategy_value_eur"] = df["total_revenue_eur"] - df["no_id_benchmark_eur"]
    df["id_plus_imbalance_eur"] = df["id_revenue_eur"] + df["imbalance_settlement_eur"]

    auction_breakdown_df = pd.DataFrame(auction_rows)

    kpi = {
        "DA revenue EUR": float(df["da_revenue_eur"].sum()),
        "ID revenue EUR": float(df["id_revenue_eur"].sum()),
        "Imbalance Settlement EUR": float(df["imbalance_settlement_eur"].sum()),
        "Total revenue EUR": float(df["total_revenue_eur"].sum()),
        "No-ID benchmark EUR": float(df["no_id_benchmark_eur"].sum()),
        "No-ID imbalance settlement EUR": float(df["no_id_imbalance_settlement_eur"].sum()),
        "ID strategy value EUR": float(df["id_strategy_value_eur"].sum()),
        "Original forecast error MWh": float(df["forecast_error_mwh"].sum()),
        "Original absolute forecast error MWh": float(df["forecast_error_mwh"].abs().sum()),
        "Strategy exposure MWh": float(df["strategy_error_mwh"].sum()),
        "Strategy absolute exposure MWh": float(df["strategy_error_mwh"].abs().sum()),
        "Total executed ID trade MWh": float(df["total_id_trade_abs_mwh"].sum()),
        "Net executed ID trade MWh": float(df["total_id_trade_mwh"].sum()),
        "Remaining imbalance MWh": float(df["remaining_imbalance_abs_mwh"].sum()),
        "Net remaining imbalance MWh": float(df["remaining_imbalance_mwh"].sum()),
    }
    return df, auction_breakdown_df, kpi


# =============================================================================
# Charts and downloads
# =============================================================================

def make_da_chart(df: pd.DataFrame) -> go.Figure:
    pos = df["da_revenue_eur"].clip(lower=0)
    neg = df["da_revenue_eur"].clip(upper=0)
    y_min, y_max = nice_axis_limits(df["da_revenue_eur"])

    fig = go.Figure()
    fig.add_bar(x=df["time_label"], y=pos, name="Positive DA Revenue EUR", marker_color=DA_POSITIVE_COLOR)
    fig.add_bar(x=df["time_label"], y=neg, name="Negative DA Revenue EUR", marker_color=DA_NEGATIVE_COLOR)
    fig.update_layout(
        title="DA Revenue EUR - 5-min contribution",
        dragmode=False,
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

    strategy_mode = str(df["strategy_mode"].iloc[0]) if "strategy_mode" in df.columns and not df.empty else ""
    is_da_only = strategy_mode == "DA Only"

    if is_da_only:
        positive_imbalance_name = "Positive Imbalance Settlement EUR"
        negative_imbalance_name = "Negative Imbalance Settlement EUR"
        positive_imbalance_color = DA_ONLY_IMBALANCE_POSITIVE_COLOR
        negative_imbalance_color = DA_ONLY_IMBALANCE_NEGATIVE_COLOR
        daily_net = float(pd.to_numeric(df["imbalance_settlement_eur"], errors="coerce").fillna(0.0).sum())
        chart_title = f"DA-only Imbalance Settlement EUR - 5-min | Daily net: {daily_net:,.0f} EUR"
    else:
        positive_imbalance_name = "Positive residual imbalance settlement EUR"
        negative_imbalance_name = "Negative residual imbalance settlement EUR"
        positive_imbalance_color = DEFAULT_IMBALANCE_POSITIVE_COLOR
        negative_imbalance_color = DEFAULT_IMBALANCE_NEGATIVE_COLOR
        chart_title = "ID Revenue / Residual Imbalance Settlement EUR - 5-min"

    fig = go.Figure()
    fig.add_bar(x=df["time_label"], y=id_pos, name="Positive ID contribution", marker_color=ID_POSITIVE_COLOR)
    fig.add_bar(x=df["time_label"], y=id_neg, name="Negative ID contribution", marker_color=ID_NEGATIVE_COLOR)
    fig.add_bar(x=df["time_label"], y=imb_pos, name=positive_imbalance_name, marker_color=positive_imbalance_color)
    fig.add_bar(x=df["time_label"], y=imb_neg, name=negative_imbalance_name, marker_color=negative_imbalance_color)
    fig.update_layout(
        title=chart_title,
        dragmode=False,
        barmode="relative",
        height=500,
        margin=dict(l=40, r=20, t=90, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(tickmode="array", tickvals=df["time_label"].iloc[::12], tickangle=-45),
        yaxis=dict(range=[y_min, y_max], title="EUR per 5-min interval"),
    )
    return fig


def make_da_only_settlement_detail_chart(df: pd.DataFrame) -> go.Figure:
    """Detailed DA-only imbalance settlement chart.

    The main DA-only graph shows the 5-minute settlement bars. This detail chart
    keeps those bars and adds the cumulative imbalance settlement line so that
    the user can see both interval-level impact and daily accumulation.
    """
    settlement = pd.to_numeric(df["imbalance_settlement_eur"], errors="coerce").fillna(0.0)
    pos = settlement.clip(lower=0)
    neg = settlement.clip(upper=0)
    cumulative = settlement.cumsum()

    if "remaining_imbalance_mwh" in df.columns:
        net_remaining = pd.to_numeric(df["remaining_imbalance_mwh"], errors="coerce").fillna(0.0)
    else:
        net_remaining = pd.Series(0.0, index=df.index)

    if "remaining_imbalance_abs_mwh" in df.columns:
        abs_remaining = pd.to_numeric(df["remaining_imbalance_abs_mwh"], errors="coerce").fillna(net_remaining.abs())
    else:
        abs_remaining = net_remaining.abs()

    customdata = np.column_stack(
        [
            net_remaining.to_numpy(dtype=float),
            abs_remaining.to_numpy(dtype=float),
            cumulative.to_numpy(dtype=float),
        ]
    )

    y_min, y_max = nice_axis_limits(settlement)

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    hover_template = (
        "Time: %{x}<br>"
        "Settlement: %{y:,.0f} EUR<br>"
        "Net remaining imbalance: %{customdata[0]:,.2f} MWh<br>"
        "Abs remaining imbalance: %{customdata[1]:,.2f} MWh<br>"
        "Cumulative settlement: %{customdata[2]:,.0f} EUR"
        "<extra></extra>"
    )

    fig.add_bar(
        x=df["time_label"],
        y=pos,
        name="Positive Imbalance Settlement EUR",
        marker_color=DA_ONLY_IMBALANCE_POSITIVE_COLOR,
        customdata=customdata,
        hovertemplate=hover_template,
        secondary_y=False,
    )
    fig.add_bar(
        x=df["time_label"],
        y=neg,
        name="Negative Imbalance Settlement EUR",
        marker_color=DA_ONLY_IMBALANCE_NEGATIVE_COLOR,
        customdata=customdata,
        hovertemplate=hover_template,
        secondary_y=False,
    )
    fig.add_scatter(
        x=df["time_label"],
        y=cumulative,
        mode="lines",
        name="Cumulative Imbalance Settlement EUR",
        line=dict(color="rgba(46, 125, 50, 0.95)", width=2.5),
        secondary_y=True,
    )

    fig.update_layout(
        title="DA-only Imbalance Settlement Detail - 5-min and cumulative",
        dragmode=False,
        barmode="relative",
        height=420,
        margin=dict(l=55, r=60, t=85, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(tickmode="array", tickvals=df["time_label"].iloc[::12], tickangle=-45),
    )
    fig.update_yaxes(title_text="Settlement EUR per 5-min interval", range=[y_min, y_max], secondary_y=False)
    fig.update_yaxes(title_text="Cumulative settlement EUR", secondary_y=True)
    return fig


def make_auction_chart(auction_breakdown_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    if auction_breakdown_df.empty:
        fig.update_layout(height=420, dragmode=False, margin=dict(l=45, r=20, t=30, b=70))
        return fig

    fig.add_bar(
        x=auction_breakdown_df["auction"],
        y=auction_breakdown_df["id_revenue_eur"],
        name="ID revenue EUR",
    )
    fig.update_layout(
        height=420,
        dragmode=False,
        margin=dict(l=55, r=20, t=35, b=70),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0.0),
        yaxis=dict(title="EUR"),
    )
    return fig


def make_cumulative_chart(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_scatter(x=df["time_label"], y=df["da_revenue_eur"].cumsum(), mode="lines", name="DA")
    fig.add_scatter(x=df["time_label"], y=df["id_revenue_eur"].cumsum(), mode="lines", name="ID")
    fig.add_scatter(x=df["time_label"], y=df["imbalance_settlement_eur"].cumsum(), mode="lines", name="Imbalance Settlement EUR")
    fig.add_scatter(x=df["time_label"], y=df["total_revenue_eur"].cumsum(), mode="lines", name="Total")
    fig.update_layout(
        height=420,
        dragmode=False,
        margin=dict(l=55, r=25, t=45, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.04, xanchor="left", x=0.0),
        xaxis=dict(tickmode="array", tickvals=df["time_label"].iloc[::12], tickangle=-45),
        yaxis=dict(title="EUR"),
    )
    return fig




def make_strategy_waterfall_chart(df: pd.DataFrame, kpi: Dict[str, float]) -> go.Figure:
    """Excel v17-style bridge from DA-only benchmark to ID-correction result."""
    da_revenue = float(kpi.get("DA revenue EUR", 0.0))
    no_id_imbalance = float(kpi.get("No-ID imbalance settlement EUR", 0.0))
    no_id_total = float(kpi.get("No-ID benchmark EUR", 0.0))
    id_revenue = float(kpi.get("ID revenue EUR", 0.0))
    residual_imbalance = float(kpi.get("Imbalance Settlement EUR", 0.0))
    id_case_total = float(kpi.get("Total revenue EUR", 0.0))
    strategy_value = float(kpi.get("ID strategy value EUR", 0.0))

    fig = go.Figure(
        go.Waterfall(
            name="Strategy bridge",
            orientation="v",
            measure=["relative", "relative", "total", "relative", "relative", "total"],
            x=[
                "DA Revenue",
                "No-ID Imbalance",
                "DA-only Benchmark",
                "ID Revenue",
                "Residual Imbalance",
                "ID-Correction Total",
            ],
            y=[da_revenue, no_id_imbalance, no_id_total, id_revenue, residual_imbalance, id_case_total],
            connector={"line": {"width": 1}},
            text=[
                fmt_eur(da_revenue),
                fmt_eur(no_id_imbalance),
                fmt_eur(no_id_total),
                fmt_eur(id_revenue),
                fmt_eur(residual_imbalance),
                fmt_eur(id_case_total),
            ],
            textposition="outside",
        )
    )
    fig.update_layout(
        title=f"Strategy Result Waterfall | ID strategy value: {strategy_value:,.0f} EUR",
        dragmode=False,
        height=430,
        margin=dict(l=55, r=25, t=85, b=85),
        yaxis=dict(title="EUR"),
    )
    return fig


def make_case_comparison_table(kpi: Dict[str, float]) -> pd.DataFrame:
    da_revenue = float(kpi.get("DA revenue EUR", 0.0))
    no_id_imbalance = float(kpi.get("No-ID imbalance settlement EUR", 0.0))
    no_id_total = float(kpi.get("No-ID benchmark EUR", 0.0))
    id_revenue = float(kpi.get("ID revenue EUR", 0.0))
    residual_imbalance = float(kpi.get("Imbalance Settlement EUR", 0.0))
    id_case_total = float(kpi.get("Total revenue EUR", 0.0))
    return pd.DataFrame(
        [
            {"case": "DA Only / No ID", "DA revenue EUR": da_revenue, "ID revenue EUR": 0.0, "Imbalance Settlement EUR": no_id_imbalance, "Total revenue EUR": no_id_total},
            {"case": "DA + ID Correction", "DA revenue EUR": da_revenue, "ID revenue EUR": id_revenue, "Imbalance Settlement EUR": residual_imbalance, "Total revenue EUR": id_case_total},
            {"case": "ID strategy value", "DA revenue EUR": 0.0, "ID revenue EUR": id_revenue, "Imbalance Settlement EUR": residual_imbalance - no_id_imbalance, "Total revenue EUR": id_case_total - no_id_total},
        ]
    )


def make_excel_download(
    input_df: pd.DataFrame,
    settlement_df: pd.DataFrame,
    kpi: Dict[str, float],
    source_status_df: pd.DataFrame,
    auction_breakdown_df: pd.DataFrame,
    v3_columns_df: pd.DataFrame,
    v4_score_df: pd.DataFrame,
    v5_score_df: pd.DataFrame,
    epex_info_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([kpi]).T.rename(columns={0: "value"}).to_excel(writer, sheet_name="KPI")
        auction_breakdown_df.to_excel(writer, sheet_name="Auction_Breakdown", index=False)
        source_status_df.to_excel(writer, sheet_name="Data_Source_Status", index=False)
        input_df.to_excel(writer, sheet_name="Input_Data", index=False)
        settlement_df.to_excel(writer, sheet_name="Settlement", index=False)
        v3_columns_df.to_excel(writer, sheet_name="v3_Columns", index=False)
        v4_score_df.to_excel(writer, sheet_name="v4_Detection", index=False)
        v5_score_df.to_excel(writer, sheet_name="v5_Detection", index=False)
        epex_info_df.to_excel(writer, sheet_name="EPEX_Mapping", index=False)
    return output.getvalue()


def fmt_eur(value: float) -> str:
    return f"{value:,.0f} EUR"


def fmt_mwh(value: float) -> str:
    return f"{value:,.1f} MWh"


def fmt_pct(value: float) -> str:
    return f"{value:.0f}%"


def auction_display_table(auction_breakdown_df: pd.DataFrame) -> pd.DataFrame:
    if auction_breakdown_df.empty:
        return auction_breakdown_df
    preferred_cols = [
        "auction",
        "window",
        "use_auction",
        "target_capture_pct",
        "price_source",
        "price_status",
        "avg_price_display",
        "abs_executed_trade_mwh",
        "net_executed_trade_mwh",
        "id_revenue_eur",
        "mwh_basis",
    ]
    cols = [c for c in preferred_cols if c in auction_breakdown_df.columns]
    out = auction_breakdown_df[cols].copy()
    out = out.rename(
        columns={
            "avg_price_display": "avg_price_eur_per_mwh",
            "abs_executed_trade_mwh": "executed_trade_abs_mwh",
            "net_executed_trade_mwh": "executed_trade_net_mwh",
        }
    )
    return out


# =============================================================================
# UI
# =============================================================================

st.title("⚡ Convex Asset Trader Experience Board v17 - Web MVP v3.11")
st.caption(
    "Streamlit version of the v17 trader board concept. "
    "Uses v3/v4/v5 and optional raw EPEX vintage files. v16 files are not used."
)

demo_available = demo_files_available()

with st.sidebar:
    st.header("0. Data mode")
    data_mode_options = ["Use built-in demo data", "Upload custom CSV files"]
    data_mode = st.radio(
        "Data source",
        data_mode_options,
        index=0 if demo_available else 1,
        help=(
            "Use built-in demo data for a shareable app view. "
            "Switch to upload mode when testing another day or asset file set."
        ),
    )

    st.header("1. Data files")
    if data_mode == "Use built-in demo data":
        if demo_available:
            st.success("Using built-in demo data: 2026-06-22")
            v3_file = LocalDemoFile(DEMO_FILES["v3"])
            v4_file = LocalDemoFile(DEMO_FILES["v4"])
            v5_file = LocalDemoFile(DEMO_FILES["v5"])
            st.caption("Loaded from GitHub repository:")
            st.code(
                "data/demo_2026_06_22/\n"
                "  v3_da_revenue_2026_06_22.csv\n"
                "  v4_id_correction_pnl_2026_06_22.csv\n"
                "  v5_no_correction_imbalance_pnl_2026_06_22.csv",
                language="text",
            )
        else:
            v3_file = None
            v4_file = None
            v5_file = None
            st.error("Built-in demo files were not found in the repository.")
            st.caption("Missing files:")
            st.code("\n".join(missing_demo_files()), language="text")
        epex_files = st.file_uploader(
            "Optional: raw EPEX ID vintage CSVs",
            type=["csv"],
            accept_multiple_files=True,
        )
    else:
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
    strategy_mode = st.selectbox("Strategy Mode", ["DA + ID Correction", "DA Only"])
    da_sold_pct = st.slider("DA Sold %", 0.0, 100.0, 100.0, 1.0)

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

    st.header("5. ID auction controls")
    use_flags: Dict[str, bool] = {}
    capture_pcts: Dict[str, float] = {}
    price_sources: Dict[str, str] = {}

    for auction_key in ["IDA3", "IDA1", "IDA2"]:
        default_use, default_capture, default_price = DEFAULT_AUCTION_SETTINGS[auction_key]
        with st.expander(auction_key, expanded=True):
            use_flags[auction_key] = st.checkbox(f"Use {auction_key}", value=default_use, key=f"use_{auction_key}")
            capture_pcts[auction_key] = st.slider(
                f"{auction_key} Target Capture %",
                0.0,
                100.0,
                default_capture,
                1.0,
                key=f"capture_{auction_key}",
            )
            # v3.10 change:
            # Streamlit's selectbox dropdown can be clipped at the bottom of the sidebar,
            # especially for IDA2. A radio control avoids the hidden-scroll problem and
            # keeps all price-source options directly visible in the sidebar.
            price_sources[auction_key] = st.radio(
                f"{auction_key} price source",
                PRICE_OPTIONS,
                index=PRICE_OPTIONS.index(default_price),
                key=f"price_{auction_key}",
                horizontal=False,
            )

if v3_file is None:
    if data_mode == "Use built-in demo data":
        st.info(
            "Built-in demo mode is selected, but the demo CSV files are missing. "
            "Add the files under data/demo_2026_06_22/ in GitHub, or switch to Upload custom CSV files."
        )
    else:
        st.info("Upload v3_da_revenue CSV to start. v4/v5 and raw EPEX files can be added after that.")
    st.stop()

try:
    input_df, sources, v3_norm_df = build_input_data(v3_file.getvalue(), asset_label)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not build input data from v3 file: {exc}")
    st.stop()

v4_id_base, v4_info, v4_norm_df, v4_score_df = load_contribution_file(v4_file, input_df, asset, purpose="id")
v5_imbalance_base, v5_info, v5_norm_df, v5_score_df = load_contribution_file(v5_file, input_df, asset, purpose="imbalance")

id3_benchmark_price, id3_benchmark_info = build_id3_benchmark_price(
    input_df=input_df,
    v4_norm_df=v4_norm_df,
    v5_norm_df=v5_norm_df,
    v4_id_base=v4_id_base,
)

vintage_prices, epex_info_df = load_raw_vintage_prices(
    epex_files,
    input_df,
    id3_benchmark_price=id3_benchmark_price,
    id3_benchmark_info=id3_benchmark_info,
)

settlement_df, auction_breakdown_df, kpi = run_simulation(
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

source_rows = [
    {
        "item": "Data mode",
        "status": data_mode,
        "selected_column": "built-in demo files" if data_mode == "Use built-in demo data" else "uploaded files",
        "sum": np.nan,
    },
    {
        "item": "v3 DA revenue",
        "status": "ok",
        "selected_column": sources.get("DA revenue column"),
        "sum": float(input_df["da_revenue_eur_5min"].sum()),
    },
    {
        "item": "v4 ID contribution",
        "status": "ok" if v4_info.get("selected_column") not in ["not uploaded", "not detected"] else str(v4_info.get("selected_column")),
        "selected_column": v4_info.get("selected_column"),
        "sum": v4_info.get("sum", 0.0),
    },
    {
        "item": "v5 imbalance settlement EUR contribution",
        "status": v5_info.get("source_type", "ok") if v5_info.get("selected_column") not in ["not uploaded", "not detected"] else str(v5_info.get("selected_column")),
        "selected_column": v5_info.get("selected_column"),
        "sum": v5_info.get("sum", 0.0),
    },
    {
        "item": "No-ID imbalance settlement EUR",
        "status": "explicit benchmark used for DA-only comparison",
        "selected_column": "no_id_imbalance_settlement_eur",
        "sum": kpi.get("No-ID imbalance settlement EUR", 0.0),
    },
    {
        "item": "ID3 Benchmark Price",
        "status": id3_benchmark_info.get("status", "not loaded"),
        "selected_column": id3_benchmark_info.get("selected_column", ""),
        "sum": id3_benchmark_info.get("avg_price", np.nan),
    },
    {
        "item": "ID revenue source used",
        "status": str(settlement_df["id_revenue_source"].iloc[0]),
        "selected_column": "",
        "sum": float(settlement_df["id_revenue_eur"].sum()),
    },
    {
        "item": "Imbalance Settlement EUR source used",
        "status": str(settlement_df["imbalance_source"].iloc[0]),
        "selected_column": "",
        "sum": float(settlement_df["imbalance_settlement_eur"].sum()),
    },
]
source_status_df = pd.DataFrame(source_rows)

st.subheader(f"Trader Board Summary - {asset_label}")

kpi_cols = st.columns(4)
kpi_cols[0].metric("DA Revenue", fmt_eur(kpi["DA revenue EUR"]))
kpi_cols[1].metric("ID Revenue", fmt_eur(kpi["ID revenue EUR"]))
kpi_cols[2].metric("Imbalance Settlement EUR", fmt_eur(kpi["Imbalance Settlement EUR"]))
kpi_cols[3].metric("Total Revenue", fmt_eur(kpi["Total revenue EUR"]))

kpi_cols2 = st.columns(4)
kpi_cols2[0].metric("No-ID Benchmark", fmt_eur(kpi["No-ID benchmark EUR"]))
kpi_cols2[1].metric("ID Strategy Value", fmt_eur(kpi["ID strategy value EUR"]))
kpi_cols2[2].metric("Executed ID Trade", fmt_mwh(kpi["Total executed ID trade MWh"]))
kpi_cols2[3].metric("Remaining Imbalance", fmt_mwh(kpi["Remaining imbalance MWh"]))

with st.expander("Current data-source status", expanded=True):
    st.dataframe(source_status_df, use_container_width=True)
    if str(settlement_df["id_revenue_source"].iloc[0]).startswith("uploaded_v4"):
        st.caption(
            "Note: ID Revenue uses the uploaded v4 benchmark contribution. "
            "Executed ID Trade MWh is an approximate volume based on forecast-error exposure × auction capture, "
            "so it remains visible even when raw EPEX price files are not loaded."
        )
    if v5_info.get("source_type") == "converted_from_no_correction_total_pnl":
        st.caption(
            "v5 conversion note: the selected v5 column is total no-correction PnL, not imbalance settlement alone. "
            "The app converts it as Imbalance Settlement EUR = selected v5 total PnL - v3 DA Revenue. "
            f"Converted daily imbalance settlement = {float(v5_info.get('sum', 0.0)):,.0f} EUR."
        )

st.divider()

left, right = st.columns([0.85, 1.15])
with left:
    st.markdown("### Trader Decision Board")
    decision_rows = [
        {"metric": "Asset", "value": asset.display_name},
        {"metric": "Strategy Mode", "value": strategy_mode},
        {"metric": "DA Sold %", "value": fmt_pct(da_sold_pct)},
        {"metric": "Data Mode", "value": data_mode},
        {"metric": "Contribution Source", "value": contribution_source_mode},
    ]
    st.dataframe(pd.DataFrame(decision_rows), use_container_width=True, hide_index=True)

with right:
    st.markdown("### Auction Settings / Results")
    st.dataframe(auction_display_table(auction_breakdown_df), use_container_width=True, hide_index=True)
    if not epex_files:
        st.caption(
            "Raw EPEX vintage price files are not loaded. ID3 Benchmark Price is loaded from the best available v3/v4/v5 source when possible; "
            "other raw vintage prices remain 'not loaded' until matching raw EPEX CSVs are uploaded."
        )

st.divider()

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "Graphs",
        "Trader Decision Board",
        "Settlement table",
        "Input check",
        "Column diagnostics",
        "Download",
    ]
)

with tab1:
    st.plotly_chart(make_da_chart(settlement_df), use_container_width=True, config=PLOTLY_CONFIG)
    st.plotly_chart(make_id_imbalance_chart(settlement_df), use_container_width=True, config=PLOTLY_CONFIG)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### ID Revenue by Auction Window")
        st.plotly_chart(make_auction_chart(auction_breakdown_df), use_container_width=True, config=PLOTLY_CONFIG)
    with c2:
        st.markdown("#### Cumulative Revenue Components")
        st.plotly_chart(make_cumulative_chart(settlement_df), use_container_width=True, config=PLOTLY_CONFIG)

with tab2:
    c1, c2 = st.columns([0.8, 1.2])
    with c1:
        st.markdown("#### Strategy Inputs")
        strategy_table = pd.DataFrame(
            [
                {"item": "Asset", "value": asset_label},
                {"item": "Strategy Mode", "value": strategy_mode},
                {"item": "DA Sold %", "value": fmt_pct(da_sold_pct)},
                {"item": "Data Mode", "value": data_mode},
                {"item": "ID / imbalance source", "value": contribution_source_mode},
                {"item": "Original forecast error", "value": fmt_mwh(kpi["Original forecast error MWh"])},
                {"item": "Strategy exposure", "value": fmt_mwh(kpi["Strategy exposure MWh"])},
            ]
        )
        st.dataframe(strategy_table, use_container_width=True, hide_index=True)
    with c2:
        st.markdown("#### KPI")
        kpi_table = pd.DataFrame([{"metric": k, "value": v} for k, v in kpi.items()])
        st.dataframe(kpi_table, use_container_width=True, hide_index=True)

    st.markdown("#### Auction Breakdown")
    st.dataframe(auction_display_table(auction_breakdown_df), use_container_width=True, hide_index=True)

    st.markdown("#### DA-only vs DA + ID Correction")
    st.dataframe(make_case_comparison_table(kpi), use_container_width=True, hide_index=True)
    st.plotly_chart(make_strategy_waterfall_chart(settlement_df, kpi), use_container_width=True, config=PLOTLY_CONFIG)

with tab3:
    display_cols = [
        "time_label",
        "da_revenue_eur",
        "id_revenue_eur",
        "no_id_imbalance_settlement_eur",
        "imbalance_settlement_eur",
        "residual_imbalance_ratio",
        "id_plus_imbalance_eur",
        "total_revenue_eur",
        "forecast_error_mwh",
        "strategy_error_mwh",
        "capture_factor",
        "total_id_trade_mwh",
        "total_id_trade_abs_mwh",
        "remaining_imbalance_mwh",
        "remaining_imbalance_abs_mwh",
    ]
    st.dataframe(settlement_df[display_cols], use_container_width=True, height=520)

with tab4:
    st.write("Detected v3 source columns")
    st.dataframe(pd.DataFrame([sources]).T.rename(columns={0: "detected_column"}), use_container_width=True)
    st.write("Input data preview")
    st.dataframe(input_df.head(30), use_container_width=True)

with tab5:
    st.write("v3 normalized columns")
    v3_columns_df = pd.DataFrame({"v3_columns": list(v3_norm_df.columns)})
    st.dataframe(v3_columns_df, use_container_width=True, height=220)

    st.write("v4 ID contribution detection")
    st.json(v4_info)
    if not v4_score_df.empty:
        st.dataframe(v4_score_df.head(40), use_container_width=True, height=300)
    elif v4_file is None:
        st.info("v4 file was not uploaded.")
    else:
        st.warning("No numeric v4 columns were detected.")

    st.write("v5 imbalance settlement EUR contribution detection")
    st.json(v5_info)
    if not v5_score_df.empty:
        st.dataframe(v5_score_df.head(40), use_container_width=True, height=300)
    elif v5_file is None:
        st.info("v5 file was not uploaded.")
    else:
        st.warning("No numeric v5 columns were detected.")

    st.write("ID3 benchmark / raw EPEX price mapping")
    st.dataframe(epex_info_df, use_container_width=True)
    if not epex_files:
        st.info("No raw EPEX vintage files uploaded. Only the ID3 Benchmark Price source is available if detected or implied.")

with tab6:
    v3_columns_df = pd.DataFrame({"v3_columns": list(v3_norm_df.columns)})
    excel_bytes = make_excel_download(
        input_df=input_df,
        settlement_df=settlement_df,
        kpi=kpi,
        source_status_df=source_status_df,
        auction_breakdown_df=auction_breakdown_df,
        v3_columns_df=v3_columns_df,
        v4_score_df=v4_score_df,
        v5_score_df=v5_score_df,
        epex_info_df=epex_info_df,
    )
    st.download_button(
        label="Download calculated result as Excel",
        data=excel_bytes,
        file_name="v17_streamlit_mvp_v3_11_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        label="Download settlement table as CSV",
        data=settlement_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="v17_streamlit_mvp_v3_11_settlement.csv",
        mime="text/csv",
    )
