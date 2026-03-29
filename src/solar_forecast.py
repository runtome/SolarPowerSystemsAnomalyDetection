#!/usr/bin/env python3
"""
Solar Power Forecasting with TimesFM — Site-Level Analysis

Uses real solar monitoring data (generation, irradiance, temperature, load)
to forecast solar production using TimesFM 2.5 with covariates.

Data layout (datasets/site_1/):
  gen_*.csv          — Solar production (kW), ~1-min intervals
  Irradiance_*.csv   — Solar irradiance (W/m²), ~1-min intervals
  Temp*.csv          — Ambient temperature (°C), ~1-min intervals
  load_*.csv         — Site consumption (kW), 15-min intervals

Pipeline:
  1. Load & resample all series to 15-min intervals (mean aggregation)
  2. Forward-fill small gaps, flag large gaps
  3. Split into context (training) and horizon (forecast evaluation)
  4. Forecast solar generation with irradiance + temperature as covariates
  5. Forecast load independently
  6. Compute net energy (generation - load)
  7. Produce 4-panel visualization + output CSV/JSON

Requires: pip install timesfm[xreg] torch numpy pandas matplotlib
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = REPO_ROOT / "datasets" / "site_1"
OUTPUT_DIR = Path(__file__).parent / "output"

# ── Parameters ───────────────────────────────────────────────────────────────
RESAMPLE_FREQ = "15min"
MAX_GAP_FILL = 4          # forward-fill up to 4 periods (1 hour)
HORIZON_DAYS = 7           # forecast 7 days ahead
CONTEXT_DAYS = 90          # use last 90 days as context
HORIZON_STEPS = HORIZON_DAYS * 24 * 4   # 15-min steps
CONTEXT_STEPS = CONTEXT_DAYS * 24 * 4


# ── Step 1: Load & Resample ─────────────────────────────────────────────────

def load_and_resample(filepath: Path, value_col: str) -> pd.Series:
    """Load CSV, parse datetime index, resample to 15-min mean."""
    df = pd.read_csv(filepath, parse_dates=["date/time"], index_col="date/time")
    df.columns = [value_col]
    # Remove duplicates, sort, resample
    df = df[~df.index.duplicated(keep="first")].sort_index()
    resampled = df[value_col].resample(RESAMPLE_FREQ).mean()
    return resampled


def load_site_data() -> pd.DataFrame:
    """Load all 4 series for site_1, align on common 15-min index."""
    print("Loading site_1 data...")

    # Find files by prefix
    gen_file = next(DATA_DIR.glob("gen*.csv"))
    irr_file = next(DATA_DIR.glob("Irradiance*.csv"))
    temp_file = next(DATA_DIR.glob("Temp*.csv"))
    load_file = next(DATA_DIR.glob("load*.csv"))

    gen = load_and_resample(gen_file, "generation_kw")
    irr = load_and_resample(irr_file, "irradiance_wm2")
    temp = load_and_resample(temp_file, "temperature_c")
    load = load_and_resample(load_file, "load_kw")

    # Combine on common index
    df = pd.DataFrame({
        "generation_kw": gen,
        "irradiance_wm2": irr,
        "temperature_c": temp,
        "load_kw": load,
    })

    print(f"  Combined shape: {df.shape}")
    print(f"  Date range: {df.index[0]} -> {df.index[-1]}")
    print(f"  Nulls before fill:\n{df.isnull().sum().to_dict()}")

    return df


# ── Step 2: Clean & Fill Gaps ────────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Forward-fill small gaps, clip negatives for generation, report quality."""
    # Forward fill small gaps
    df = df.ffill(limit=MAX_GAP_FILL)

    # For remaining nulls, interpolate linearly
    null_before = df.isnull().sum()
    df = df.interpolate(method="linear", limit_direction="both")

    # Solar generation can't be negative (meter artifacts)
    df["generation_kw"] = df["generation_kw"].clip(lower=0)
    df["irradiance_wm2"] = df["irradiance_wm2"].clip(lower=0)
    df["load_kw"] = df["load_kw"].clip(lower=0)

    null_after = df.isnull().sum()
    print(f"  Nulls after cleaning:\n{null_after.to_dict()}")

    # Drop any remaining null rows (edges)
    df = df.dropna()

    # Data quality stats
    total_expected = (df.index[-1] - df.index[0]) / pd.Timedelta(RESAMPLE_FREQ) + 1
    completeness = len(df) / total_expected * 100
    print(f"  Final shape: {df.shape}, completeness: {completeness:.1f}%")

    return df


# ── Step 3: Split Context / Horizon ──────────────────────────────────────────

def split_data(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Split into context (for model input) and horizon (for evaluation)."""
    # Use the last HORIZON_STEPS as evaluation, preceding CONTEXT_STEPS as context
    end_idx = len(df)
    horizon_start = end_idx - HORIZON_STEPS
    context_start = max(0, horizon_start - CONTEXT_STEPS)

    context = df.iloc[context_start:horizon_start].copy()
    horizon = df.iloc[horizon_start:end_idx].copy()

    print(f"\n  Context: {context.index[0]} -> {context.index[-1]} ({len(context)} steps)")
    print(f"  Horizon: {horizon.index[0]} -> {horizon.index[-1]} ({len(horizon)} steps)")

    return context, horizon


# ── Step 4 & 5: Forecast with TimesFM ────────────────────────────────────────

def _check_xreg_available() -> bool:
    """Check if the XReg (covariates) module is available."""
    try:
        from timesfm.utils import xreg_lib  # noqa: F401
        return True
    except ImportError:
        return False


def _round_up(n: int, multiple: int) -> int:
    """Round n up to the nearest multiple."""
    return ((n + multiple - 1) // multiple) * multiple


def forecast_with_timesfm(
    context: pd.DataFrame,
    horizon: pd.DataFrame,
) -> dict:
    """Run TimesFM forecasts for generation and load.

    If timesfm[xreg] is installed, uses irradiance + temperature as covariates
    for solar generation. Otherwise falls back to univariate forecasting.
    """
    import torch
    import timesfm

    torch.set_float32_matmul_precision("high")

    print("\nLoading TimesFM 2.5...")
    model = timesfm.TimesFM_2p5_200M_torch.from_pretrained(
        "google/timesfm-2.5-200m-pytorch"
    )

    horizon_len = len(horizon)
    gen_context = context["generation_kw"].values.astype(np.float32)
    load_context = context["load_kw"].values.astype(np.float32)
    use_xreg = _check_xreg_available()

    # Round max_context and max_horizon to model patch sizes
    # (input patch = 32, output patch = 128)
    max_ctx = min(len(gen_context), 1024)
    max_ctx = _round_up(max_ctx, 32)
    max_hor = _round_up(horizon_len, 128)

    if use_xreg:
        # ── Solar generation WITH covariates ─────────────────────────────
        print(f"Forecasting solar generation (horizon={horizon_len} steps)...")
        print("  Using covariates: irradiance, temperature (xreg mode)")

        irr_full = np.concatenate([
            context["irradiance_wm2"].values,
            horizon["irradiance_wm2"].values,
        ]).astype(np.float32)
        temp_full = np.concatenate([
            context["temperature_c"].values,
            horizon["temperature_c"].values,
        ]).astype(np.float32)

        model.compile(
            timesfm.ForecastConfig(
                max_context=max_ctx,
                max_horizon=max_hor,
                normalize_inputs=True,
                use_continuous_quantile_head=True,
                force_flip_invariance=True,
                infer_is_positive=True,
                fix_quantile_crossing=True,
                per_core_batch_size=32,
                return_backcast=True,  # required for xreg
            )
        )

        gen_point_list, gen_quant_list = model.forecast_with_covariates(
            inputs=[gen_context],
            dynamic_numerical_covariates={
                "irradiance": [irr_full],
                "temperature": [temp_full],
            },
            xreg_mode="timesfm + xreg",
        )
        gen_point = np.array(gen_point_list[0])[:horizon_len]
        gen_quantiles = np.array(gen_quant_list[0])[:horizon_len]

        # Recompile without return_backcast for univariate load forecast
        model.compile(
            timesfm.ForecastConfig(
                max_context=max_ctx,
                max_horizon=max_hor,
                normalize_inputs=True,
                use_continuous_quantile_head=True,
                force_flip_invariance=True,
                infer_is_positive=True,
                fix_quantile_crossing=True,
                per_core_batch_size=32,
            )
        )
    else:
        # ── Univariate fallback ──────────────────────────────────────────
        print("  Note: timesfm[xreg] not installed — using univariate forecast")
        print("  Install with: pip install jax scikit-learn  (for covariate support)")

        model.compile(
            timesfm.ForecastConfig(
                max_context=max_ctx,
                max_horizon=max_hor,
                normalize_inputs=True,
                use_continuous_quantile_head=True,
                force_flip_invariance=True,
                infer_is_positive=True,
                fix_quantile_crossing=True,
                per_core_batch_size=32,
            )
        )

        print(f"Forecasting solar generation (horizon={horizon_len} steps)...")
        gen_point_raw, gen_quant_raw = model.forecast(
            inputs=[gen_context],
            horizon=horizon_len,
        )
        gen_point = gen_point_raw[0][:horizon_len]
        gen_quantiles = gen_quant_raw[0, :horizon_len, :]

    # ── Load forecast (always univariate) ────────────────────────────────
    print("Forecasting load consumption...")
    load_point_raw, load_quant_raw = model.forecast(
        inputs=[load_context],
        horizon=horizon_len,
    )
    load_point = load_point_raw[0][:horizon_len]
    load_quantiles = load_quant_raw[0, :horizon_len, :]

    results = {
        "generation": {
            "point": np.clip(gen_point, 0, None),
            "q10": np.clip(gen_quantiles[:, 1], 0, None),
            "q90": gen_quantiles[:, 9],
        },
        "load": {
            "point": np.clip(load_point, 0, None),
            "q10": np.clip(load_quantiles[:, 1], 0, None),
            "q90": load_quantiles[:, 9],
        },
        "used_covariates": use_xreg,
    }

    return results


# ── Step 6: Net Energy ───────────────────────────────────────────────────────

def compute_net_energy(
    horizon: pd.DataFrame,
    forecasts: dict,
) -> dict:
    """Compute net energy = generation - load for both actual and forecast."""
    actual_net = horizon["generation_kw"].values - horizon["load_kw"].values
    forecast_net = forecasts["generation"]["point"] - forecasts["load"]["point"]

    return {
        "actual_net": actual_net,
        "forecast_net": forecast_net,
    }


# ── Step 7: Visualization ───────────────────────────────────────────────────

def create_visualization(
    context: pd.DataFrame,
    horizon: pd.DataFrame,
    forecasts: dict,
    net_energy: dict,
    used_covariates: bool = False,
) -> None:
    """Create 4-panel solar power system dashboard."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(2, 2, figsize=(18, 12), sharex=False)
    fig.suptitle(
        "Solar Power System Forecast — Site 1\n"
        f"Context: {context.index[0].strftime('%Y-%m-%d')} to "
        f"{context.index[-1].strftime('%Y-%m-%d')} | "
        f"Forecast: {horizon.index[0].strftime('%Y-%m-%d')} to "
        f"{horizon.index[-1].strftime('%Y-%m-%d')}",
        fontsize=14, fontweight="bold", y=1.02,
    )

    h_dates = horizon.index
    # Show last 3 days of context + full horizon
    ctx_tail = context.iloc[-3 * 96:]
    ctx_dates = ctx_tail.index

    # ── Panel (0,0): Solar Generation Forecast ──────────────────────────
    ax = axes[0, 0]
    ax.plot(ctx_dates, ctx_tail["generation_kw"], color="#f59e0b",
            lw=1.2, alpha=0.7, label="Context (actual)")
    ax.plot(h_dates, horizon["generation_kw"], color="#f59e0b",
            lw=1.5, label="Actual generation")
    ax.plot(h_dates, forecasts["generation"]["point"], color="#dc2626",
            lw=2, ls="--", label="Forecast")
    ax.fill_between(h_dates, forecasts["generation"]["q10"],
                    forecasts["generation"]["q90"],
                    alpha=0.2, color="#dc2626", label="80% CI")
    ax.axvline(horizon.index[0], color="gray", ls=":", lw=1, alpha=0.6)
    ax.set_ylabel("Solar Generation (kW)")
    gen_title = ("Solar Generation Forecast (with Irradiance + Temp covariates)"
                 if used_covariates else "Solar Generation Forecast (univariate)")
    ax.set_title(gen_title, fontweight="bold")
    ax.legend(fontsize=8, loc="upper right")
    ax.grid(True, alpha=0.2)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))

    # ── Panel (0,1): Load Consumption Forecast ──────────────────────────
    ax = axes[0, 1]
    ax.plot(ctx_dates, ctx_tail["load_kw"], color="#3b82f6",
            lw=1.2, alpha=0.7, label="Context (actual)")
    ax.plot(h_dates, horizon["load_kw"], color="#3b82f6",
            lw=1.5, label="Actual load")
    ax.plot(h_dates, forecasts["load"]["point"], color="#7c3aed",
            lw=2, ls="--", label="Forecast")
    ax.fill_between(h_dates, forecasts["load"]["q10"],
                    forecasts["load"]["q90"],
                    alpha=0.2, color="#7c3aed", label="80% CI")
    ax.axvline(horizon.index[0], color="gray", ls=":", lw=1, alpha=0.6)
    ax.set_ylabel("Load Consumption (kW)")
    ax.set_title("Load Consumption Forecast", fontweight="bold")
    ax.legend(fontsize=8, loc="upper right")
    ax.grid(True, alpha=0.2)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))

    # ── Panel (1,0): Covariates during horizon ──────────────────────────
    ax = axes[1, 0]
    ax2 = ax.twinx()
    ln1 = ax.plot(h_dates, horizon["irradiance_wm2"], color="#eab308",
                  lw=1.2, label="Irradiance (W/m²)")
    ax.set_ylabel("Irradiance (W/m²)", color="#eab308")
    ln2 = ax2.plot(h_dates, horizon["temperature_c"], color="#ef4444",
                   lw=1.2, alpha=0.8, label="Temperature (°C)")
    ax2.set_ylabel("Temperature (°C)", color="#ef4444")
    lns = ln1 + ln2
    labs = [l.get_label() for l in lns]
    ax.legend(lns, labs, fontsize=8, loc="upper right")
    ax.set_title("Covariates — Irradiance & Temperature (Forecast Period)",
                 fontweight="bold")
    ax.grid(True, alpha=0.2)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))

    # ── Panel (1,1): Net Energy (Generation - Load) ─────────────────────
    ax = axes[1, 1]
    actual_net = net_energy["actual_net"]
    forecast_net = net_energy["forecast_net"]
    ax.fill_between(h_dates, actual_net, 0,
                    where=(actual_net >= 0), alpha=0.4, color="#22c55e",
                    label="Actual surplus (gen > load)")
    ax.fill_between(h_dates, actual_net, 0,
                    where=(actual_net < 0), alpha=0.4, color="#ef4444",
                    label="Actual deficit (gen < load)")
    ax.plot(h_dates, forecast_net, color="black", lw=1.8, ls="--",
            label="Forecast net energy")
    ax.axhline(0, color="gray", lw=1, alpha=0.5)
    ax.set_ylabel("Net Energy (kW)")
    ax.set_title("Net Energy = Generation − Load", fontweight="bold")
    ax.legend(fontsize=8, loc="upper right")
    ax.grid(True, alpha=0.2)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))

    for row in range(2):
        for col in range(2):
            axes[row, col].tick_params(axis="x", rotation=30)

    plt.tight_layout()
    out_path = OUTPUT_DIR / "solar_forecast.png"
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"\nSaved visualization: {out_path}")


def create_daily_summary_chart(
    horizon: pd.DataFrame,
    forecasts: dict,
) -> None:
    """Create daily summary bar chart comparing actual vs forecast."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Aggregate to daily totals (kWh = sum of 15-min kW readings / 4)
    horizon_copy = horizon.copy()
    horizon_copy["gen_forecast"] = forecasts["generation"]["point"]
    horizon_copy["load_forecast"] = forecasts["load"]["point"]

    daily = horizon_copy.resample("D").sum() / 4  # kW * 15min -> kWh

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    fig.suptitle("Daily Energy Summary — Actual vs Forecast (kWh)",
                 fontsize=13, fontweight="bold")

    dates = daily.index.strftime("%m-%d")
    x = np.arange(len(daily))
    w = 0.35

    # Generation
    ax = axes[0]
    ax.bar(x - w/2, daily["generation_kw"], w, color="#f59e0b", label="Actual")
    ax.bar(x + w/2, daily["gen_forecast"], w, color="#dc2626", alpha=0.7, label="Forecast")
    ax.set_xticks(x)
    ax.set_xticklabels(dates, rotation=45)
    ax.set_ylabel("Energy (kWh)")
    ax.set_title("Daily Solar Generation")
    ax.legend()
    ax.grid(True, alpha=0.2, axis="y")

    # Load
    ax = axes[1]
    ax.bar(x - w/2, daily["load_kw"], w, color="#3b82f6", label="Actual")
    ax.bar(x + w/2, daily["load_forecast"], w, color="#7c3aed", alpha=0.7, label="Forecast")
    ax.set_xticks(x)
    ax.set_xticklabels(dates, rotation=45)
    ax.set_ylabel("Energy (kWh)")
    ax.set_title("Daily Load Consumption")
    ax.legend()
    ax.grid(True, alpha=0.2, axis="y")

    plt.tight_layout()
    out_path = OUTPUT_DIR / "daily_summary.png"
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"Saved daily summary: {out_path}")


# ── Step 8: Export Results ───────────────────────────────────────────────────

def export_results(
    horizon: pd.DataFrame,
    forecasts: dict,
    net_energy: dict,
    used_covariates: bool = False,
) -> None:
    """Export forecast results to CSV and metadata to JSON."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # CSV with actual + forecast
    out_df = pd.DataFrame({
        "datetime": horizon.index,
        "actual_generation_kw": horizon["generation_kw"].values,
        "forecast_generation_kw": forecasts["generation"]["point"],
        "gen_q10": forecasts["generation"]["q10"],
        "gen_q90": forecasts["generation"]["q90"],
        "actual_load_kw": horizon["load_kw"].values,
        "forecast_load_kw": forecasts["load"]["point"],
        "load_q10": forecasts["load"]["q10"],
        "load_q90": forecasts["load"]["q90"],
        "actual_net_kw": net_energy["actual_net"],
        "forecast_net_kw": net_energy["forecast_net"],
        "irradiance_wm2": horizon["irradiance_wm2"].values,
        "temperature_c": horizon["temperature_c"].values,
    })
    csv_path = OUTPUT_DIR / "solar_forecast_results.csv"
    out_df.to_csv(csv_path, index=False)
    print(f"Saved results CSV: {csv_path} ({len(out_df)} rows)")

    # Error metrics
    gen_actual = horizon["generation_kw"].values
    gen_pred = forecasts["generation"]["point"]
    load_actual = horizon["load_kw"].values
    load_pred = forecasts["load"]["point"]

    def calc_metrics(actual, pred):
        mae = np.mean(np.abs(actual - pred))
        rmse = np.sqrt(np.mean((actual - pred) ** 2))
        mask = actual > 0
        mape = np.mean(np.abs((actual[mask] - pred[mask]) / actual[mask])) * 100 if mask.any() else None
        return {"mae": round(float(mae), 2), "rmse": round(float(rmse), 2), "mape_pct": round(float(mape), 2) if mape else None}

    metadata = {
        "site": "site_1",
        "model": "TimesFM 2.5 (200M)",
        "context_days": CONTEXT_DAYS,
        "horizon_days": HORIZON_DAYS,
        "resample_freq": RESAMPLE_FREQ,
        "horizon_steps": HORIZON_STEPS,
        "forecast_period": {
            "start": str(horizon.index[0]),
            "end": str(horizon.index[-1]),
        },
        "covariates_used": {
            "generation_forecast": (
                ["irradiance (W/m²)", "temperature (°C)"] if used_covariates
                else "univariate (xreg not installed)"
            ),
            "load_forecast": "univariate",
            "xreg_mode": "timesfm + xreg" if used_covariates else "n/a",
        },
        "metrics": {
            "generation": calc_metrics(gen_actual, gen_pred),
            "load": calc_metrics(load_actual, load_pred),
        },
        "daily_totals_kwh": {
            "actual_generation": round(float(gen_actual.sum() / 4), 1),
            "forecast_generation": round(float(gen_pred.sum() / 4), 1),
            "actual_load": round(float(load_actual.sum() / 4), 1),
            "forecast_load": round(float(load_pred.sum() / 4), 1),
        },
    }

    meta_path = OUTPUT_DIR / "solar_forecast_metadata.json"
    with open(meta_path, "w") as f:
        json.dump(metadata, f, indent=2)
    print(f"Saved metadata: {meta_path}")


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 70)
    print("  SOLAR POWER FORECASTING WITH TIMESFM")
    print("=" * 70)

    # 1-2. Load & clean
    df = load_site_data()
    df = clean_data(df)

    # 3. Split
    context, horizon = split_data(df)

    # 4-5. Forecast
    forecasts = forecast_with_timesfm(context, horizon)

    # 6. Net energy
    net_energy = compute_net_energy(horizon, forecasts)

    # 7. Visualize
    create_visualization(context, horizon, forecasts, net_energy, forecasts.get("used_covariates", False))
    create_daily_summary_chart(horizon, forecasts)

    # 8. Export
    export_results(horizon, forecasts, net_energy, forecasts.get("used_covariates", False))

    print("\n" + "=" * 70)
    print("  SOLAR FORECAST COMPLETE")
    print("=" * 70)
    gen_mae = np.mean(np.abs(
        horizon["generation_kw"].values - forecasts["generation"]["point"]
    ))
    load_mae = np.mean(np.abs(
        horizon["load_kw"].values - forecasts["load"]["point"]
    ))
    print(f"""
Results:
  Generation MAE: {gen_mae:.1f} kW
  Load MAE:       {load_mae:.1f} kW

Output files:
  output/solar_forecast.png           — 4-panel forecast dashboard
  output/daily_summary.png            — Daily actual vs forecast bars
  output/solar_forecast_results.csv   — Full 15-min results
  output/solar_forecast_metadata.json — Metrics and config
""")


if __name__ == "__main__":
    main()
