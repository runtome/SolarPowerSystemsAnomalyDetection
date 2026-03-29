"""
TimesFM foundation model wrapper for Solar Anomaly Detection pipeline.

TimesFM is a pretrained time-series foundation model — no training needed.
This module handles loading the model and performing rolling forecasts
on the test period using actual history as context.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from src.config import TimesFMConfig, DataConfig


def is_timesfm_available() -> bool:
    """Check if timesfm package is installed."""
    try:
        import timesfm  # noqa: F401
        return True
    except ImportError:
        return False


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


def load_timesfm_model(cfg: TimesFMConfig, use_xreg: bool = False):
    """Load pretrained TimesFM model and compile with ForecastConfig."""
    import torch
    import timesfm

    torch.set_float32_matmul_precision("high")

    print("  Loading TimesFM 2.5 (200M)...")
    model = timesfm.TimesFM_2p5_200M_torch.from_pretrained(cfg.model_id)

    max_ctx = _round_up(cfg.context_length, 32)
    max_hor = _round_up(cfg.max_horizon, 128)

    compile_kwargs = dict(
        max_context=max_ctx,
        max_horizon=max_hor,
        normalize_inputs=True,
        use_continuous_quantile_head=True,
        force_flip_invariance=True,
        infer_is_positive=True,
        fix_quantile_crossing=True,
        per_core_batch_size=32,
    )

    if use_xreg:
        compile_kwargs["return_backcast"] = True

    model.compile(timesfm.ForecastConfig(**compile_kwargs))
    print(f"  Model compiled (context={max_ctx}, horizon={max_hor}, xreg={use_xreg})")

    return model


def rolling_forecast(model, all_values: np.ndarray, test_start_idx: int,
                     n_test: int, max_horizon: int, context_length: int) -> np.ndarray:
    """Univariate rolling forecast over the test period.

    Uses actual history as context (not model's own predictions).
    This is correct for retrospective anomaly detection.

    Args:
        model: Compiled TimesFM model.
        all_values: Full time series (train + test), 1D float32.
        test_start_idx: Index where test period begins.
        n_test: Number of test steps to forecast.
        max_horizon: Max steps per forecast chunk.
        context_length: Number of trailing actual values as context.

    Returns:
        predictions: 1D array of length n_test.
    """
    predictions = []
    pos = test_start_idx

    while len(predictions) < n_test:
        remaining = n_test - len(predictions)
        chunk_size = min(max_horizon, remaining)

        # Context: actual values ending just before current position
        ctx_start = max(0, pos - context_length)
        context = all_values[ctx_start:pos].astype(np.float32)

        point_forecast, _ = model.forecast(
            inputs=[context],
            horizon=chunk_size,
        )
        chunk_pred = point_forecast[0][:chunk_size]
        predictions.extend(chunk_pred.tolist())

        pos += chunk_size

    return np.array(predictions[:n_test], dtype=np.float32)


def rolling_forecast_with_covariates(model, all_gen: np.ndarray,
                                     all_irr: np.ndarray, all_temp: np.ndarray,
                                     test_start_idx: int, n_test: int,
                                     max_horizon: int, context_length: int) -> np.ndarray:
    """Rolling forecast with irradiance + temperature covariates.

    Args:
        model: Compiled TimesFM model (with return_backcast=True).
        all_gen: Full generation series (train + test), 1D float32.
        all_irr: Full irradiance series (train + test), 1D float32.
        all_temp: Full temperature series (train + test), 1D float32.
        test_start_idx: Index where test period begins.
        n_test: Number of test steps to forecast.
        max_horizon: Max steps per forecast chunk.
        context_length: Number of trailing actual values as context.

    Returns:
        predictions: 1D array of length n_test.
    """
    predictions = []
    pos = test_start_idx

    while len(predictions) < n_test:
        remaining = n_test - len(predictions)
        chunk_size = min(max_horizon, remaining)

        # Context window
        ctx_start = max(0, pos - context_length)
        gen_context = all_gen[ctx_start:pos].astype(np.float32)

        # Covariates: context + horizon period
        irr_cov = all_irr[ctx_start:pos + chunk_size].astype(np.float32)
        temp_cov = all_temp[ctx_start:pos + chunk_size].astype(np.float32)

        point_forecast, _ = model.forecast_with_covariates(
            inputs=[gen_context],
            dynamic_numerical_covariates={
                "irradiance": [irr_cov],
                "temperature": [temp_cov],
            },
            xreg_mode="timesfm + xreg",
        )
        chunk_pred = np.array(point_forecast[0])[:chunk_size]
        predictions.extend(chunk_pred.tolist())

        pos += chunk_size

    return np.array(predictions[:n_test], dtype=np.float32)


def forecast_test_period(train_df: pd.DataFrame, test_df: pd.DataFrame,
                         cfg: TimesFMConfig, data_cfg: DataConfig):
    """Run TimesFM rolling forecast on the test period.

    Args:
        train_df: Training DataFrame with datetime index.
        test_df: Test DataFrame with datetime index.
        cfg: TimesFM configuration.
        data_cfg: Data configuration (for feature/target names).

    Returns:
        actual: 1D array of actual generation values (test period).
        predicted: 1D array of predicted generation values (test period).
    """
    target = data_cfg.target
    n_test = len(test_df)
    n_train = len(train_df)

    # Concat train + test generation (raw, unscaled)
    all_gen = np.concatenate([
        train_df[target].values,
        test_df[target].values,
    ]).astype(np.float32)

    # Check if covariates are available and desired
    use_xreg = cfg.use_covariates and _check_xreg_available()

    if use_xreg:
        print("  Using covariates: irradiance, temperature (xreg mode)")
        all_irr = np.concatenate([
            train_df["irradiance_wm2"].values,
            test_df["irradiance_wm2"].values,
        ]).astype(np.float32)
        all_temp = np.concatenate([
            train_df["temperature_c"].values,
            test_df["temperature_c"].values,
        ]).astype(np.float32)

        model = load_timesfm_model(cfg, use_xreg=True)
        predicted = rolling_forecast_with_covariates(
            model, all_gen, all_irr, all_temp,
            test_start_idx=n_train,
            n_test=n_test,
            max_horizon=cfg.max_horizon,
            context_length=cfg.context_length,
        )
    else:
        if cfg.use_covariates:
            print("  Note: xreg not available, falling back to univariate")
        else:
            print("  Using univariate forecasting")

        model = load_timesfm_model(cfg, use_xreg=False)
        predicted = rolling_forecast(
            model, all_gen,
            test_start_idx=n_train,
            n_test=n_test,
            max_horizon=cfg.max_horizon,
            context_length=cfg.context_length,
        )

    # Clip negative predictions (generation can't be negative)
    predicted = np.clip(predicted, 0, None)

    actual = test_df[target].values.astype(np.float32)

    return actual, predicted
