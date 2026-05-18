"""Evaluation metrics and anomaly detection."""

import numpy as np
import pandas as pd
from pathlib import Path

from sklearn.metrics import (
    mean_squared_error, mean_absolute_error, r2_score,
    precision_score, recall_score, f1_score, confusion_matrix,
)


def evaluate_model(actual: np.ndarray, predicted: np.ndarray, model_name: str) -> dict:
    """Compute regression metrics. Handles multi-step by averaging across steps."""
    if actual.ndim == 2:
        actual = actual.mean(axis=1)
    if predicted.ndim == 2:
        predicted = predicted.mean(axis=1)

    rmse = np.sqrt(mean_squared_error(actual, predicted))
    mae = mean_absolute_error(actual, predicted)
    r2 = r2_score(actual, predicted)

    print(f"  {model_name:20s} | RMSE: {rmse:8.4f} | MAE: {mae:8.4f} | R2: {r2:.4f}")
    return {"model": model_name, "RMSE": rmse, "MAE": mae, "R2": r2}


def detect_anomalies_3sigma(actual: np.ndarray, predicted: np.ndarray, sigma: float = 3.0):
    """Detect anomalies using 3-sigma rule on MAE.

    For multi-step predictions, averages across output steps first.
    """
    if actual.ndim == 2:
        actual = actual.mean(axis=1)
    if predicted.ndim == 2:
        predicted = predicted.mean(axis=1)

    mae = np.abs(actual - predicted)
    threshold = mae.mean() + sigma * mae.std()
    anomalies = (mae > threshold).astype(int)
    return mae, threshold, anomalies


def detect_anomalies_reconstruction(X_original, X_reconstructed, sigma: float = 3.0,
                                     train_error: float = None, train_std: float = None):
    """Detect anomalies via autoencoder reconstruction error.

    Args:
        X_original: (samples, timesteps, features)
        X_reconstructed: (samples, timesteps, features)
        sigma: number of standard deviations for threshold
        train_error: pre-computed mean training error (for threshold)
        train_std: pre-computed std of training error

    Returns:
        recon_error, threshold, anomalies
    """
    recon_error = np.mean(np.abs(X_original - X_reconstructed), axis=(1, 2))

    if train_error is not None and train_std is not None:
        threshold = train_error + sigma * train_std
    else:
        threshold = recon_error.mean() + sigma * recon_error.std()

    anomalies = (recon_error > threshold).astype(int)
    return recon_error, threshold, anomalies


def ensemble_voting(anomaly_arrays: dict, min_votes: int = 3):
    """Combine anomaly predictions from multiple models via majority voting.

    Args:
        anomaly_arrays: dict of {model_name: anomaly_array}
        min_votes: minimum number of models that must agree

    Returns:
        votes_per_point, ensemble_anomalies, vote_distribution
    """
    # Find minimum length across all arrays
    min_len = min(len(a) for a in anomaly_arrays.values())
    matrix = np.column_stack([a[:min_len] for a in anomaly_arrays.values()])

    votes = matrix.sum(axis=1)
    ensemble = (votes >= min_votes).astype(int)

    distribution = {}
    for v in range(len(anomaly_arrays) + 1):
        count = (votes == v).sum()
        if count > 0:
            distribution[v] = int(count)

    return votes, ensemble, distribution


def classify_anomalies(anomaly_mask: np.ndarray, timestamps,
                        test_df: pd.DataFrame) -> dict:
    """Classify detected anomalies into 5 categories.

    Uses the aligned test_df (irradiance, temperature, generation) to determine
    which physical category each anomaly belongs to.

    Categories:
        1. High Irr / Low Gen    - high irradiance but low generation
        2. Sudden Drop           - large drop from previous consecutive step
        3. Efficiency Decline    - efficiency below rolling trend
        4. Gen Spike             - generation far above normal for irradiance bin
        5. Gen / Zero Irr        - generation recorded with zero irradiance

    Returns dict of {category_name: boolean_mask} (same length as anomaly_mask).
    """
    n = len(anomaly_mask)
    ts = timestamps[:n]
    anom_idx = anomaly_mask[:n] == 1

    # Build aligned DataFrame for classification
    irr_col = "irradiance_wm2"
    gen_col = "generation_kw"
    has_irr = irr_col in test_df.columns
    has_gen = gen_col in test_df.columns

    # Align test_df to timestamps
    aligned = test_df.reindex(ts)

    categories = {}
    if not (has_irr and has_gen) or len(aligned) == 0:
        categories["Anomaly"] = anom_idx
        return categories

    irr = aligned[irr_col].values
    gen = aligned[gen_col].values

    # Start with all anomalies unclassified
    classified = np.zeros(n, dtype=bool)

    # 5. Gen / Zero Irr  (check first - most definitive)
    cat5 = anom_idx & (irr <= 0) & (gen > 0)
    categories["Gen / Zero Irr"] = cat5
    classified |= cat5

    # 1. High Irr / Low Gen
    daytime_mask = irr > 50
    if daytime_mask.sum() > 0:
        irr_q75 = np.nanpercentile(irr[daytime_mask], 75)
        gen_q15 = np.nanpercentile(gen[daytime_mask], 15)
        cat1 = anom_idx & ~classified & (irr > irr_q75) & (gen < gen_q15)
    else:
        cat1 = np.zeros(n, dtype=bool)
    categories["High Irr / Low Gen"] = cat1
    classified |= cat1

    # 2. Sudden Drop  (>50% drop from previous step within 20 min)
    gen_diff = pd.Series(gen, index=ts).diff()
    time_diff = pd.Series(ts).diff().dt.total_seconds().values
    prev_gen = pd.Series(gen, index=ts).shift(1)
    pct_drop = gen_diff.values / np.where(prev_gen.values > 0, prev_gen.values, np.nan)
    consecutive = np.zeros(n, dtype=bool)
    consecutive[1:] = time_diff[1:] <= 1200  # 20 min
    cat2 = (anom_idx & ~classified & consecutive
            & (pct_drop < -0.5) & (np.abs(gen_diff.values) > 30))
    categories["Sudden Drop"] = cat2
    classified |= cat2

    # 3. Efficiency Decline (daytime efficiency below rolling trend)
    with np.errstate(divide="ignore", invalid="ignore"):
        eff = np.where(irr > 50, gen / irr, np.nan)
    eff_series = pd.Series(eff, index=ts)
    roll_mean = eff_series.rolling(96, min_periods=10).mean()
    roll_std = eff_series.rolling(96, min_periods=10).std()
    below_trend = eff_series < (roll_mean - 2 * roll_std)
    cat3 = anom_idx & ~classified & below_trend.fillna(False).values & daytime_mask
    categories["Efficiency Decline"] = cat3
    classified |= cat3

    # 4. Gen Spike (generation above 3*std for its irradiance bin)
    cat4 = np.zeros(n, dtype=bool)
    if daytime_mask.sum() > 20:
        try:
            irr_daytime = irr.copy()
            irr_daytime[~daytime_mask] = np.nan
            bins = pd.qcut(pd.Series(irr_daytime).dropna(), q=10, duplicates="drop")
            bin_series = pd.qcut(pd.Series(irr_daytime), q=10, duplicates="drop")
            bin_stats = pd.Series(gen).groupby(bin_series).agg(["mean", "std"])
            for idx in range(n):
                if anom_idx[idx] and not classified[idx] and daytime_mask[idx]:
                    b = bin_series.iloc[idx]
                    if pd.notna(b) and b in bin_stats.index:
                        if gen[idx] > bin_stats.loc[b, "mean"] + 3 * bin_stats.loc[b, "std"]:
                            cat4[idx] = True
        except Exception:
            pass
    categories["Gen Spike"] = cat4
    classified |= cat4

    # Remaining unclassified anomalies go into the largest applicable bucket
    remaining = anom_idx & ~classified
    if remaining.sum() > 0:
        # Assign to Efficiency Decline as catch-all for model-detected anomalies
        categories["Efficiency Decline"] = categories["Efficiency Decline"] | remaining

    return categories


def build_results_table(all_results: list) -> pd.DataFrame:
    """Build comparison DataFrame from evaluation results."""
    return pd.DataFrame(all_results).set_index("model").round(4)


_LABEL_BOOL_COLS = ["HighIrr-LowGen", "SuddenDrop", "EfficiencyDecline", "GenSpike", "Gen-ZeroIrr"]


def align_labels_to_model(anomaly_mask: np.ndarray, timestamps, label_csv) -> tuple:
    """Return (y_true, y_pred) aligned to model timestamps from the label CSV.

    Returns (None, None) if label CSV doesn't exist or has no matching columns.
    """
    label_csv = Path(label_csv)
    if not label_csv.exists():
        return None, None

    label_df = pd.read_csv(label_csv, index_col=0, parse_dates=True)
    bool_cols = [c for c in _LABEL_BOOL_COLS if c in label_df.columns]
    if not bool_cols:
        return None, None

    gt = label_df.reindex(timestamps)
    y_true = gt[bool_cols].any(axis=1).fillna(False).astype(int).values
    y_pred = anomaly_mask.astype(int)
    min_len = min(len(y_true), len(y_pred))
    return y_true[:min_len], y_pred[:min_len]


def print_confusion_matrix(y_true: np.ndarray, y_pred: np.ndarray, model_name: str):
    """Print 2×2 text confusion matrix with precision / recall / F1."""
    cm = confusion_matrix(y_true, y_pred, labels=[0, 1])
    tn, fp, fn, tp = cm.ravel()
    prec = precision_score(y_true, y_pred, zero_division=0)
    rec  = recall_score(y_true, y_pred, zero_division=0)
    f1   = f1_score(y_true, y_pred, zero_division=0)

    print(f"\n  Confusion Matrix ({model_name} vs Labels):")
    print(f"  {'':22s}  {'Pred Normal':>12} {'Pred Anomaly':>13}")
    print(f"  {'Actual Normal':22s}  {tn:>12,} {fp:>13,}")
    print(f"  {'Actual Anomaly':22s}  {fn:>12,} {tp:>13,}")
    print(f"  Precision: {prec:.4f}  Recall: {rec:.4f}  F1: {f1:.4f}")


def evaluate_anomaly_performance(
    anomaly_store: dict,
    timestamps_store: dict,
    label_csv,
) -> pd.DataFrame | None:
    """Compare per-model anomaly predictions against ground truth label CSV.

    Returns DataFrame (index=model) with columns: precision, recall, f1, tp, fp, tn, fn.
    Returns None if label CSV is missing.
    """
    label_csv = Path(label_csv)
    if not label_csv.exists():
        return None

    label_df = pd.read_csv(label_csv, index_col=0, parse_dates=True)
    bool_cols = [c for c in _LABEL_BOOL_COLS if c in label_df.columns]
    if not bool_cols:
        return None

    rows = []
    for model_name, anomaly_mask in anomaly_store.items():
        ts = timestamps_store.get(model_name)
        if ts is None:
            continue

        gt = label_df.reindex(ts)
        y_true = gt[bool_cols].any(axis=1).fillna(False).astype(int).values
        y_pred = anomaly_mask.astype(int)

        min_len = min(len(y_true), len(y_pred))
        y_true = y_true[:min_len]
        y_pred = y_pred[:min_len]

        cm = confusion_matrix(y_true, y_pred, labels=[0, 1])
        tn, fp, fn, tp = cm.ravel()

        rows.append({
            "model":     model_name,
            "precision": round(precision_score(y_true, y_pred, zero_division=0), 4),
            "recall":    round(recall_score(y_true, y_pred, zero_division=0), 4),
            "f1":        round(f1_score(y_true, y_pred, zero_division=0), 4),
            "tp": int(tp), "fp": int(fp), "tn": int(tn), "fn": int(fn),
        })

    if not rows:
        return None
    return pd.DataFrame(rows).set_index("model")


def print_anomaly_performance(metrics_df: pd.DataFrame):
    """Print formatted anomaly detection performance table."""
    print("\n" + "=" * 74)
    print("  Anomaly Detection Performance (vs Ground Truth Labels)")
    print("=" * 74)
    hdr = f"  {'Model':<22} {'Precision':>10} {'Recall':>8} {'F1':>8}  {'TP':>5} {'FP':>5} {'TN':>7} {'FN':>5}"
    print(hdr)
    print("  " + "-" * 72)
    for model, row in metrics_df.iterrows():
        print(
            f"  {model:<22} {row['precision']:>10.4f} {row['recall']:>8.4f}"
            f" {row['f1']:>8.4f}  {row['tp']:>5} {row['fp']:>5}"
            f" {row['tn']:>7} {row['fn']:>5}"
        )
    print("=" * 74)
