"""Evaluation metrics and anomaly detection."""

import numpy as np
import pandas as pd
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score


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


def build_results_table(all_results: list) -> pd.DataFrame:
    """Build comparison DataFrame from evaluation results."""
    return pd.DataFrame(all_results).set_index("model").round(4)
