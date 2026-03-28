"""All visualization functions. Each saves to the appropriate output directory."""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from pathlib import Path


plt.style.use("seaborn-v0_8-whitegrid")
plt.rcParams["figure.figsize"] = (14, 5)
plt.rcParams["font.size"] = 12


def plot_train_test_split(train_df, test_df, target_col: str, save_path: Path):
    fig, ax = plt.subplots(figsize=(16, 4))
    ax.plot(train_df.index, train_df[target_col], color="blue", lw=0.5, label="Train")
    ax.plot(test_df.index, test_df[target_col], color="orange", lw=0.5, label="Test")
    ax.axvline(test_df.index[0], color="red", ls="--", label="Split")
    ax.set_title("Train / Test Split")
    ax.set_ylabel("Generation (kW)")
    ax.legend()
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_training_loss(train_losses, val_losses, model_name: str, save_path: Path):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(train_losses, label="Train Loss", color="blue")
    ax.plot(val_losses, label="Val Loss", color="orange")
    ax.set_title(f"{model_name} - Training Loss")
    ax.set_xlabel("Epoch")
    ax.set_ylabel("Loss (MSE)")
    ax.legend()
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_prediction_vs_actual(timestamps, actual, predicted,
                               model_name: str, save_path: Path):
    """Plot actual vs predicted generation."""
    fig, ax = plt.subplots(figsize=(16, 5))
    ax.plot(timestamps[:len(actual)], actual, "r-", lw=0.5, alpha=0.7, label="Actual")
    ax.plot(timestamps[:len(predicted)], predicted, "b-", lw=0.5, alpha=0.7, label="Predicted")
    ax.set_title(f"{model_name} - Actual vs Predicted Generation")
    ax.set_ylabel("Generation (kW)")
    ax.legend()
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_error_distribution(mae, threshold, model_name: str, save_path: Path):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.hist(mae, bins=80, color="steelblue", alpha=0.7, edgecolor="black", linewidth=0.3)
    ax.axvline(threshold, color="red", ls="--", lw=2, label=f"Threshold={threshold:.2f}")
    ax.set_title(f"{model_name} - Error Distribution")
    ax.set_xlabel("Prediction Error (MAE)")
    ax.set_ylabel("Count")
    ax.legend()
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_anomalies(timestamps, actual, predicted, anomaly_mask,
                   model_name: str, threshold: float, save_path: Path):
    """Full 3-panel anomaly plot: pred vs actual, MAE+threshold, anomaly markers."""
    n = min(len(timestamps), len(actual), len(predicted))
    ts = timestamps[:n]
    act = actual[:n]
    pred = predicted[:n]
    mask = anomaly_mask[:n]

    fig, axes = plt.subplots(3, 1, figsize=(16, 12))

    # Panel 1: Actual vs Predicted
    axes[0].plot(ts, act, "r-", lw=0.5, alpha=0.7, label="Actual")
    axes[0].plot(ts, pred, "b-", lw=0.5, alpha=0.7, label="Predicted")
    axes[0].set_title(f"{model_name} - Actual vs Predicted")
    axes[0].set_ylabel("Generation (kW)")
    axes[0].legend()

    # Panel 2: MAE with threshold
    mae = np.abs(act - pred)
    axes[1].plot(ts, mae, color="gray", lw=0.5)
    axes[1].axhline(threshold, color="red", ls="--", label=f"Threshold={threshold:.2f}")
    axes[1].set_title(f"{model_name} - Prediction Error")
    axes[1].set_ylabel("MAE")
    axes[1].legend()

    # Panel 3: Anomalies marked
    axes[2].plot(ts, act, color="gray", lw=0.5, label="Generation")
    idx = mask == 1
    if idx.sum() > 0:
        axes[2].scatter(ts[idx], act[idx], c="red", s=20, zorder=5,
                       label=f"Anomalies ({idx.sum()})")
    axes[2].set_title(f"{model_name} - Detected Anomalies")
    axes[2].set_ylabel("Generation (kW)")
    axes[2].legend()

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_isolation_forest(timestamps, generation, anomalies, scores,
                           model_name: str, save_path: Path):
    fig, axes = plt.subplots(2, 1, figsize=(16, 8))

    axes[0].plot(timestamps, generation, color="gray", lw=0.5, label="Generation")
    mask = anomalies == 1
    axes[0].scatter(timestamps[mask], generation[mask],
                   color="red", s=15, label=f"Anomalies ({mask.sum()})", zorder=5)
    axes[0].set_title(f"{model_name} - Anomaly Detection")
    axes[0].set_ylabel("Generation (kW)")
    axes[0].legend()

    axes[1].plot(timestamps, scores, color="blue", lw=0.5)
    axes[1].axhline(0, color="red", ls="--")
    axes[1].set_title("Anomaly Score (negative = anomalous)")
    axes[1].set_ylabel("Score")

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_autoencoder_results(timestamps, actual, recon_error, anomalies,
                              threshold, model_name: str, save_path: Path):
    fig, axes = plt.subplots(2, 1, figsize=(16, 8))

    n = min(len(timestamps), len(recon_error))
    ts = timestamps[:n]

    axes[0].plot(ts, recon_error[:n], color="gray", lw=0.5)
    axes[0].axhline(threshold, color="red", ls="--", label=f"Threshold={threshold:.4f}")
    axes[0].set_title(f"{model_name} - Reconstruction Error")
    axes[0].set_ylabel("MAE")
    axes[0].legend()

    axes[1].plot(ts, actual[:n], color="gray", lw=0.5, label="Generation")
    mask = anomalies[:n] == 1
    if mask.sum() > 0:
        axes[1].scatter(ts[mask], actual[:n][mask], c="red", s=15, zorder=5,
                       label=f"Anomalies ({mask.sum()})")
    axes[1].set_title(f"{model_name} - Detected Anomalies")
    axes[1].set_ylabel("Generation (kW)")
    axes[1].legend()

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_model_comparison(results_df: pd.DataFrame, save_path: Path):
    fig, axes = plt.subplots(1, 3, figsize=(16, 5))
    for ax, metric, color in zip(axes, ["RMSE", "MAE", "R2"],
                                  ["#2196F3", "#FF9800", "#4CAF50"]):
        results_df[metric].plot(kind="barh", ax=ax, color=color, edgecolor="black")
        ax.set_title(metric)
    plt.suptitle("Model Performance Comparison", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_anomaly_comparison(anomaly_counts: dict, save_path: Path):
    models = list(anomaly_counts.keys())
    counts = list(anomaly_counts.values())

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.barh(models, counts, color="crimson", edgecolor="black")
    ax.set_xlabel("Anomalies Detected")
    ax.set_title("Anomaly Detection Comparison")
    for bar, v in zip(bars, counts):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                str(v), va="center")
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_ensemble(timestamps, actual, votes, ensemble_anomalies,
                   min_votes: int, save_path: Path):
    n = min(len(timestamps), len(actual), len(votes))
    ts = timestamps[:n]

    fig, axes = plt.subplots(2, 1, figsize=(16, 10))

    axes[0].plot(ts, actual[:n], color="gray", lw=0.5, label="Generation")
    mask = ensemble_anomalies[:n] == 1
    if mask.sum() > 0:
        axes[0].scatter(ts[mask], actual[:n][mask], c="red", s=20, zorder=5,
                       label=f"Anomalies ({mask.sum()})")
    axes[0].set_title("Ensemble Anomaly Detection (Majority Vote)")
    axes[0].set_ylabel("Generation (kW)")
    axes[0].legend()

    axes[1].bar(ts, votes[:n], color="steelblue", width=0.01, alpha=0.7)
    axes[1].axhline(min_votes, color="red", ls="--", label=f"Threshold ({min_votes} votes)")
    axes[1].set_title("Model Agreement Count")
    axes[1].set_ylabel("Votes")
    axes[1].legend()

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()
