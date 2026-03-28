#!/usr/bin/env python
"""
Main entry point for Solar Anomaly Detection training pipeline.

Usage examples:
    # Default (auto trial name: run_1, run_2, ...)
    python run_training.py

    # Named trial
    python run_training.py --trial-name baseline

    # Custom timesteps with trial name
    python run_training.py --trial-name exp_6step --input-steps 6 --output-steps 2

    # Specific models
    python run_training.py --trial-name lstm_only --models LSTM

    # Adjust training
    python run_training.py --trial-name tuned --epochs 200 --batch-size 64 --lr 0.0005
"""

import argparse
import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

import joblib
import numpy as np
import torch

warnings.filterwarnings("ignore")


class TeeLogger:
    """Duplicates writes to both a file and the original stream."""

    def __init__(self, filepath, stream):
        self.file = open(filepath, "w", encoding="utf-8")
        self.stream = stream

    def write(self, data):
        self.stream.write(data)
        self.file.write(data)
        self.file.flush()

    def flush(self):
        self.stream.flush()
        self.file.flush()

    def close(self):
        self.file.close()


def resolve_trial_name(output_dir: str, trial_name: str | None) -> str:
    """Resolve trial name. Auto-generate run_1, run_2, ... if not specified."""
    if trial_name:
        return trial_name

    base = Path(output_dir)
    if not base.exists():
        return "run_1"

    existing = []
    for d in base.iterdir():
        if d.is_dir() and d.name.startswith("run_"):
            try:
                num = int(d.name.split("_", 1)[1])
                existing.append(num)
            except ValueError:
                pass

    next_num = max(existing, default=0) + 1
    return f"run_{next_num}"


def setup_logging(args, trial_name: str):
    """Set up logging to both console and a log file in logs/."""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    model_tag = "_".join(args.models) if args.models != ["all"] else "all"
    filename = f"{trial_name}_{timestamp}_in{args.input_steps}_out{args.output_steps}_{model_tag}.log"
    log_path = log_dir / filename

    tee_stdout = TeeLogger(log_path, sys.stdout)
    tee_stderr = TeeLogger(log_path, sys.stderr)
    sys.stdout = tee_stdout
    sys.stderr = tee_stderr

    print(f"  Log file: {log_path}")
    return tee_stdout, tee_stderr


from src.config import (
    DataConfig, SequenceConfig, TrainConfig, ModelConfig,
    AnomalyConfig, PathConfig, DL_MODELS, ML_MODELS, ALL_MODELS,
)
from src.data import SolarDataModule
from src.models import LSTMModel, CNNLSTM, LSTMAutoencoder, TransformerModel
from src.models.ml_models import (
    train_isolation_forest, predict_isolation_forest, train_random_forest,
)
from src.trainer import train_model, predict
from src.evaluate import (
    evaluate_model, detect_anomalies_3sigma, detect_anomalies_reconstruction,
    ensemble_voting, build_results_table,
)
from src.eda import run_eda
from src.visualize import (
    plot_train_test_split, plot_training_loss, plot_prediction_vs_actual,
    plot_error_distribution, plot_anomalies, plot_isolation_forest,
    plot_autoencoder_results, plot_model_comparison, plot_anomaly_comparison,
    plot_ensemble,
)


def build_dl_model(name: str, n_features: int, seq_cfg: SequenceConfig,
                   model_cfg: ModelConfig, device: torch.device):
    """Instantiate a DL model by name."""
    inp = seq_cfg.input_steps
    out = seq_cfg.output_steps

    if name == "LSTM":
        return LSTMModel(
            n_features=n_features, output_steps=out,
            hidden_size=model_cfg.lstm_hidden,
            num_layers=model_cfg.lstm_layers,
            dropout=model_cfg.lstm_dropout,
        ).to(device)

    elif name == "CNN_LSTM":
        return CNNLSTM(
            n_features=n_features, output_steps=out,
            cnn_filters=model_cfg.cnn_filters,
            lstm_hidden=model_cfg.cnn_lstm_hidden,
        ).to(device)

    elif name == "LSTM_Autoencoder":
        return LSTMAutoencoder(
            n_features=n_features, input_steps=inp,
            hidden_enc=model_cfg.ae_hidden_enc,
            bottleneck=model_cfg.ae_bottleneck,
            dropout=model_cfg.ae_dropout,
        ).to(device)

    elif name == "Transformer":
        return TransformerModel(
            n_features=n_features, output_steps=out,
            d_model=model_cfg.tf_d_model,
            nhead=model_cfg.tf_nhead,
            num_layers=model_cfg.tf_num_layers,
            dim_feedforward=model_cfg.tf_ff_dim,
            dropout=model_cfg.tf_dropout,
        ).to(device)

    else:
        raise ValueError(f"Unknown DL model: {name}")


def run_ml_models(data_mod: SolarDataModule, model_cfg: ModelConfig,
                  anomaly_cfg: AnomalyConfig, path_cfg: PathConfig):
    """Train and evaluate ML models."""
    results = []
    anomaly_store = {}

    # --- Isolation Forest ---
    name = "Isolation_Forest"
    print(f"\n{'='*60}\n  Training: {name}\n{'='*60}")

    iso_model = train_isolation_forest(data_mod.train_df, model_cfg)
    iso_anomalies, iso_scores = predict_isolation_forest(iso_model, data_mod.test_df)

    # Save model
    joblib.dump(iso_model, path_cfg.model_dir(name) / f"{name}.joblib")
    print(f"  Anomalies: {iso_anomalies.sum()} ({iso_anomalies.sum()/len(iso_anomalies)*100:.1f}%)")

    # Plot
    plot_isolation_forest(
        data_mod.test_df.index,
        data_mod.test_df[data_mod.data_cfg.target].values,
        iso_anomalies, iso_scores, name,
        path_cfg.plots_dir(name) / "anomaly_detection.png",
    )
    anomaly_store[name] = iso_anomalies

    # --- Random Forest ---
    name = "Random_Forest"
    print(f"\n{'='*60}\n  Training: {name}\n{'='*60}")

    features = data_mod.data_cfg.features
    target = data_mod.data_cfg.target

    rf_model = train_random_forest(
        data_mod.train_df[features], data_mod.train_df[target], model_cfg
    )
    rf_preds = rf_model.predict(data_mod.test_df[features])
    y_test = data_mod.test_df[target].values

    # Save model
    joblib.dump(rf_model, path_cfg.model_dir(name) / f"{name}.joblib")

    # Evaluate
    rf_result = evaluate_model(y_test, rf_preds, name)
    results.append(rf_result)

    # Feature importance
    imp = dict(zip(features, rf_model.feature_importances_))
    print(f"  Feature importance: {imp}")

    # Anomaly detection
    rf_mae, rf_thresh, rf_anomalies = detect_anomalies_3sigma(
        y_test, rf_preds, anomaly_cfg.sigma
    )
    print(f"  Anomalies: {rf_anomalies.sum()} ({rf_anomalies.sum()/len(rf_anomalies)*100:.1f}%)")

    # Plots
    pdir = path_cfg.plots_dir(name)
    plot_anomalies(data_mod.test_df.index, y_test, rf_preds,
                   rf_anomalies, name, rf_thresh, pdir / "anomaly_detection.png")
    plot_prediction_vs_actual(data_mod.test_df.index, y_test, rf_preds,
                              name, pdir / "prediction_vs_actual.png")
    plot_error_distribution(rf_mae, rf_thresh, name, pdir / "error_distribution.png")

    anomaly_store[name] = rf_anomalies

    return results, anomaly_store


def run_dl_models(model_names: list, data_mod: SolarDataModule,
                  seq_cfg: SequenceConfig, train_cfg: TrainConfig,
                  model_cfg: ModelConfig, anomaly_cfg: AnomalyConfig,
                  path_cfg: PathConfig, device: torch.device):
    """Train and evaluate DL models."""
    results = []
    anomaly_store = {}

    for name in model_names:
        model = build_dl_model(name, data_mod.n_features, seq_cfg, model_cfg, device)

        if name == "LSTM_Autoencoder":
            # --- Autoencoder path ---
            train_loader, val_loader = data_mod.get_ae_loaders(
                train_cfg.batch_size, train_cfg.val_split
            )
            loss_data = train_model(
                model, train_loader, val_loader,
                train_cfg, path_cfg, name, device,
            )

            # Plot training loss
            plot_training_loss(
                loss_data["train_loss"], loss_data["val_loss"], name,
                path_cfg.loss_dir(name) / "training_loss.png",
            )

            # Reconstruction error on test
            X_test_ae = data_mod.get_test_ae_tensors()
            X_test_np = data_mod.X_test_ae
            recon = predict(model, X_test_ae)

            # Threshold from training data
            X_train_ae = data_mod.get_train_ae_tensors()
            train_recon = predict(model, X_train_ae)
            train_error = np.mean(np.abs(data_mod.X_train_ae - train_recon), axis=(1, 2))

            recon_error, threshold, anomalies = detect_anomalies_reconstruction(
                X_test_np, recon, anomaly_cfg.sigma,
                train_error.mean(), train_error.std(),
            )

            n_anom = anomalies.sum()
            print(f"  Anomalies: {n_anom} ({n_anom/len(anomalies)*100:.1f}%)")

            # Use test timestamps for AE (offset by input_steps only, no output_steps)
            ae_timestamps = data_mod.test_df.index[seq_cfg.input_steps:]
            n = min(len(ae_timestamps), len(recon_error))
            y_actual_ae = data_mod.test_df[data_mod.data_cfg.target].values[seq_cfg.input_steps:]

            pdir = path_cfg.plots_dir(name)
            plot_autoencoder_results(
                ae_timestamps[:n], y_actual_ae[:n], recon_error[:n], anomalies[:n],
                threshold, name, pdir / "anomaly_detection.png",
            )
            plot_error_distribution(recon_error, threshold, name,
                                    pdir / "reconstruction_error_dist.png")

            anomaly_store[name] = anomalies

        else:
            # --- Prediction model path ---
            train_loader, val_loader = data_mod.get_pred_loaders(
                train_cfg.batch_size, train_cfg.val_split
            )
            loss_data = train_model(
                model, train_loader, val_loader,
                train_cfg, path_cfg, name, device,
            )

            # Plot training loss
            plot_training_loss(
                loss_data["train_loss"], loss_data["val_loss"], name,
                path_cfg.loss_dir(name) / "training_loss.png",
            )

            # Predict on test
            X_test = data_mod.get_test_tensors()
            pred_scaled = predict(model, X_test)

            # Inverse-transform
            pred_actual = data_mod.inverse_predictions(pred_scaled)
            y_actual = data_mod.y_test_actual

            # For metrics: average across output steps
            result = evaluate_model(y_actual, pred_actual, name)
            results.append(result)

            # Anomaly detection
            mae, threshold, anomalies = detect_anomalies_3sigma(
                y_actual, pred_actual, anomaly_cfg.sigma
            )

            # Average across steps for 1D arrays
            y_act_1d = y_actual.mean(axis=1) if y_actual.ndim == 2 else y_actual
            pred_1d = pred_actual.mean(axis=1) if pred_actual.ndim == 2 else pred_actual

            n_anom = anomalies.sum()
            print(f"  Anomalies: {n_anom} ({n_anom/len(anomalies)*100:.1f}%)")

            # Plots
            ts = data_mod.test_timestamps
            pdir = path_cfg.plots_dir(name)
            plot_anomalies(ts, y_act_1d, pred_1d, anomalies,
                          name, threshold, pdir / "anomaly_detection.png")
            plot_prediction_vs_actual(ts, y_act_1d, pred_1d,
                                      name, pdir / "prediction_vs_actual.png")
            plot_error_distribution(mae, threshold, name,
                                    pdir / "error_distribution.png")

            anomaly_store[name] = anomalies

    return results, anomaly_store


def main():
    parser = argparse.ArgumentParser(
        description="Solar Power Anomaly Detection Training Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python run_training.py                                        # auto trial: run_1, run_2, ...
  python run_training.py --trial-name baseline                  # named trial
  python run_training.py --trial-name exp_lr --lr 0.0005        # experiment with name
  python run_training.py --input-steps 6 --output-steps 2       # 1.5h in, 30min out
  python run_training.py --models LSTM Transformer               # specific models only
        """,
    )
    parser.add_argument("--input-steps", type=int, default=4,
                        help="Number of input timesteps (default: 4 = 1 hour)")
    parser.add_argument("--output-steps", type=int, default=1,
                        help="Number of output timesteps to predict (default: 1 = 15min)")
    parser.add_argument("--models", nargs="+", default=["all"],
                        help="Models to train: all, LSTM, CNN_LSTM, LSTM_Autoencoder, "
                             "Transformer, Isolation_Forest, Random_Forest")
    parser.add_argument("--epochs", type=int, default=100)
    parser.add_argument("--batch-size", type=int, default=32)
    parser.add_argument("--lr", type=float, default=0.001)
    parser.add_argument("--patience", type=int, default=15)
    parser.add_argument("--sigma", type=float, default=3.0,
                        help="Sigma for anomaly threshold (default: 3.0)")
    parser.add_argument("--trial-name", type=str, default=None,
                        help="Trial name for output folder (default: auto run_1, run_2, ...)")
    parser.add_argument("--data-dir", type=str, default="datasets/site_1")
    parser.add_argument("--output-dir", type=str, default="outputs")
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    # --- Resolve trial name ---
    trial_name = resolve_trial_name(args.output_dir, args.trial_name)
    args.output_dir = str(Path(args.output_dir) / trial_name)

    # --- Setup logging ---
    tee_out, tee_err = setup_logging(args, trial_name)

    # --- Resolve models ---
    if "all" in args.models:
        ml_to_run = ML_MODELS[:]
        dl_to_run = DL_MODELS[:]
    else:
        ml_to_run = [m for m in args.models if m in ML_MODELS]
        dl_to_run = [m for m in args.models if m in DL_MODELS]
        unknown = [m for m in args.models if m not in ALL_MODELS]
        if unknown:
            print(f"Unknown models: {unknown}")
            print(f"Available: {ALL_MODELS}")
            sys.exit(1)

    # --- Configs ---
    data_cfg = DataConfig(data_dir=args.data_dir)
    seq_cfg = SequenceConfig(input_steps=args.input_steps, output_steps=args.output_steps)
    train_cfg = TrainConfig(
        epochs=args.epochs, batch_size=args.batch_size,
        learning_rate=args.lr, patience=args.patience, seed=args.seed,
    )
    model_cfg = ModelConfig()
    anomaly_cfg = AnomalyConfig(sigma=args.sigma)
    path_cfg = PathConfig(base_output=args.output_dir)

    # --- Device ---
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    torch.manual_seed(train_cfg.seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed(train_cfg.seed)

    print("=" * 60)
    print("  Solar Power Anomaly Detection Pipeline")
    print("=" * 60)
    print(f"  Trial name:   {trial_name}")
    print(f"  Device:       {device}")
    if torch.cuda.is_available():
        print(f"  GPU:          {torch.cuda.get_device_name(0)}")
    print(f"  Input steps:  {seq_cfg.input_steps} ({seq_cfg.input_steps * 15} min)")
    print(f"  Output steps: {seq_cfg.output_steps} ({seq_cfg.output_steps * 15} min)")
    print(f"  Models (ML):  {ml_to_run}")
    print(f"  Models (DL):  {dl_to_run}")
    print(f"  Epochs:       {train_cfg.epochs}")
    print(f"  Batch size:   {train_cfg.batch_size}")
    print(f"  Sigma:        {anomaly_cfg.sigma}")
    print(f"  Output dir:   {path_cfg.base_output}")
    print("=" * 60)

    # --- EDA & Data Cleaning ---
    print("\n[1/6] Running EDA & data cleaning...")
    run_eda(args.data_dir, args.output_dir)

    # --- Load Data ---
    print("\n[2/6] Loading and preparing data...")
    data_mod = SolarDataModule(data_cfg, seq_cfg, device)
    data_mod.setup()

    # Save scaler + config
    comp_dir = path_cfg.comparison_dir()
    joblib.dump(data_mod.scaler, comp_dir / "scaler.joblib")
    run_config = {
        "trial_name": trial_name,
        "input_steps": seq_cfg.input_steps,
        "output_steps": seq_cfg.output_steps,
        "features": data_cfg.features,
        "target": data_cfg.target,
        "train_ratio": data_cfg.train_ratio,
        "sigma": anomaly_cfg.sigma,
        "epochs": train_cfg.epochs,
        "batch_size": train_cfg.batch_size,
        "lr": train_cfg.learning_rate,
    }
    with open(comp_dir / "run_config.json", "w") as f:
        json.dump(run_config, f, indent=2)

    # Plot train/test split
    plot_train_test_split(
        data_mod.train_df, data_mod.test_df, data_cfg.target,
        comp_dir / "train_test_split.png",
    )

    # --- Train ML Models ---
    all_results = []
    all_anomalies = {}

    if ml_to_run:
        print("\n[3/6] Training ML models...")
        ml_results, ml_anomalies = run_ml_models(
            data_mod, model_cfg, anomaly_cfg, path_cfg
        )
        all_results.extend(ml_results)
        all_anomalies.update(ml_anomalies)

    # --- Train DL Models ---
    if dl_to_run:
        print("\n[4/6] Training DL models...")
        dl_results, dl_anomalies = run_dl_models(
            dl_to_run, data_mod, seq_cfg, train_cfg,
            model_cfg, anomaly_cfg, path_cfg, device,
        )
        all_results.extend(dl_results)
        all_anomalies.update(dl_anomalies)

    # --- Comparison ---
    print("\n[5/6] Generating comparison plots...")

    if all_results:
        results_df = build_results_table(all_results)
        print("\n=== Model Comparison ===")
        print(results_df)
        results_df.to_csv(comp_dir / "model_comparison.csv")
        plot_model_comparison(results_df, comp_dir / "model_comparison.png")

    anomaly_counts = {k: int(v.sum()) for k, v in all_anomalies.items()}
    if anomaly_counts:
        print(f"\n=== Anomaly Counts ===")
        for k, v in anomaly_counts.items():
            print(f"  {k:25s}: {v}")
        plot_anomaly_comparison(anomaly_counts, comp_dir / "anomaly_comparison.png")

        with open(comp_dir / "anomaly_counts.json", "w") as f:
            json.dump(anomaly_counts, f, indent=2)

    # --- Ensemble ---
    if len(all_anomalies) >= 2:
        print("\n[6/6] Ensemble anomaly detection...")

        votes, ensemble, distribution = ensemble_voting(
            all_anomalies, anomaly_cfg.ensemble_min_votes
        )
        print(f"  Ensemble anomalies (>={anomaly_cfg.ensemble_min_votes} votes): {ensemble.sum()}")
        print(f"  Vote distribution: {distribution}")

        # For ensemble plot, use the shortest aligned timestamps
        n = len(votes)
        # Use test timestamps based on whether we have DL models
        if dl_to_run:
            ts = data_mod.test_timestamps[:n]
            y_act = data_mod.y_test_actual[:n]
            y_act_1d = y_act.mean(axis=1) if y_act.ndim == 2 else y_act
        else:
            ts = data_mod.test_df.index[:n]
            y_act_1d = data_mod.test_df[data_cfg.target].values[:n]

        ens_dir = path_cfg.ensemble_dir()
        plot_ensemble(ts, y_act_1d, votes, ensemble,
                      anomaly_cfg.ensemble_min_votes, ens_dir / "ensemble_anomalies.png")

        with open(ens_dir / "ensemble_results.json", "w") as f:
            json.dump({
                "total_anomalies": int(ensemble.sum()),
                "min_votes": anomaly_cfg.ensemble_min_votes,
                "vote_distribution": distribution,
            }, f, indent=2)

    print("\n" + "=" * 60)
    print("  Pipeline complete!")
    print(f"  Results saved to: {path_cfg.base_output}/")
    print("=" * 60)

    # Close log file
    sys.stdout = tee_out.stream
    sys.stderr = tee_err.stream
    tee_out.close()
    tee_err.close()


if __name__ == "__main__":
    main()
