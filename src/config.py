"""Centralized configuration for Solar Anomaly Detection pipeline."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


@dataclass
class DataConfig:
    """Data loading and preprocessing settings."""
    data_dir: str = "datasets/site_1"
    cleaned_csv: str = ""  # auto-derived from data_dir if empty
    features: List[str] = field(default_factory=lambda: ["irradiance_wm2", "temperature_c"])
    target: str = "generation_kw"
    train_ratio: float = 0.7
    resample_interval: str = "15min"
    ffill_limit: int = 4       # forward-fill up to 4 intervals (1 hour)
    interp_limit: int = 4      # interpolation up to 4 intervals

    def __post_init__(self):
        if not self.cleaned_csv:
            p = Path(self.data_dir)
            self.cleaned_csv = str(p.parent / f"{p.name}_cleaned.csv")


@dataclass
class SequenceConfig:
    """Variable-length input/output timestep settings.

    Examples:
        input_steps=4,  output_steps=1  -> classic: 1h in, predict next 15min
        input_steps=6,  output_steps=2  -> 1.5h in, predict next 30min
        input_steps=8,  output_steps=3  -> 2h in, predict next 45min
    """
    input_steps: int = 4
    output_steps: int = 1


@dataclass
class TrainConfig:
    """Training hyperparameters."""
    epochs: int = 100
    batch_size: int = 32
    learning_rate: float = 0.001
    patience: int = 15
    lr_factor: float = 0.5
    lr_patience: int = 5
    val_split: float = 0.15
    grad_clip: float = 1.0
    seed: int = 42


@dataclass
class ModelConfig:
    """Model architecture hyperparameters."""
    # LSTM
    lstm_hidden: int = 64
    lstm_layers: int = 2
    lstm_dropout: float = 0.2

    # CNN-LSTM
    cnn_filters: int = 32
    cnn_lstm_hidden: int = 32

    # Autoencoder
    ae_hidden_enc: int = 32
    ae_bottleneck: int = 16
    ae_dropout: float = 0.2

    # Transformer
    tf_d_model: int = 32
    tf_nhead: int = 2
    tf_num_layers: int = 2
    tf_ff_dim: int = 64
    tf_dropout: float = 0.1

    # ML
    iso_n_estimators: int = 200
    iso_contamination: float = 0.02
    rf_n_estimators: int = 200
    rf_max_depth: int = 15


@dataclass
class AnomalyConfig:
    """Anomaly detection settings."""
    sigma: float = 3.0
    ensemble_min_votes: int = 3


@dataclass
class PathConfig:
    """Output directory structure."""
    base_output: str = "outputs"

    def model_dir(self, model_name: str) -> Path:
        p = Path(self.base_output) / model_name / "model"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def loss_dir(self, model_name: str) -> Path:
        p = Path(self.base_output) / model_name / "loss"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def plots_dir(self, model_name: str) -> Path:
        p = Path(self.base_output) / model_name / "plots"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def ensemble_dir(self) -> Path:
        p = Path(self.base_output) / "Ensemble" / "plots"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def comparison_dir(self) -> Path:
        p = Path(self.base_output) / "Comparison"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def eda_dir(self, site_name: str) -> Path:
        p = Path(self.base_output) / "data_EDA" / site_name
        p.mkdir(parents=True, exist_ok=True)
        return p


# --- All available model names ---
DL_MODELS = ["LSTM", "CNN_LSTM", "LSTM_Autoencoder", "Transformer"]
ML_MODELS = ["Isolation_Forest", "Random_Forest"]
ALL_MODELS = ML_MODELS + DL_MODELS
