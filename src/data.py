"""Data loading, preprocessing, and sequence creation."""

import numpy as np
import pandas as pd
import torch
from torch.utils.data import DataLoader, TensorDataset
from sklearn.preprocessing import MinMaxScaler
from pathlib import Path

from .config import DataConfig, SequenceConfig


def load_raw_data(cfg: DataConfig) -> dict:
    """Load raw CSV files from site directory."""
    data_dir = Path(cfg.data_dir)
    csvs = sorted(data_dir.glob("*.csv"))
    raw = {}
    for f in csvs:
        df = pd.read_csv(f, parse_dates=["date/time"])
        col_name = df.columns[1]
        df.columns = ["datetime", col_name]
        raw[col_name] = df
    return raw


def preprocess_and_merge(cfg: DataConfig, force_reload: bool = False) -> pd.DataFrame:
    """Load cleaned CSV or build from raw data.

    Returns DataFrame indexed by datetime with columns:
    [feature_1, feature_2, ..., target]
    """
    cleaned_path = Path(cfg.cleaned_csv)
    if cleaned_path.exists() and not force_reload:
        df = pd.read_csv(cleaned_path, index_col="datetime", parse_dates=True)
        all_cols = [c for c in cfg.features + [cfg.target] if c in df.columns]
        return df[all_cols]

    # Build from raw
    raw = load_raw_data(cfg)
    dfs = {}
    name_map = {}
    for col_name, df in raw.items():
        clean_name = col_name.lower().replace(" ", "_")
        if "gen" in clean_name:
            clean_name = "generation_kw"
        elif "irradiance" in clean_name:
            clean_name = "irradiance_wm2"
        elif "temp" in clean_name:
            clean_name = "temperature_c"
        elif "load" in clean_name:
            clean_name = "load_kw"
        name_map[col_name] = clean_name
        df = df.rename(columns={col_name: clean_name})
        df_rs = df.set_index("datetime").resample(cfg.resample_interval).mean()
        dfs[clean_name] = df_rs

    merged = pd.concat(dfs.values(), axis=1, join="outer").sort_index()
    merged = merged.ffill(limit=cfg.ffill_limit)
    merged = merged.interpolate(method="time", limit=cfg.interp_limit)
    merged = merged.dropna()

    merged.to_csv(cleaned_path)
    all_cols = cfg.features + [cfg.target]
    return merged[all_cols]


def create_sequences(data: np.ndarray, input_steps: int, output_steps: int):
    """Create variable-length input/output sequences.

    Args:
        data: (total_timesteps, n_features)
        input_steps: number of past timesteps as input
        output_steps: number of future timesteps to predict

    Returns:
        X: (samples, input_steps, n_features)
        Y: (samples, output_steps, n_features)
    """
    X, Y = [], []
    for i in range(input_steps, len(data) - output_steps + 1):
        X.append(data[i - input_steps:i, :])
        Y.append(data[i:i + output_steps, :])
    return np.array(X), np.array(Y)


def create_ae_sequences(data: np.ndarray, input_steps: int):
    """Create sequences for autoencoder (input = target).

    Returns:
        X: (samples, input_steps, n_features)
    """
    X = []
    for i in range(input_steps, len(data) + 1):
        X.append(data[i - input_steps:i, :])
    return np.array(X)


class SolarDataModule:
    """Manages all data operations: load, scale, split, create dataloaders."""

    def __init__(self, data_cfg: DataConfig, seq_cfg: SequenceConfig, device: torch.device):
        self.data_cfg = data_cfg
        self.seq_cfg = seq_cfg
        self.device = device
        self.scaler = MinMaxScaler()
        self.df = None
        self.train_df = None
        self.test_df = None
        self.n_features = None
        self.target_idx = None

    def setup(self):
        """Load data, split, scale, create sequences."""
        all_cols = self.data_cfg.features + [self.data_cfg.target]
        self.df = preprocess_and_merge(self.data_cfg)
        self.n_features = len(all_cols)
        self.target_idx = all_cols.index(self.data_cfg.target)

        # Time-based split
        split_idx = int(len(self.df) * self.data_cfg.train_ratio)
        self.train_df = self.df.iloc[:split_idx]
        self.test_df = self.df.iloc[split_idx:]

        # Scale
        self.train_scaled = self.scaler.fit_transform(self.train_df)
        self.test_scaled = self.scaler.transform(self.test_df)

        # Sequences for prediction models
        inp, out = self.seq_cfg.input_steps, self.seq_cfg.output_steps
        self.X_train, self.Y_train = create_sequences(self.train_scaled, inp, out)
        self.X_test, self.Y_test = create_sequences(self.test_scaled, inp, out)

        # Sequences for autoencoder
        self.X_train_ae = create_ae_sequences(self.train_scaled, inp)
        self.X_test_ae = create_ae_sequences(self.test_scaled, inp)

        # Test timestamps (aligned to prediction sequences)
        self.test_timestamps = self.test_df.index[inp + out - 1:]

        # Actual test target values (inverse-scaled)
        self.y_test_actual = self._inverse_target(self.Y_test)

        print(f"  Data loaded: {len(self.df)} rows")
        print(f"  Train: {self.train_df.shape[0]} | Test: {self.test_df.shape[0]}")
        print(f"  Sequence config: input={inp} -> output={out}")
        print(f"  X_train: {self.X_train.shape} | Y_train: {self.Y_train.shape}")
        print(f"  X_test:  {self.X_test.shape}  | Y_test:  {self.Y_test.shape}")
        print(f"  X_train_ae: {self.X_train_ae.shape} | X_test_ae: {self.X_test_ae.shape}")

    def get_pred_loaders(self, batch_size: int, val_split: float):
        """Get train/val DataLoaders for prediction models."""
        X_t = torch.FloatTensor(self.X_train).to(self.device)
        # Flatten output steps: (samples, output_steps * n_features)
        Y_flat = self.Y_train.reshape(self.Y_train.shape[0], -1)
        Y_t = torch.FloatTensor(Y_flat).to(self.device)

        n_val = int(len(X_t) * val_split)
        n_train = len(X_t) - n_val

        train_ds = TensorDataset(X_t[:n_train], Y_t[:n_train])
        val_ds = TensorDataset(X_t[n_train:], Y_t[n_train:])

        train_loader = DataLoader(train_ds, batch_size=batch_size, shuffle=True)
        val_loader = DataLoader(val_ds, batch_size=batch_size)
        return train_loader, val_loader

    def get_ae_loaders(self, batch_size: int, val_split: float):
        """Get train/val DataLoaders for autoencoder (input=target)."""
        X_t = torch.FloatTensor(self.X_train_ae).to(self.device)

        n_val = int(len(X_t) * val_split)
        n_train = len(X_t) - n_val

        train_ds = TensorDataset(X_t[:n_train], X_t[:n_train])
        val_ds = TensorDataset(X_t[n_train:], X_t[n_train:])

        train_loader = DataLoader(train_ds, batch_size=batch_size, shuffle=True)
        val_loader = DataLoader(val_ds, batch_size=batch_size)
        return train_loader, val_loader

    def get_test_tensors(self):
        """Get test tensors for prediction models."""
        X = torch.FloatTensor(self.X_test).to(self.device)
        return X

    def get_test_ae_tensors(self):
        """Get test tensors for autoencoder."""
        X = torch.FloatTensor(self.X_test_ae).to(self.device)
        return X

    def get_train_ae_tensors(self):
        """Get train tensors for autoencoder (for threshold computation)."""
        X = torch.FloatTensor(self.X_train_ae).to(self.device)
        return X

    def _inverse_target(self, Y_seq: np.ndarray) -> np.ndarray:
        """Inverse-scale and extract target column from multi-step predictions.

        Args:
            Y_seq: (samples, output_steps, n_features) or (samples, output_steps * n_features)

        Returns:
            (samples, output_steps) array of target values in original scale
        """
        if Y_seq.ndim == 2:
            # Flattened: reshape to (samples, output_steps, n_features)
            Y_seq = Y_seq.reshape(Y_seq.shape[0], -1, self.n_features)

        samples, out_steps, n_feat = Y_seq.shape
        result = np.zeros((samples, out_steps))
        for t in range(out_steps):
            inversed = self.scaler.inverse_transform(Y_seq[:, t, :])
            result[:, t] = inversed[:, self.target_idx]
        return result

    def inverse_predictions(self, preds: np.ndarray) -> np.ndarray:
        """Inverse-scale predictions. Public wrapper."""
        return self._inverse_target(preds)
