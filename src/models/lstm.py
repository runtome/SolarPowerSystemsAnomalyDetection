"""LSTM model for multi-step solar generation prediction."""

import torch.nn as nn


class LSTMModel(nn.Module):
    def __init__(self, n_features: int, output_steps: int,
                 hidden_size: int = 64, num_layers: int = 2, dropout: float = 0.2):
        super().__init__()
        self.output_steps = output_steps
        self.n_features = n_features

        self.lstm = nn.LSTM(
            input_size=n_features,
            hidden_size=hidden_size,
            num_layers=num_layers,
            batch_first=True,
            dropout=dropout if num_layers > 1 else 0,
        )
        self.fc = nn.Sequential(
            nn.Linear(hidden_size, 64),
            nn.ReLU(),
            nn.Dropout(dropout),
            nn.Linear(64, output_steps * n_features),
        )

    def forward(self, x):
        # x: (batch, input_steps, n_features)
        out, _ = self.lstm(x)
        out = self.fc(out[:, -1, :])  # last timestep -> (batch, output_steps * n_features)
        return out
