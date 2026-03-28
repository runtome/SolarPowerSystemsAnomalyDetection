"""Transformer model for multi-step solar generation prediction."""

import numpy as np
import torch
import torch.nn as nn


class PositionalEncoding(nn.Module):
    def __init__(self, d_model: int, max_len: int = 500):
        super().__init__()
        pe = torch.zeros(max_len, d_model)
        position = torch.arange(0, max_len, dtype=torch.float).unsqueeze(1)
        div_term = torch.exp(torch.arange(0, d_model, 2).float() * (-np.log(10000.0) / d_model))
        pe[:, 0::2] = torch.sin(position * div_term)
        if d_model > 1:
            pe[:, 1::2] = torch.cos(position * div_term[:d_model // 2])
        self.register_buffer("pe", pe.unsqueeze(0))

    def forward(self, x):
        return x + self.pe[:, :x.size(1), :]


class TransformerModel(nn.Module):
    def __init__(self, n_features: int, output_steps: int,
                 d_model: int = 32, nhead: int = 2, num_layers: int = 2,
                 dim_feedforward: int = 64, dropout: float = 0.1):
        super().__init__()
        self.output_steps = output_steps
        self.n_features = n_features

        self.input_proj = nn.Linear(n_features, d_model)
        self.pos_enc = PositionalEncoding(d_model)

        encoder_layer = nn.TransformerEncoderLayer(
            d_model=d_model,
            nhead=nhead,
            dim_feedforward=dim_feedforward,
            dropout=dropout,
            batch_first=True,
        )
        self.transformer_enc = nn.TransformerEncoder(encoder_layer, num_layers=num_layers)

        self.fc = nn.Sequential(
            nn.Linear(d_model, 64),
            nn.ReLU(),
            nn.Dropout(dropout),
            nn.Linear(64, output_steps * n_features),
        )

    def forward(self, x):
        # x: (batch, input_steps, n_features)
        x = self.input_proj(x)
        x = self.pos_enc(x)
        x = self.transformer_enc(x)
        x = x.mean(dim=1)                  # global avg pool
        out = self.fc(x)                    # (batch, output_steps * n_features)
        return out
