"""LSTM Autoencoder for reconstruction-based anomaly detection."""

import torch
import torch.nn as nn


class LSTMAutoencoder(nn.Module):
    """Reconstructs input sequences. High reconstruction error = anomaly.

    Input/Output shape: (batch, input_steps, n_features) -> (batch, input_steps, n_features)
    This model does NOT use output_steps; it reconstructs the input window.
    """

    def __init__(self, n_features: int, input_steps: int,
                 hidden_enc: int = 32, bottleneck: int = 16, dropout: float = 0.2,
                 **kwargs):
        super().__init__()
        self.input_steps = input_steps

        # Encoder
        self.enc_lstm1 = nn.LSTM(n_features, hidden_enc, batch_first=True)
        self.enc_drop = nn.Dropout(dropout)
        self.enc_lstm2 = nn.LSTM(hidden_enc, bottleneck, batch_first=True)

        # Decoder
        self.dec_lstm1 = nn.LSTM(bottleneck, bottleneck, batch_first=True)
        self.dec_drop = nn.Dropout(dropout)
        self.dec_lstm2 = nn.LSTM(bottleneck, hidden_enc, batch_first=True)
        self.output_layer = nn.Linear(hidden_enc, n_features)

    def forward(self, x):
        # Encode
        enc, _ = self.enc_lstm1(x)
        enc = self.enc_drop(enc)
        enc, _ = self.enc_lstm2(enc)
        bottleneck = enc[:, -1, :]  # (batch, bottleneck)

        # Decode: repeat bottleneck for each timestep
        dec_input = bottleneck.unsqueeze(1).repeat(1, self.input_steps, 1)
        dec, _ = self.dec_lstm1(dec_input)
        dec = self.dec_drop(dec)
        dec, _ = self.dec_lstm2(dec)
        out = self.output_layer(dec)  # (batch, input_steps, n_features)
        return out
