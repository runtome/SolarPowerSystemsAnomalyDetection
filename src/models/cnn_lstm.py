"""CNN-LSTM hybrid model for multi-step solar generation prediction."""

import torch.nn as nn


class CNNLSTM(nn.Module):
    def __init__(self, n_features: int, output_steps: int,
                 cnn_filters: int = 32, lstm_hidden: int = 32, dropout: float = 0.2):
        super().__init__()
        self.output_steps = output_steps
        self.n_features = n_features

        self.cnn = nn.Sequential(
            nn.Conv1d(n_features, cnn_filters, kernel_size=2, padding=1),
            nn.ReLU(),
            nn.Conv1d(cnn_filters, 16, kernel_size=2, padding=1),
            nn.ReLU(),
        )
        self.lstm = nn.LSTM(
            input_size=16,
            hidden_size=lstm_hidden,
            num_layers=1,
            batch_first=True,
        )
        self.fc = nn.Sequential(
            nn.Linear(lstm_hidden, 64),
            nn.ReLU(),
            nn.Dropout(dropout),
            nn.Linear(64, output_steps * n_features),
        )

    def forward(self, x):
        # x: (batch, input_steps, n_features)
        x_cnn = x.permute(0, 2, 1)         # (batch, n_features, input_steps)
        cnn_out = self.cnn(x_cnn)           # (batch, 16, seq)
        cnn_out = cnn_out.permute(0, 2, 1)  # (batch, seq, 16)
        lstm_out, _ = self.lstm(cnn_out)
        out = self.fc(lstm_out[:, -1, :])   # (batch, output_steps * n_features)
        return out
