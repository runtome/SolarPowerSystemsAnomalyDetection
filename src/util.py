"""Print torch model summaries for all DL models with default config."""

import sys
import io

import torch
from torchinfo import summary

from src.config import SequenceConfig, ModelConfig
from src.models import LSTMModel, CNNLSTM, LSTMAutoencoder, TransformerModel


def print_all_summaries():
    seq_cfg = SequenceConfig()       # input_steps=4, output_steps=1
    model_cfg = ModelConfig()
    n_features = 3                   # irradiance, temperature, generation

    inp = seq_cfg.input_steps
    out = seq_cfg.output_steps
    batch = 32

    models = {
        "LSTM": LSTMModel(
            n_features=n_features, output_steps=out,
            hidden_size=model_cfg.lstm_hidden,
            num_layers=model_cfg.lstm_layers,
            dropout=model_cfg.lstm_dropout,
        ),
        "CNN_LSTM": CNNLSTM(
            n_features=n_features, output_steps=out,
            cnn_filters=model_cfg.cnn_filters,
            lstm_hidden=model_cfg.cnn_lstm_hidden,
        ),
        "LSTM_Autoencoder": LSTMAutoencoder(
            n_features=n_features, input_steps=inp,
            hidden_enc=model_cfg.ae_hidden_enc,
            bottleneck=model_cfg.ae_bottleneck,
            dropout=model_cfg.ae_dropout,
        ),
        "Transformer": TransformerModel(
            n_features=n_features, output_steps=out,
            d_model=model_cfg.tf_d_model,
            nhead=model_cfg.tf_nhead,
            num_layers=model_cfg.tf_num_layers,
            dim_feedforward=model_cfg.tf_ff_dim,
            dropout=model_cfg.tf_dropout,
        ),
    }

    print("=" * 70)
    print(f"  Default config: n_features={n_features}, "
          f"input_steps={inp}, output_steps={out}, batch={batch}")
    print("=" * 70)

    for name, model in models.items():
        print(f"\n{'='*70}")
        print(f"  {name}")
        print(f"{'='*70}")

        if name == "LSTM_Autoencoder":
            input_size = (batch, inp, n_features)
        else:
            input_size = (batch, inp, n_features)

        result = summary(model, input_size=input_size, col_names=[
            "input_size", "output_size", "num_params", "kernel_size",
        ], verbose=0)
        # Handle Windows console encoding by replacing unicode box chars
        text = str(result)
        text = text.replace("\u2500", "-").replace("\u2550", "=")
        text = text.replace("\u2502", "|").replace("\u2551", "|")
        text = text.replace("\u250c", "+").replace("\u2510", "+")
        text = text.replace("\u2514", "+").replace("\u2518", "+")
        text = text.replace("\u251c", "+").replace("\u2524", "+")
        text = text.replace("\u252c", "+").replace("\u2534", "+")
        text = text.replace("\u253c", "+")
        text = text.replace("\u2554", "+").replace("\u2557", "+")
        text = text.replace("\u255a", "+").replace("\u255d", "+")
        text = text.replace("\u2560", "+").replace("\u2563", "+")
        text = text.replace("\u2566", "+").replace("\u2569", "+")
        text = text.replace("\u256c", "+")
        print(text)


if __name__ == "__main__":
    print_all_summaries()
