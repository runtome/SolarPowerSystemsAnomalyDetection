from .lstm import LSTMModel
from .cnn_lstm import CNNLSTM
from .autoencoder import LSTMAutoencoder
from .transformer import TransformerModel

__all__ = ["LSTMModel", "CNNLSTM", "LSTMAutoencoder", "TransformerModel"]

# Optional: TimesFM foundation model (requires `pip install timesfm`)
try:
    from .timesfm_wrapper import is_timesfm_available, load_timesfm_model, forecast_test_period
    __all__ += ["is_timesfm_available", "load_timesfm_model", "forecast_test_period"]
except ImportError:
    pass
