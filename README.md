# Solar Power Systems Anomaly Detection

Detect anomalies in solar power generation systems using Machine Learning and Deep Learning. The system learns the normal relationship between weather conditions (irradiance, temperature) and solar output, then flags significant deviations as anomalies indicating equipment faults, panel degradation, inverter failures, or sensor errors.

## Key Results (Site 1)

| Model | RMSE | MAE | R2 | Anomalies |
|-------|------|-----|-----|-----------|
| Random Forest | 83.37 | 34.80 | 0.6346 | 159 (3.5%) |
| LSTM | 32.46 | 15.43 | 0.9445 | 102 (2.3%) |
| CNN-LSTM | 32.87 | 15.45 | 0.9431 | 94 (2.1%) |
| **Transformer** | **32.05** | **13.76** | **0.9459** | 100 (2.2%) |
| Isolation Forest | - | - | - | 244 (5.4%) |
| LSTM Autoencoder | - | - | - | 265 (5.9%) |
| **Ensemble (>=3/6)** | - | - | - | **131 (2.9%)** |

## Project Structure

```
SolarPowerSystemsAnomalyDetection/
|
|-- run_training.py                 # Main CLI - train & evaluate all models
|-- requirements.txt                # Python dependencies
|
|-- src/                            # Modular source code
|   |-- config.py                   # All configurations (data, model, training, paths)
|   |-- data.py                     # Data loading, cleaning, variable-length sequences
|   |-- trainer.py                  # PyTorch training loop with early stopping
|   |-- evaluate.py                 # Metrics, 3-sigma anomaly detection, ensemble voting
|   |-- visualize.py                # All plotting functions (auto-save to output dirs)
|   |-- models/
|       |-- lstm.py                 # LSTM model
|       |-- cnn_lstm.py             # CNN-LSTM hybrid model
|       |-- autoencoder.py          # LSTM Autoencoder (reconstruction-based)
|       |-- transformer.py          # Transformer model
|       |-- ml_models.py            # Isolation Forest, Random Forest (scikit-learn)
|
|-- notebooks/                      # Jupyter notebooks (EDA & initial experiments)
|   |-- 01_EDA.ipynb                # Exploratory Data Analysis
|   |-- 02_training_anomaly_detection.ipynb   # Original training notebook
|
|-- datasets/
|   |-- site_1/                     # Raw CSV files (per site)
|   |   |-- gen_337_*.csv           # Solar generation (kW)
|   |   |-- Irradiance_339_*.csv    # Solar irradiance (W/m2)
|   |   |-- Temp-Ambient_340_*.csv  # Ambient temperature (C)
|   |   |-- load_338_*.csv          # Power consumption (kW)
|   |-- site_1_cleaned.csv          # Merged & cleaned data (auto-generated)
|
|-- outputs/                        # Auto-generated per training run
|   |-- {ModelName}/
|   |   |-- model/                  # Saved model weights (.pt / .joblib)
|   |   |-- loss/                   # Loss history (.json, .csv, training_loss.png)
|   |   |-- plots/                  # Anomaly detection, prediction, error distribution
|   |-- Ensemble/plots/             # Ensemble voting results
|   |-- Comparison/                 # Cross-model comparison (CSV, PNG, configs)
|
|-- Result.md                       # Detailed results summary
|-- Presentation_Summary.md         # Slide-by-slide presentation content
|-- Solar_Anomaly_Detection_Presentation.pptx      # PowerPoint (English)
|-- Solar_Anomaly_Detection_Presentation_TH.pptx   # PowerPoint (Thai)
|-- References/                     # Research papers & reference project
```

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

**Requirements:** Python 3.10+, PyTorch 2.0+, scikit-learn, pandas, matplotlib, seaborn

### 2. Run Training (Default)

```bash
python run_training.py
```

This trains all 6 models with default settings: input=4 timesteps (1 hour), output=1 timestep (15 min).

### 3. Custom Timesteps

The pipeline supports **variable input/output window sizes**:

```bash
# Input: 6 steps (1.5 hours) -> Output: 2 steps (30 min)
python run_training.py --input-steps 6 --output-steps 2

# Input: 8 steps (2 hours) -> Output: 3 steps (45 min)
python run_training.py --input-steps 8 --output-steps 3
```

### 4. Select Specific Models

```bash
# Only LSTM and Transformer
python run_training.py --models LSTM Transformer

# Only ML models
python run_training.py --models Isolation_Forest Random_Forest

# Single model
python run_training.py --models CNN_LSTM
```

**Available models:** `LSTM`, `CNN_LSTM`, `LSTM_Autoencoder`, `Transformer`, `Isolation_Forest`, `Random_Forest`

### 5. Tune Hyperparameters

```bash
python run_training.py \
    --epochs 200 \
    --batch-size 64 \
    --lr 0.0005 \
    --patience 20 \
    --sigma 2.5
```

### 6. Full CLI Reference

```
python run_training.py --help

options:
  --input-steps N     Input timesteps (default: 4 = 1 hour)
  --output-steps N    Output timesteps to predict (default: 1 = 15 min)
  --models [M ...]    Models to train (default: all)
  --epochs N          Training epochs (default: 100)
  --batch-size N      Batch size (default: 32)
  --lr F              Learning rate (default: 0.001)
  --patience N        Early stopping patience (default: 15)
  --sigma F           Anomaly threshold sigma (default: 3.0)
  --data-dir PATH     Data directory (default: datasets/site_1)
  --output-dir PATH   Output directory (default: outputs)
  --seed N            Random seed (default: 42)
```

## Methodology

### Approach

```
Weather Conditions -----> Model -----> Predicted Generation
(Irradiance, Temp)         |
                           v
                    Actual Generation
                           |
                    Prediction Error = |Actual - Predicted|
                           |
                    Error > Threshold? ---> ANOMALY
```

1. **Train** models to predict solar generation from weather features
2. **Anomaly** = when actual generation significantly deviates from expected output
3. **3-sigma threshold**: `threshold = mean(error) + 3 * std(error)`
4. **Ensemble voting**: flag as anomaly if >= 3 of 6 models agree

### Models

| Model | Type | Approach |
|-------|------|----------|
| Isolation Forest | ML, Unsupervised | Statistically isolated points in feature space |
| Random Forest | ML, Regression | Prediction error exceeds threshold |
| LSTM | DL, Sequential | Temporal patterns from sliding window |
| CNN-LSTM | DL, Hybrid | CNN local features + LSTM temporal dynamics |
| LSTM Autoencoder | DL, Reconstruction | Normal pattern compression; high reconstruction error = anomaly |
| Transformer | DL, Self-Attention | Multi-head attention across all timesteps |

### Variable Timestep Design

All models support configurable input/output windows:

```
Example: --input-steps 6 --output-steps 2

Time: [t-5] [t-4] [t-3] [t-2] [t-1] [t]  |  [t+1] [t+2]
      |_____________ INPUT _____________|  |  |__ OUTPUT __|
              6 x 15min = 1.5h                2 x 15min = 30min
```

- **Prediction models** (LSTM, CNN-LSTM, Transformer): output shape = `(output_steps * n_features)`
- **Autoencoder**: always reconstructs the input window (independent of output_steps)

## Output Structure

After training, the `outputs/` directory contains:

```
outputs/
|-- LSTM/
|   |-- model/LSTM.pt                    # Saved PyTorch weights
|   |-- loss/
|   |   |-- loss_history.json            # Train & val loss per epoch
|   |   |-- loss_history.csv             # Same in CSV format
|   |   |-- training_loss.png            # Loss curve plot
|   |-- plots/
|       |-- anomaly_detection.png        # 3-panel: pred vs actual, MAE, anomalies
|       |-- prediction_vs_actual.png     # Predicted vs actual generation
|       |-- error_distribution.png       # Error histogram with threshold
|
|-- CNN_LSTM/model/ loss/ plots/         # Same structure
|-- LSTM_Autoencoder/model/ loss/ plots/ # Reconstruction error plots
|-- Transformer/model/ loss/ plots/
|-- Isolation_Forest/model/ plots/       # No loss (not DL)
|-- Random_Forest/model/ plots/
|
|-- Ensemble/plots/
|   |-- ensemble_anomalies.png           # Majority vote results
|   |-- ensemble_results.json            # Vote distribution
|
|-- Comparison/
    |-- model_comparison.csv             # RMSE, MAE, R2 per model
    |-- model_comparison.png             # Bar chart comparison
    |-- anomaly_comparison.png           # Anomaly count per model
    |-- anomaly_counts.json
    |-- run_config.json                  # Hyperparameters used
    |-- scaler.joblib                    # Fitted scaler (for inference)
    |-- train_test_split.png
```

## Data

### Input Format (Per Site)

Each site has 4 CSV files with columns `date/time` and a value column:

| File | Variable | Unit | Resolution |
|------|----------|------|------------|
| `gen_*` | Solar generation | kW | ~1 min |
| `Irradiance_*` | Solar irradiance | W/m2 | ~1 min |
| `Temp-Ambient_*` | Ambient temperature | C | ~1 min |
| `load_*` | Power consumption | kW | 15 min |

### Preprocessing Pipeline

1. **Resample** all to uniform 15-min intervals (mean aggregation)
2. **Merge** via outer join on timestamp
3. **Forward-fill** gaps up to 1 hour
4. **Interpolate** remaining short gaps
5. **Drop** rows with unrecoverable gaps

### Adding New Sites

Place CSV files in `datasets/site_N/` and run:

```bash
python run_training.py --data-dir datasets/site_N --output-dir outputs_site_N
```

## Using Saved Models for Inference

```python
import torch
import joblib
from src.config import SequenceConfig
from src.data import create_sequences
from src.models import LSTMModel

# Load config and scaler
scaler = joblib.load("outputs/Comparison/scaler.joblib")

# Load model
model = LSTMModel(n_features=3, output_steps=1)
model.load_state_dict(torch.load("outputs/LSTM/model/LSTM.pt"))
model.eval()

# Prepare new data
new_data_scaled = scaler.transform(new_df[["irradiance_wm2", "temperature_c", "generation_kw"]])
X, _ = create_sequences(new_data_scaled, input_steps=4, output_steps=1)
X_tensor = torch.FloatTensor(X)

# Predict
with torch.no_grad():
    predictions = model(X_tensor).numpy()
```

## References

- Research papers in `References/` directory
- Reference implementation: `References/Data_Mining_Project/`
  - Models: LSTM, MLP+LSTM Encoder, CNN-LSTM
  - Dataset: Solar power generation plant data (Kaggle)
  - Anomaly detection via 3-sigma rule on prediction error
