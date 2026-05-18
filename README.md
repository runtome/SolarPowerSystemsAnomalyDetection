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
|   |-- eda.py                      # EDA, cleaning, anomaly indicators, 12 EDA plots
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
|   |-- 01_EDA.ipynb
|   |-- 02_training_anomaly_detection.ipynb
|
|-- datasets/                       # One subfolder per site
|   |-- site_1/
|   |   |-- gen_337_*.csv           # Solar generation (kW)
|   |   |-- Irradiance_339_*.csv    # Solar irradiance (W/m2)
|   |   |-- Temp-Ambient_340_*.csv  # Ambient temperature (C)
|   |   |-- load_338_*.csv          # Power consumption (kW)
|   |-- S1/                         # Sites with multiple generators
|   |   |-- gen 1_397_*.csv         # Generator 1 (summed with others automatically)
|   |   |-- gen 2_430_*.csv         # Generator 2
|   |   |-- ...
|   |-- site_1_cleaned.csv          # Merged & cleaned data (auto-generated per site)
|
|-- outputs/                        # Auto-generated per training run
|   |-- {trial_name}/
|       |-- {ModelName}/{site_name}/
|       |   |-- model/              # Saved weights (.pt / .joblib)
|       |   |-- loss/               # Loss history (.json, .csv, .png)
|       |   |-- plots/              # Anomaly, prediction, error plots
|       |-- Ensemble/{site_name}/plots/
|       |-- Comparison/{site_name}/
|       |-- data_EDA/{site_name}/   # 12 EDA plots + cleaning summary
|
|-- documents/                      # Technical documentation
|-- References/                     # Research papers & reference project
```

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

**Requirements:** Python 3.10+, PyTorch 2.0+, scikit-learn, pandas, matplotlib, seaborn

### 2. Run Training — All Sites (Default)

```bash
python run_training.py
```

Scans all subfolders in `datasets/` and trains all 6 models for each site in sequence. Each site's results are saved to `outputs/{trial}/{model}/{site_name}/`.

### 3. Run Training — Single Site

```bash
python run_training.py --data-dir datasets/site_1
```

### 4. Custom Timesteps

The pipeline supports **variable input/output window sizes**:

```bash
# Input: 6 steps (1.5 hours) -> Output: 2 steps (30 min)
python run_training.py --input-steps 6 --output-steps 2

# Input: 8 steps (2 hours) -> Output: 3 steps (45 min)
python run_training.py --input-steps 8 --output-steps 3
```

### 5. Select Specific Models

```bash
python run_training.py --models LSTM Transformer
python run_training.py --models Isolation_Forest Random_Forest
python run_training.py --models CNN_LSTM
```

**Available models:** `LSTM`, `CNN_LSTM`, `LSTM_Autoencoder`, `Transformer`, `Isolation_Forest`, `Random_Forest`

### 6. Tune Hyperparameters

```bash
python run_training.py \
    --epochs 200 \
    --batch-size 64 \
    --lr 0.0005 \
    --patience 20 \
    --sigma 2.5
```

### 7. Full CLI Reference

```
options:
  --input-steps N       Input timesteps (default: 4 = 1 hour)
  --output-steps N      Output timesteps to predict (default: 1 = 15 min)
  --models [M ...]      Models to train (default: all)
  --epochs N            Training epochs (default: 100)
  --batch-size N        Batch size (default: 32)
  --lr F                Learning rate (default: 0.001)
  --patience N          Early stopping patience (default: 15)
  --sigma F             Anomaly threshold sigma (default: 3.0)
  --trial-name NAME     Output folder name (default: auto run_1, run_2, ...)
  --datasets-dir PATH   Root folder of site subfolders (default: datasets/)
  --data-dir PATH       Single site folder — overrides --datasets-dir
  --output-dir PATH     Base output directory (default: outputs)
  --seed N              Random seed (default: 42)
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

After training, the `outputs/` directory contains one subfolder per trial, then per model, then per site:

```
outputs/{trial_name}/
|-- LSTM/{site_name}/
|   |-- model/LSTM.pt                    # Saved PyTorch weights
|   |-- loss/
|   |   |-- loss_history.json
|   |   |-- loss_history.csv
|   |   |-- training_loss.png
|   |-- plots/
|       |-- anomaly_detection.png
|       |-- prediction_vs_actual.png
|       |-- error_distribution.png
|
|-- CNN_LSTM/{site_name}/  LSTM_Autoencoder/{site_name}/  ...  (same structure)
|
|-- Ensemble/{site_name}/plots/
|   |-- ensemble_anomalies.png
|   |-- ensemble_results.json
|
|-- Comparison/{site_name}/
|   |-- model_comparison.csv             # RMSE, MAE, R2 per model
|   |-- model_comparison.png
|   |-- anomaly_comparison.png
|   |-- anomaly_counts.json
|   |-- run_config.json                  # Hyperparameters + site_name
|   |-- scaler.joblib                    # Fitted scaler (for inference)
|   |-- train_test_split.png
|
|-- data_EDA/{site_name}/
    |-- time_series.png
    |-- missing_data_pattern.png
    |-- before_after_cleaning.png
    |-- distributions.png  boxplots.png  correlation_matrix.png
    |-- hourly_patterns.png  monthly_trends.png
    |-- scatter_vs_generation.png
    |-- efficiency_analysis.png
    |-- anomaly_indicators.png
    |-- net_power.png
```

## Data

### Input Format (Per Site)

Each site folder contains CSV files with columns `date/time` and a value column. Files are matched by **case-insensitive keyword** in the filename:

| Keyword | Variable | Unit | Example filenames |
|---------|----------|------|-------------------|
| `gen` | Solar generation | kW | `gen_337_*.csv`, `gen 1_397_*.csv` |
| `irradiance` | Solar irradiance | W/m2 | `Irradiance_339_*.csv`, `Actual Irradiance_*.csv` |
| `temp` | Ambient temperature | C | `Temp-Ambient_340_*.csv`, `Actual Temp-Ambient_*.csv` |
| `load` | Power consumption | kW | `load_338_*.csv`, `load1_*.csv` |

**Multiple generation files** (e.g. `gen 1_*.csv`, `gen 2_*.csv`, `gen 3_*.csv`) are automatically detected and **summed** at each timestamp into a single `generation_kw` column.

### Preprocessing Pipeline

1. **Resample** all to uniform 15-min intervals (mean aggregation)
2. **Merge** via outer join on timestamp
3. **Forward-fill** gaps up to 1 hour
4. **Interpolate** remaining short gaps
5. **Drop** rows with unrecoverable gaps
6. **Zero** generation readings where irradiance ≤ 0 (physical constraint)
7. **Save** cleaned CSV to `datasets/{site_name}_cleaned.csv`

### Adding New Sites

Place a new folder under `datasets/` with the required CSV files:

```
datasets/
|-- my_new_site/
    |-- gen_*.csv
    |-- Irradiance_*.csv  (or Actual Irradiance_*.csv)
    |-- Temp-Ambient_*.csv  (or Actual Temp-Ambient_*.csv)
    |-- load_*.csv
```

Then run `python run_training.py` — the new site is picked up automatically.

To train only the new site:

```bash
python run_training.py --data-dir datasets/my_new_site
```

## Using Saved Models for Inference

```python
import torch
import joblib
from src.data import create_sequences
from src.models import LSTMModel

trial = "run_1"
site  = "site_1"

# Load scaler saved during training
scaler = joblib.load(f"outputs/{trial}/Comparison/{site}/scaler.joblib")

# Load model
model = LSTMModel(n_features=3, output_steps=1)
model.load_state_dict(torch.load(f"outputs/{trial}/LSTM/{site}/model/LSTM.pt"))
model.eval()

# Prepare new data
new_data_scaled = scaler.transform(new_df[["irradiance_wm2", "temperature_c", "generation_kw"]])
X, _ = create_sequences(new_data_scaled, input_steps=4, output_steps=1)
X_tensor = torch.FloatTensor(X)

# Predict
with torch.no_grad():
    predictions = model(X_tensor).numpy()
    predictions_actual = scaler.inverse_transform(predictions)
```

## References

- Research papers in `References/` directory
- Reference implementation: `References/Data_Mining_Project/`
  - Models: LSTM, MLP+LSTM Encoder, CNN-LSTM
  - Dataset: Solar power generation plant data (Kaggle)
  - Anomaly detection via 3-sigma rule on prediction error
