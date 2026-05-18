# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Solar Power Systems Anomaly Detection is a machine learning pipeline that detects anomalies in solar power generation systems by learning the relationship between weather conditions (irradiance, temperature) and solar output, then flagging deviations as anomalies (equipment faults, panel degradation, inverter failures, sensor errors).

The system trains 6 models (4 deep learning + 2 traditional ML) and uses ensemble voting to identify anomalies with high confidence.

## Quick Start Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run full training pipeline with all 6 models (default: input=4 steps, output=1 step)
python run_training.py

# Custom experiment with trial name and different timestep windows
python run_training.py --trial-name exp_6step --input-steps 6 --output-steps 2

# Train specific models only
python run_training.py --models LSTM Transformer

# Tune hyperparameters
python run_training.py --trial-name tuned --epochs 200 --batch-size 64 --lr 0.0005 --patience 20

# Full CLI help
python run_training.py --help
```

### Available Models
- **Deep Learning:** LSTM, CNN_LSTM, LSTM_Autoencoder, Transformer
- **Traditional ML:** Isolation_Forest, Random_Forest

## Architecture & Data Flow

### High-Level Pipeline

The pipeline flows through 6 sequential stages:

1. **EDA & Preprocessing** (src/eda.py, src/data.py)
   - Resample CSV files to uniform 15-min intervals
   - Merge multiple CSVs, forward-fill gaps, interpolate
   - Auto-generate cleaned CSV in datasets/site_N_cleaned.csv

2. **Data Loading & Scaling** (SolarDataModule in src/data.py)
   - MinMax scale all features (irradiance_wm2, temperature_c)
   - Time-based train/test split (70/30, no shuffling)
   - Create variable-length sequences (input_steps -> output_steps)

3. **Model Training** (src/trainer.py, src/models/)
   - **Prediction Path:** LSTM, CNN_LSTM, Transformer predict next output_steps of generation; Random Forest does direct regression
   - **Reconstruction Path:** LSTM Autoencoder learns to reconstruct input window
   - **Anomaly Scorer:** Isolation Forest scores points in feature space
   - All DL models use PyTorch with early stopping, LR scheduling, gradient clipping

4. **Anomaly Detection** (src/evaluate.py)
   - **Prediction models:** 3-sigma rule on |actual - predicted| error
   - **Autoencoder:** 3-sigma rule on reconstruction error (threshold from training baseline)
   - **Isolation Forest:** statistical isolation scores

5. **Ensemble & Classification** (src/evaluate.py)
   - Combine 6 anomaly arrays via majority voting (min_votes=3 default)
   - Classify each anomaly into 5 physical categories (High_Irr_Low_Gen, Sudden_Drop, Efficiency_Decline, etc.)

6. **Visualization & Export** (src/visualize.py)
   - Save all plots: loss curves, predictions vs actual, anomaly highlights
   - Generate comparison tables: RMSE, MAE, R2, anomaly counts per model
   - Save trained models, loss history JSON, fitted scaler for inference

### Key Modules

**src/config.py** - Centralized configuration
- DataConfig: features list, target column, preprocessing parameters
- SequenceConfig: input_steps, output_steps (enables variable-length experiments)
- TrainConfig: epochs, batch_size, learning_rate, early stopping patience
- ModelConfig: architecture hyperparams for all 6 models
- AnomalyConfig: sigma multiplier, ensemble voting threshold
- PathConfig: constructs output directory structure

**src/data.py** - Data pipeline
- load_raw_data(): load CSVs by filename pattern
- preprocess_and_merge(): resample to 15min, merge, fill/interpolate gaps
- create_sequences(): sliding window for prediction models
- create_ae_sequences(): create windows for autoencoder
- SolarDataModule: orchestrates all operations, creates PyTorch DataLoaders

**src/trainer.py** - PyTorch training loop
- Generic training with early stopping (patience-based)
- Learning rate scheduler (ReduceLROnPlateau)
- Saves best checkpoint + loss history (JSON + CSV)

**src/models/** - Neural network architectures
- lstm.py: LSTM with FC layers, outputs (output_steps * n_features)
- cnn_lstm.py: 1D CNN -> LSTM -> FC
- autoencoder.py: LSTM encoder-decoder, reconstructs input window
- transformer.py: Positional encoding + TransformerEncoder + FC
- ml_models.py: Isolation Forest, Random Forest from scikit-learn

**src/evaluate.py** - Metrics and anomaly detection
- evaluate_model(): RMSE, MAE, R2 (multi-step averaged)
- detect_anomalies_3sigma(): threshold = mean(error) + sigma * std(error)
- detect_anomalies_reconstruction(): MAE on autoencoder reconstruction
- ensemble_voting(): majority vote across models
- classify_anomalies(): categorize into 5 types based on irradiance/temp/generation relationships

**src/visualize.py** - All plotting functions
- Loss curves, predictions, error distributions, anomaly highlights, ensemble results
- Auto-saves to output dirs with consistent formatting

**src/eda.py** - Data analysis and cleaning
- Data cleaning, merging, gap filling
- Generates EDA plots (availability, distribution, time series)

### Variable Input/Output Windows

All models support flexible sequence lengths. Example: input_steps=6, output_steps=2

```
15-min resolution: 6 steps = 1.5 hours, 2 steps = 30 minutes

Time: [t-5] [t-4] [t-3] [t-2] [t-1] [t]  |  [t+1] [t+2]
      |_____________ INPUT _____________|  |  |__ OUTPUT __|
```

- **Prediction models** (LSTM, CNN-LSTM, Transformer): output shape = (output_steps * n_features)
- **Autoencoder**: reconstructs input window only (ignores output_steps)
- **Random Forest, Isolation Forest**: work on raw features, not sequences

## Output Structure

After training, `outputs/` contains organized results:

```
outputs/{trial_name}/
|- LSTM/
|  |- model/LSTM.pt                    # PyTorch weights
|  |- loss/loss_history.json           # Train & val loss per epoch
|  +- plots/anomaly_detection.png
|- CNN_LSTM/, LSTM_Autoencoder/, Transformer/, Isolation_Forest/, Random_Forest/ (same structure)
|- Ensemble/plots/
|  |- ensemble_anomalies.png
|  +- ensemble_results.json
+- Comparison/
   |- model_comparison.csv             # RMSE, MAE, R2 per model
   |- scaler.joblib                    # Fitted MinMaxScaler for inference
   +- run_config.json
```

## Key Design Patterns

**Modular Configuration via Dataclasses**
All hyperparameters centralized in src/config.py. No hard-coded values scattered through code.

**Time-Based Train/Test Split**
Preserves temporal order (70/30). No random shuffling to prevent data leakage.

**Dual Sequence Handling**
- Prediction models: (input_steps, output_steps) sliding windows
- Autoencoder: input_steps only (reconstruction target = input)
- Test timestamps aligned post-sequence-creation

**Reconstruction vs Prediction Anomalies**
- Autoencoder: reconstruction error threshold from training baseline
- Prediction models: 3-sigma on test set errors only

**Anomaly Classification**
5-category system based on physical signatures:
1. High Irradiance / Low Generation (equipment fault)
2. Sudden Drop (transient)
3. Efficiency Decline (gradual degradation)
4. High Generation / Low Irradiance (sensor error)
5. Other

## Testing & Validation

No automated unit tests. Validation via:
- Manual inspection of output plots and metrics
- Model RMSE/MAE/R2 comparison on test set
- Visual anomaly plot inspection (overlay on time series)
- Ensemble vote distribution (consensus on obvious anomalies)

Historical runs in `outputs/Before/` and `logs/Before/` for reference.

## Data Input Format

Each site requires 4 CSVs in `datasets/site_N/`:

| File | Pattern | Variable | Unit |
|------|---------|----------|------|
| Generation | gen_* | Solar output | kW |
| Irradiance | Irradiance_* | Solar irradiance | W/m2 |
| Temperature | Temp* | Ambient temperature | C |
| Load | load_* | Power consumption | kW |

All resampled to uniform 15-min intervals during preprocessing.

## Hyperparameter Tuning Guide

Common tuning targets:

```bash
# Vary prediction window (longer context)
--input-steps 6 --output-steps 2

# More intensive training
--epochs 200 --batch-size 64 --patience 20

# Lower learning rate for fine-tuning
--lr 0.0005

# Anomaly sensitivity (lower sigma = more anomalies)
--sigma 2.5

# Test specific models
--models LSTM Transformer
```

Results auto-saved with trial name for comparison.

## Inference with Saved Models

```python
import torch
import joblib
from src.data import create_sequences
from src.models import LSTMModel

# Load training scaler
scaler = joblib.load("outputs/{trial_name}/Comparison/scaler.joblib")

# Load and initialize model
model = LSTMModel(n_features=3, output_steps=1)
model.load_state_dict(torch.load("outputs/{trial_name}/LSTM/model/LSTM.pt"))
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

## Notable Implementation Details

- **Logging**: Dual console + file output via TeeLogger in logs/
- **Trial naming**: Auto-generates run_1, run_2, ... if not specified
- **Anomaly categories**: Dynamic detection based on irradiance/temp/generation correlations
- **Ensemble alignment**: Handles mismatched sequence lengths (uses shortest common length)
- **GPU auto-detection**: Automatically uses CUDA if available, falls back to CPU

## References

Research papers and reference projects in References/ directory.
Historical results in outputs/Before/ and documents/ (markdown summaries + PowerPoint).
