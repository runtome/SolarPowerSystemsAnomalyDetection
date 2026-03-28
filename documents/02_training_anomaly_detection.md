## Goal
- Train Isolation Forest, Random Forest regression, and four PyTorch models (LSTM, CNN-LSTM, LSTM Autoencoder, Transformer) to predict solar generation from irradiance and temperature so that large prediction or reconstruction errors (3-sigma rule) reveal anomalies.

## Steps
1. Load `datasets/site_1_cleaned.csv` and keep the modeling columns `irradiance_wm2`, `temperature_c`, `generation_kw`.
2. Split chronologically into train/test sets, scale with `MinMaxScaler`, and plot the target split.
3. Define helpers for creating sequences (4 timesteps = 1 hour), tensor conversion, inverse scaling, 3-sigma detection, training loop (early stopping + LR scheduler), and plotting.
4. Build DataLoaders from the scaled sequences and persist the actual test targets/timestamps for evaluation plots.
5. Fit Isolation Forest for unsupervised flags and visualize anomaly scores.
6. Train Random Forest regression, evaluate RMSE/MAE/R², apply 3-sigma on MAE, and plot actual vs predicted plus flagged points.
7. Define each PyTorch architecture, train with `train_model`, infer on the test set, inverse-scale outputs, evaluate metrics, run 3-sigma anomaly detection, and visualize errors/anomalies (for the autoencoder, use reconstruction error directly).
8. Compare all models’ metrics and anomaly counts, build a majority-vote ensemble (>=3 models agree), and plot ensemble votes.
9. Save sklearn models (Isolation Forest, Random Forest, scaler) and PyTorch weights plus model configs under `models/`.

## Key Variables & Functions
- `FEATURES`, `TARGET`, `ALL_COLS`: column names used throughout.
- `TIMESTEPS` (4), `EPOCHS` (100), `BATCH_SIZE`, `LR`, `PATIENCE`, `N_FEATURES`, `TARGET_IDX`: training hyperparameters/constants.
- `create_sequences`, `to_tensors`, `inverse_target`, `detect_anomalies_3sigma`, `evaluate_model`, `plot_anomalies`, `plot_loss`: general-purpose utilities for data prep, scoring, and visualization.
- `train_model`: reusable PyTorch training loop with gradient clipping, ReduceLROnPlateau scheduler, and early stopping.
- `predict`: inference helper returning NumPy arrays from tensors.
- Architecture classes: `LSTMModel`, `CNNLSTM`, `LSTMAutoencoder`, `PositionalEncoding`, `TransformerModel`.
- `all_results` / `all_anomaly_counts`: accumulate metrics/anomaly counts for comparison/ensemble.
- `ensemble_matrix`, `ensemble_votes`, `ensemble_anomalies`: construct majority-vote detection.
- `MODEL_DIR`, joblib/torch saves: persist each trained model + scaler + configs for later inference.

## Execution
1. Run `01_EDA.ipynb` to generate `datasets/site_1_cleaned.csv`.
2. Execute the training notebook: `jupyter nbconvert --execute 02_training_anomaly_detection.ipynb`.
3. You will see printed metrics for each model, multiple matplotlib plots (loss, anomaly detections, ensemble votes), and saved artifacts under `models/` (sklearn + PyTorch + configs).

## Suggestions
- Accept CLI/config input for noise threshold (sigma multiplier), model hyperparameters, and output directories so retraining for different sites or deployment targets can be automated.
- Separate plotting from training (headless mode) so CI systems can run the notebook without rendering while still generating numeric logs. 
