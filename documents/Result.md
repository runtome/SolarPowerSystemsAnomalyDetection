# Solar Power Systems Anomaly Detection - Results

## 1. Data Overview

| Item | Detail |
|------|--------|
| **Site** | Site 1 |
| **Date Range** | 2025-06-01 to 2025-12-01 (6 months) |
| **Raw Data** | 4 CSV files: generation, irradiance, temperature, load |
| **Raw Resolution** | ~1 min (gen/irr/temp), 15 min (load) |
| **Resampled Resolution** | 15-minute intervals (uniform) |
| **Cleaned Rows** | 15,013 (dropped 2,651 rows with large gaps) |
| **Train Set** | 10,509 rows (Jun 1 - Oct 6, 2025) |
| **Test Set** | 4,504 rows (Oct 6 - Dec 1, 2025) |
| **GPU Used** | NVIDIA GeForce RTX 3060 Ti |
| **Framework** | PyTorch 2.9.1+cu126 |

### Variable Statistics (Cleaned Data)

| Variable | Mean | Std | Min | Max |
|----------|------|-----|-----|-----|
| Generation (kW) | 96.53 | 150.86 | 0.00 | 638.00 |
| Irradiance (W/m2) | 157.76 | 243.96 | 0.00 | 1,073.00 |
| Temperature (C) | 28.73 | 3.01 | 0.00 | 37.05 |
| Load (kW) | - | - | - | - |

---

## 2. EDA Key Findings

### Correlations with Generation

| Feature | Correlation |
|---------|-------------|
| Irradiance (W/m2) | **0.901** (very strong) |
| Temperature (C) | **0.564** (moderate) |
| Load (kW) | 0.026 (negligible) |

### Observations

- **Generation is strongly driven by irradiance** (r=0.90), confirming irradiance as the primary predictor.
- **Temperature has moderate positive correlation** (r=0.56) with generation, likely because both are higher during sunny daytime hours.
- **Load is independent** of solar generation (r=0.03), following its own consumption pattern.
- **Generation efficiency** (Gen/Irradiance during daytime): mean = 0.71, std = 0.31.
- **1,655 data points** showed generation with zero sunlight - potential sensor/data logging issues.
- **0 data points** had high irradiance with zero generation (no obvious equipment failures in simple check).
- **Surplus vs Deficit:** Solar generation exceeds load only 10.8% of the time; 89.2% of the time the site relies on grid power.

### Missing Data Handling

- After merging 4 datasets at 15-min intervals: ~75% nulls in gen/irr/temp (due to ~1-min data not aligning perfectly with all 15-min windows).
- Applied forward-fill + time interpolation (up to 1-hour gaps).
- Remaining ~15% nulls dropped, yielding 15,013 clean rows.

---

## 3. Model Results

### 3.1 Prediction Performance (Test Set)

| Model | RMSE | MAE | R2 | Parameters |
|-------|------|-----|-----|------------|
| Random Forest | 83.37 | 34.80 | 0.6346 | - |
| LSTM | 32.46 | 15.43 | 0.9445 | 53,123 |
| CNN-LSTM | 32.87 | 15.45 | 0.9431 | 8,819 |
| **Transformer** | **32.05** | **13.76** | **0.9459** | ~37K |

**Best prediction model: Transformer** (lowest RMSE=32.05, lowest MAE=13.76, highest R2=0.9459)

All three deep learning models significantly outperform Random Forest, achieving R2 > 0.94 (explaining 94%+ of variance in solar generation).

### 3.2 Feature Importance (Random Forest)

| Feature | Importance |
|---------|------------|
| Irradiance (W/m2) | **94.6%** |
| Temperature (C) | 5.4% |

Irradiance dominates prediction, consistent with the high correlation (0.90) found in EDA.

---

## 4. Anomaly Detection Results

### 4.1 Anomalies Detected per Model

| Model | Method | Anomalies | % of Test |
|-------|--------|-----------|-----------|
| Isolation Forest | Isolation score (contamination=2%) | 244 | 5.4% |
| Random Forest | 3-sigma on prediction MAE | 159 | 3.5% |
| LSTM | 3-sigma on prediction MAE | 102 | 2.3% |
| **CNN-LSTM** | **3-sigma on prediction MAE** | **94** | **2.1%** |
| LSTM Autoencoder | 3-sigma on reconstruction error | 265 | 5.9% |
| Transformer | 3-sigma on prediction MAE | 100 | 2.2% |

- **Most conservative:** CNN-LSTM (94 anomalies, 2.1%)
- **Most sensitive:** LSTM Autoencoder (265 anomalies, 5.9%)
- Deep learning prediction models (LSTM, CNN-LSTM, Transformer) converge around 2-2.3%, suggesting consistent anomaly estimation.

### 4.2 Ensemble (Majority Vote)

Combined all 6 models; flagged as anomaly if **>= 3 models agree**.

| Agreement | Data Points |
|-----------|-------------|
| 0 models (normal) | 3,939 |
| 1 model only | 311 |
| 2 models | 119 |
| **3 models** | **114** |
| **4 models** | **13** |
| **5 models** | **3** |
| **6 models (all agree)** | **1** |
| **Total ensemble anomalies** | **131 (2.9%)** |

The ensemble approach provides a balanced result: 131 high-confidence anomalies where at least 3 independent models agree. This reduces false positives compared to any single model.

---

## 5. Method Summary

### Anomaly Detection Approach

The core idea: **train a model to learn the normal relationship between weather conditions (irradiance, temperature) and solar generation.** When actual generation deviates significantly from what the model predicts, it signals an anomaly (equipment fault, panel degradation, inverter failure, shading, etc.).

| Model | Type | How It Works |
|-------|------|--------------|
| **Isolation Forest** | ML, Unsupervised | Finds data points that are statistically "easy to isolate" in multi-dimensional feature space |
| **Random Forest** | ML, Supervised Regression | Predicts generation from weather; large errors = anomalies |
| **LSTM** | DL, Sequential | Uses 1-hour history (4 timesteps) to predict next generation value |
| **CNN-LSTM** | DL, Hybrid | CNN extracts local patterns from the sequence, LSTM captures temporal dynamics |
| **LSTM Autoencoder** | DL, Reconstruction | Learns to compress and reconstruct normal sequences; anomalies have high reconstruction error |
| **Transformer** | DL, Self-Attention | Uses multi-head attention to capture dependencies across all timesteps simultaneously |

### Anomaly Threshold

All prediction-based models use the **3-sigma rule**:
- Compute MAE (Mean Absolute Error) between prediction and actual
- Threshold = mean(MAE) + 3 * std(MAE)
- Points exceeding threshold are flagged as anomalies

---

## 6. Conclusion

1. **Deep learning models (LSTM, CNN-LSTM, Transformer) dramatically outperform traditional ML (Random Forest)** for solar generation prediction, achieving R2 > 0.94 vs 0.63.

2. **The Transformer model achieved the best overall prediction accuracy** (RMSE=32.05, MAE=13.76, R2=0.9459), benefiting from its self-attention mechanism that captures dependencies across all time steps.

3. **All DL models perform similarly** (R2: 0.943-0.946), suggesting the prediction task is well-defined and the data quality is sufficient. CNN-LSTM achieves comparable results with the fewest parameters (8,819), making it the most efficient model.

4. **Ensemble voting (majority of 6 models) identified 131 high-confidence anomalies (2.9%)** in the test period (Oct-Dec 2025). These represent time points where solar generation significantly deviated from expected output given weather conditions.

5. **Irradiance is the dominant predictor** (94.6% feature importance), confirming the physics-based relationship between solar radiation and power generation.

6. **The 1,655 "generation with no sunlight" data points** found during EDA warrant investigation as potential sensor calibration or data logging issues.

### Saved Models

All models are saved in `models/` directory for deployment to the remaining 3 sites:

```
models/
  isolation_forest.joblib    # Isolation Forest
  random_forest.joblib       # Random Forest
  scaler.joblib              # MinMaxScaler (fitted on Site 1 train data)
  model_configs.joblib       # Feature/target configuration
  lstm_model.pt              # LSTM (PyTorch)
  cnn_lstm_model.pt          # CNN-LSTM (PyTorch)
  lstm_autoencoder.pt        # LSTM Autoencoder (PyTorch)
  transformer_model.pt       # Transformer (PyTorch)
```

### Next Steps

- Extend training and evaluation to all 4 sites
- Fine-tune anomaly threshold (sigma) per site based on operational feedback
- Classify anomaly types (inverter fault, panel degradation, shading, sensor error)
- Build a real-time monitoring dashboard for continuous anomaly detection
