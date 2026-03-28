# Anomaly Detection in Solar Power Systems Using Machine Learning and Deep Learning

## Presentation Summary

---

## Slide 1: Problem Statement

**Goal:** Detect anomalies in solar power generation systems that indicate equipment faults, panel degradation, inverter failures, or sensor errors.

**Why it matters:**
- Undetected faults reduce energy output and revenue
- Early detection enables preventive maintenance
- Reduces downtime and extends equipment lifespan

**Approach:** Train models to learn the normal relationship between weather conditions and solar generation. When actual output deviates significantly from expected output, flag it as an anomaly.

---

## Slide 2: Data Description

**Source:** Site 1 solar power plant monitoring system (sample of 4 total sites)

| Sensor | Variable | Unit | Resolution | Rows |
|--------|----------|------|------------|------|
| Generation | Solar production (actual output) | kW | ~1 min | 26,226 |
| Irradiance | Solar radiation intensity | W/m2 | ~1 min | 26,177 |
| Temperature | Ambient temperature | C | ~1 min | 26,195 |
| Load | Power consumption | kW | 15 min | 17,367 |

**Period:** June 1, 2025 - December 1, 2025 (6 months)

---

## Slide 3: Methodology Overview (Pipeline)

```
Step 1: Data Collection        4 separate CSV files per site
            |
Step 2: Data Preprocessing     Resample to 15-min, merge, handle nulls
            |
Step 3: EDA                    Correlations, distributions, patterns
            |
Step 4: Feature Engineering    Time sequences (1-hour sliding window)
            |
Step 5: Model Training         2 ML + 4 DL models
            |
Step 6: Anomaly Detection      3-sigma threshold on prediction error
            |
Step 7: Ensemble Voting        Majority vote across all 6 models
```

---

## Slide 4: Data Preprocessing

### Challenge: Different Time Resolutions
- Generation, Irradiance, Temperature: **~1-minute** intervals
- Load: **15-minute** intervals

### Solution:
1. **Resample** all data to uniform **15-minute** intervals (mean aggregation)
2. **Merge** all 4 variables into a single dataframe (outer join on timestamp)
3. **Handle missing values:**
   - Forward-fill small gaps (up to 1 hour)
   - Time-based interpolation for remaining small gaps
   - Drop rows with large gaps

### Result:
| Stage | Rows |
|-------|------|
| After merge | 17,664 |
| After cleaning | **15,013** |
| Dropped | 2,651 (15%) |

---

## Slide 5: EDA - Key Findings

### Correlation with Solar Generation

| Feature | Correlation | Interpretation |
|---------|-------------|----------------|
| Irradiance | **0.901** | Very strong - primary driver of generation |
| Temperature | **0.564** | Moderate - co-varies with sunny conditions |
| Load | 0.026 | Negligible - consumption is independent |

### Important Observations
- **Solar generation peaks at midday**, following irradiance pattern (clear daily cycle)
- **Generation efficiency** (output per unit irradiance): mean = 0.71 kW/(W/m2)
- **1,655 data points** recorded generation with zero sunlight (potential sensor issue)
- Solar surplus covers only **10.8%** of time; site relies on grid **89.2%** of time

---

## Slide 6: EDA - Visualizations

*(Use screenshots from the notebook for these slides)*

- **Time series plots:** All 4 variables over 6 months showing daily/seasonal patterns
- **One-week zoom:** Clear day/night cycles visible in generation and irradiance
- **Correlation heatmap:** Strong irradiance-generation block
- **Scatter plot:** Generation vs Irradiance shows strong linear relationship
- **Hourly patterns:** Bell-curve for generation/irradiance, stable temperature, variable load
- **Monthly trends:** Seasonal variation across Jun-Dec
- **Net power chart:** Green (surplus) vs Red (deficit) areas

---

## Slide 7: Model Architecture

### Train/Test Split (Time-Based, No Shuffle)
- **Train:** 70% (10,509 rows) - Jun to Oct 2025
- **Test:** 30% (4,504 rows) - Oct to Dec 2025

### Input/Output
- **Features (X):** Irradiance (W/m2), Temperature (C)
- **Target (Y):** Generation (kW)
- **Time window:** 4 timesteps (1 hour of history) for DL models
- **Scaling:** MinMaxScaler (0 to 1)

### 6 Models Trained

| # | Model | Category | Approach |
|---|-------|----------|----------|
| 1 | Isolation Forest | ML, Unsupervised | Detects statistically isolated points in feature space |
| 2 | Random Forest | ML, Supervised | Regression-based prediction; error = anomaly |
| 3 | LSTM | DL, Sequential | Learns temporal patterns from 1-hour windows |
| 4 | CNN-LSTM | DL, Hybrid | CNN extracts local features + LSTM captures time dynamics |
| 5 | LSTM Autoencoder | DL, Reconstruction | Compresses and reconstructs normal patterns |
| 6 | Transformer | DL, Self-Attention | Attention mechanism captures long-range dependencies |

---

## Slide 8: Deep Learning Model Details

### LSTM (53,123 parameters)
- 2-layer LSTM (hidden=64) with dropout (0.2)
- Fully connected output layer
- Takes last timestep hidden state for prediction

### CNN-LSTM (8,819 parameters) - Most Efficient
- 2 Conv1D layers (32 and 16 filters) for local pattern extraction
- 1-layer LSTM (hidden=32) for temporal modeling
- Smallest model, comparable performance

### LSTM Autoencoder (16,611 parameters)
- Encoder: 2 LSTM layers (32 -> 16) compressing the sequence
- Bottleneck: 16-dim representation
- Decoder: 2 LSTM layers (16 -> 32) reconstructing the input
- Anomaly = high reconstruction error

### Transformer (~37K parameters)
- Input projection + sinusoidal positional encoding
- 2 Transformer encoder layers (2 attention heads, d_model=32)
- Global average pooling + FC output
- Best prediction accuracy

---

## Slide 9: Anomaly Detection Method

### 3-Sigma Rule (Applied to All Models)

```
1. Prediction Error = |Actual Generation - Predicted Generation|
2. Threshold = Mean(Error) + 3 x Std(Error)
3. If Error > Threshold --> ANOMALY
```

**Why 3-sigma?**
- In a normal distribution, 99.7% of data falls within 3 standard deviations
- Points beyond 3-sigma are statistically rare and likely anomalous
- Can be tuned: 2-sigma (more sensitive) or 4-sigma (fewer false positives)

### For LSTM Autoencoder:
- Instead of prediction error, uses **reconstruction error**
- If the model cannot reconstruct a sequence well, it hasn't seen that pattern during training (= abnormal)

### For Isolation Forest:
- No prediction needed; directly isolates unusual multi-dimensional points
- contamination parameter set to 2% (prior estimate of anomaly rate)

---

## Slide 10: Prediction Performance Results

| Model | RMSE | MAE | R2 |
|-------|------|-----|-----|
| Random Forest | 83.37 | 34.80 | 0.6346 |
| LSTM | 32.46 | 15.43 | 0.9445 |
| CNN-LSTM | 32.87 | 15.45 | 0.9431 |
| **Transformer** | **32.05** | **13.76** | **0.9459** |

### Key Takeaways:
- **Deep learning models outperform Random Forest** by a large margin (R2: 0.94 vs 0.63)
- **Transformer is the best predictor** (lowest error, highest R2)
- All 3 DL models perform similarly (R2 ~0.94), showing the task is well-defined
- **CNN-LSTM is the most parameter-efficient** (8,819 params vs 53,123 for LSTM)
- Feature importance confirms: **Irradiance = 94.6%**, Temperature = 5.4%

---

## Slide 11: Anomaly Detection Results

### Individual Model Results

| Model | Anomalies Detected | % of Test Set |
|-------|-------------------|---------------|
| LSTM Autoencoder | 265 | 5.9% |
| Isolation Forest | 244 | 5.4% |
| Random Forest | 159 | 3.5% |
| LSTM | 102 | 2.3% |
| Transformer | 100 | 2.2% |
| CNN-LSTM | 94 | 2.1% |

- Unsupervised methods (Autoencoder, Isolation Forest) detect more anomalies (higher sensitivity)
- Prediction-based DL models converge at ~2.1-2.3% (more conservative, fewer false positives)

---

## Slide 12: Ensemble Anomaly Detection

### Majority Voting: Flag as anomaly if >= 3 of 6 models agree

| Agreement Level | Data Points | Interpretation |
|----------------|-------------|----------------|
| 0 models | 3,939 (87.5%) | Clearly normal |
| 1 model | 311 (6.9%) | Likely normal (single model noise) |
| 2 models | 119 (2.6%) | Borderline |
| **3+ models** | **131 (2.9%)** | **High-confidence anomaly** |

### Breakdown of High-Confidence Anomalies:
- 3 models agree: 114 points
- 4 models agree: 13 points
- 5 models agree: 3 points
- All 6 models agree: 1 point (strongest anomaly)

**Total: 131 high-confidence anomalies (2.9% of test data)**

The ensemble reduces false positives while maintaining detection of genuine anomalies.

---

## Slide 13: What Do These Anomalies Mean?

### Types of Anomalies Detected:

| Anomaly Pattern | Possible Cause |
|----------------|----------------|
| High irradiance + Low generation | Panel fault, inverter failure, shading |
| Sudden generation drop | Equipment trip, grid disconnection |
| Generation with zero irradiance | Sensor malfunction, data logging error |
| Gradual efficiency decline | Panel degradation, dust accumulation |
| Unusual generation spikes | Sensor calibration drift |

### Practical Impact:
- **131 flagged time points** in 2 months of test data
- Each represents a 15-minute window where the system behaved abnormally
- Operations team can investigate clustered anomalies for root cause

---

## Slide 14: Technical Setup

| Component | Specification |
|-----------|--------------|
| Language | Python 3.14 |
| DL Framework | PyTorch 2.9.1 + CUDA 12.6 |
| ML Framework | scikit-learn |
| GPU | NVIDIA GeForce RTX 3060 Ti |
| Training Data | 10,509 samples (70%) |
| Test Data | 4,504 samples (30%) |
| Time Window | 4 timesteps (1 hour) |
| Batch Size | 32 |
| Optimizer | Adam (lr=0.001 with ReduceLROnPlateau) |
| Early Stopping | Patience = 15 epochs |
| Gradient Clipping | Max norm = 1.0 |

---

## Slide 15: Conclusion

1. **Irradiance is the dominant factor** in solar generation prediction (r=0.90, 94.6% importance), confirming physical expectations.

2. **Deep learning significantly outperforms traditional ML** for this task: R2 improved from 0.63 (Random Forest) to **0.95 (Transformer)** - a 49% reduction in unexplained variance.

3. **The Transformer model achieved the best prediction** (RMSE=32.05, R2=0.9459), while **CNN-LSTM is the most efficient** (comparable accuracy with 6x fewer parameters).

4. **Ensemble voting across 6 models identified 131 high-confidence anomalies** (2.9%) in the Oct-Dec 2025 test period, providing robust detection with reduced false positives.

5. The prediction-based anomaly detection approach is **interpretable**: an anomaly means "the system produced significantly less (or more) power than expected given the current weather conditions."

---

## Slide 16: Future Work

| Priority | Task | Description |
|----------|------|-------------|
| High | Multi-site deployment | Extend to all 4 sites (currently only Site 1) |
| High | Threshold tuning | Calibrate sigma per site using operator feedback |
| Medium | Anomaly classification | Categorize anomaly type (inverter, panel, sensor, etc.) |
| Medium | Real-time pipeline | Build streaming inference for continuous monitoring |
| Low | Transfer learning | Pre-train on Site 1, fine-tune on Sites 2-4 |
| Low | Explainability | Add SHAP/attention visualization for anomaly root cause |

---

## Appendix: Project Structure

```
SolarPowerSystemsAnomalyDetection/
  01_EDA.ipynb                          # Exploratory Data Analysis
  02_training_anomaly_detection.ipynb    # Model Training (PyTorch)
  Result.md                             # Detailed results
  Presentation_Summary.md               # This document
  requirements.txt                      # Python dependencies
  datasets/
    site_1/                             # Raw data (4 CSV files)
    site_1_cleaned.csv                  # Cleaned merged data
  models/                               # Saved trained models
    lstm_model.pt
    cnn_lstm_model.pt
    lstm_autoencoder.pt
    transformer_model.pt
    isolation_forest.joblib
    random_forest.joblib
    scaler.joblib
    model_configs.joblib
  References/                           # Research papers & reference project
```
