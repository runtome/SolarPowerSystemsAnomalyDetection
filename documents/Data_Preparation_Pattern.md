# Data Preparation Pattern

## Step 1 — Raw columns (3 features total)

```
┌─────────┬────────────────┬──────────────────────────┐
│  Role   │     Column     │       Description        │
├─────────┼────────────────┼──────────────────────────┤
│ Feature │ irradiance_wm2 │ Solar irradiance (W/m²)  │
├─────────┼────────────────┼──────────────────────────┤
│ Feature │ temperature_c  │ Ambient temperature (°C) │
├─────────┼────────────────┼──────────────────────────┤
│ Target  │ generation_kw  │ Solar generation (kW)    │
└─────────┴────────────────┴──────────────────────────┘
```

All 3 are scaled together with a single `MinMaxScaler`. So `n_features = 3`.

---

## Step 2 — Train/Test split (time-based, no shuffle)

```
Total data  →  70% train  |  30% test
                ↓
        fit_transform()   transform()
```

---

## Step 3 — Sliding window sequences

### Prediction models (LSTM, CNN-LSTM, Transformer) — `create_sequences()`

```
Default: input_steps=4, output_steps=1  (15-min data → 1 hour in, 15 min out)

       timestep:   t-3   t-2   t-1    t   |  t+1
       features: [irr, temp, gen]  × 4   |  [irr, temp, gen] × 1

X shape:  (samples,  input_steps,   n_features)  =  (N, 4, 3)
Y shape:  (samples,  output_steps,  n_features)  =  (N, 1, 3)
           └─ flattened to (N, 3) before entering model
```

The model predicts **all 3 features** at the next step, but only `generation_kw` (index 2) is used for anomaly scoring.

### LSTM Autoencoder — `create_ae_sequences()`

```
X shape:  (N, input_steps, n_features)  =  (N, 4, 3)
Y = X    (reconstruction target IS the input — no future step)
```

### Isolation Forest / Random Forest — no sequences

```
X shape:  (N, n_features)  =  (N, 3)   ← raw scaled rows, no windows
```

---

## Step 4 — What the model sees at training time

```
┌─────────────────────────────┬───────────────┬────────────────────────────────────────┐
│            Model            │   X fed in    │                Y/target                │
├─────────────────────────────┼───────────────┼────────────────────────────────────────┤
│ LSTM, CNN-LSTM, Transformer │ (batch, 4, 3) │ (batch, 3) — next step, all features   │
├─────────────────────────────┼───────────────┼────────────────────────────────────────┤
│ LSTM Autoencoder            │ (batch, 4, 3) │ same (batch, 4, 3) — reconstruct input │
├─────────────────────────────┼───────────────┼────────────────────────────────────────┤
│ Random Forest               │ (N, 3)        │ (N,) — generation_kw column only       │
├─────────────────────────────┼───────────────┼────────────────────────────────────────┤
│ Isolation Forest            │ (N, 3)        │ unsupervised — no y label              │
└─────────────────────────────┴───────────────┴────────────────────────────────────────┘
```

---

## Visual summary (default `input_steps=4, output_steps=1`)

```
Raw CSV row: [irradiance_wm2,  temperature_c,  generation_kw]
                                    ↓ MinMaxScale
Scaled row:  [0.72,             0.45,           0.61         ]

Sliding window (step i):
  X[i] = scaled[i-4 : i]      shape (4, 3)  ← past 1 hour
  Y[i] = scaled[i   : i+1]    shape (1, 3)  ← next 15 min
```

> **Key insight:** X always contains all 3 columns (not just the 2 features), because the model needs to see past `generation_kw` values as context to predict the next one.

---

## Step 5 — Anomaly judgment after prediction

### Prediction models (LSTM, CNN-LSTM, Transformer, Random Forest)

```
Given:   real data at [t-3, t-2, t-1, t]   (past 1 hour)
Predict: generation_kw at [t+1]             (next 15 min)
Compare: |actual_t+1 − predicted_t+1|  →  error

If error > threshold  →  timestep t+1 is ANOMALY
```

The threshold is calculated from the errors across the entire test set:

```python
# src/evaluate.py  line 33-35
mae = np.abs(actual - predicted)
threshold = mae.mean() + sigma * mae.std()   # default sigma=3.0
anomalies = (mae > threshold).astype(int)
```

So the model is saying: **"I expected this much generation, but the real value is so far off it cannot be normal."**

---

### LSTM Autoencoder — different logic (no future prediction)

The autoencoder does **not** predict the next step. Instead:

```
Given:   real data at [t-3, t-2, t-1, t]
Output:  reconstructed [t-3, t-2, t-1, t]   (same window, rebuilt)

If reconstruction_error > threshold  →  the window itself is ANOMALY
```

The threshold here comes from **training data** (not test set), so it represents what a "normal" window looks like to the model.

---

### Isolation Forest — no prediction at all

```
Input:  raw feature row  [irradiance_wm2, temperature_c, generation_kw]
Output: isolation score   (how hard was this point to isolate?)

If score is very isolated  →  ANOMALY
```

No comparison to a predicted value — purely statistical in feature space.

---

### Summary table

```
┌─────────────────────────────┬──────────────────────────────────────────────────────┐
│           Model             │  How it judges anomaly at timestep t+1               │
├─────────────────────────────┼──────────────────────────────────────────────────────┤
│ LSTM, CNN-LSTM, Transformer │ |actual_t+1 − predicted_t+1| > mean_err + 3σ         │
├─────────────────────────────┼──────────────────────────────────────────────────────┤
│ Random Forest               │ |actual_t+1 − predicted_t+1| > mean_err + 3σ         │
├─────────────────────────────┼──────────────────────────────────────────────────────┤
│ LSTM Autoencoder            │ reconstruction_error of window > train_mean + 3σ     │
├─────────────────────────────┼──────────────────────────────────────────────────────┤
│ Isolation Forest            │ isolation score in feature space (no prediction)     │
└─────────────────────────────┴──────────────────────────────────────────────────────┘
```

---

## Step 6 — Ensemble voting

After all 6 models produce their own anomaly flag per timestep, the ensemble combines them:

```python
# If ≥ 3 models agree a timestep is anomaly  →  final ANOMALY label
ensemble = (votes >= min_votes)   # default min_votes = 3
```

So the final answer needs at least 3 out of 6 models to agree before flagging a point.
