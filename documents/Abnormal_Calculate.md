# Anomaly Indicator Calculation

Here is the full breakdown of how each anomaly is calculated:

---

## 1. High Irr / Low Gen

**Meaning:** Daytime has strong sunlight but solar output is unexpectedly low — possible panel shading, soiling, or degradation.

**Calculation:**
1. Filter daytime rows: `irradiance > 50 W/m²`
2. Compute daytime thresholds:
   - `irr_q75` = 75th percentile of daytime irradiance
   - `gen_q10` = 10th percentile of daytime generation
3. Flag row if: `irradiance > irr_q75` AND `generation < gen_q10`

A row is only flagged if it simultaneously has **high irradiance** (top 25%) and **low generation** (bottom 10%).

---

## 2. Sudden Drop

**Meaning:** Generation falls sharply between two consecutive 15-min readings — possible inverter trip, cloud shadow, or cable fault.

**Calculation (daytime only, irradiance > 50):**
1. Check that two rows are consecutive: `time_diff ≤ 20 min`
2. Compute: `gen_diff = generation[t] − generation[t−1]`
3. Compute: `pct_drop = gen_diff / generation[t−1]`
4. Flag if: `pct_drop < −0.50` AND `|gen_diff| > 50 kW`

Both conditions must be true: the drop must be **> 50% of the previous reading** and **> 50 kW in absolute terms** (prevents flagging small values).

---

## 3. Efficiency Decline

**Meaning:** System is producing less per unit of irradiance than its own recent historical trend — gradual degradation signal.

**Calculation (daytime only, irradiance > 50):**
1. Compute per-row: `efficiency = generation_kw / irradiance_wm2`
2. Compute rolling 96-step (~1 day) statistics: `roll_mean`, `roll_std`
3. Flag if: `efficiency < roll_mean − 2 × roll_std`

This is a **2-sigma rule** against the rolling baseline — only flags when efficiency is unusually low compared to the recent 24-hour window, not the global average.

---

## 4. Gen Spike

**Meaning:** Generation is far higher than normal for its irradiance level — possible sensor error or meter malfunction.

**Calculation (daytime only):**
1. Bin irradiance into 10 quantile-equal buckets
2. For each bin, compute `mean` and `std` of generation
3. Flag if: `generation > mean + 3 × std` within its bin

By comparing within irradiance bins, the test controls for the irradiance level — a high reading on a bright day is only flagged if it's high **relative to other bright-day readings**.

---

## 5. Gen / Zero Irr

**Meaning:** Generation is recorded as > 0 when irradiance ≤ 0 — physically impossible, indicates a sensor or logging error.

**Calculation:** Direct filter on the raw merged data (captured before the zeroing step):
```
generation_kw > 0  AND  irradiance_wm2 ≤ 0
```

This is the only category that runs on **pre-zeroing data** — after `zero_no_sun_generation()` runs, these rows would show generation = 0 and no longer be detectable.

---

## Summary

| Category | Filter | Threshold Logic |
|---|---|---|
| High Irr / Low Gen | daytime (irr > 50) | irr > 75th pct AND gen < 10th pct |
| Sudden Drop | daytime, consecutive rows | pct_drop < −50% AND abs_drop > 50 kW |
| Efficiency Decline | daytime | efficiency < rolling_mean − 2σ (96-step window) |
| Gen Spike | daytime, per irr-bin | generation > bin_mean + 3σ |
| Gen / Zero Irr | all hours | gen > 0 AND irr ≤ 0 (pre-zeroing snapshot) |
