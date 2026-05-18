# Why Anomaly Counts Can Increase After Cleaning

---

## 1. Gap-filling creates new data rows

`df_merged_valid` (before cleaning) only includes rows where **all sensors reported at the same time** — pure observations, no gaps. `df_clean` has additional rows that were created by `ffill` and interpolation to fill missing intervals.

These filled-in rows are **synthetic estimates**, not real measurements. They get evaluated by the anomaly detectors just like real rows, and some land in "bad" positions relative to their neighbors.

---

## 2. Why Sudden Drop +2 and High Irr/Low Gen +1

The interpolation stitches over gaps smoothly — but at the edges of a gap, it creates a transition that didn't exist in the raw data. For example:

```
Raw (gap):       [500 kW] --- GAP --- [100 kW]
After fill:      [500 kW] → [400] → [300] → [200] → [100 kW]
```

The filled ramp creates new consecutive pairs like `500 → 400` or `300 → 100` that can cross the "sudden drop > 50% AND > 50 kW" threshold.

---

## 3. Why Efficiency Decline +7 (largest change)

This category uses a **96-step (1-day) rolling window** on efficiency (gen/irradiance). More rows from gap-filling means the rolling mean and std are computed over a slightly different distribution. Borderline points that were just inside the normal range in raw data can slip outside `rolling_mean - 2*std` after cleaning because the rolling baseline shifted.

---

## Summary

| Cause | Effect |
|---|---|
| ffill/interpolation adds synthetic rows | More rows = more candidates for anomaly detection |
| Interpolated edge-of-gap transitions | Creates artificial sudden drops |
| Rolling window computed on more rows | Slightly shifts the "normal" baseline |

---

## Is this a problem?

Not really — the increase is tiny (+7/+2/+1 out of ~15,000 rows). The "Visual Inspection" counts reflect anomalies in **raw measurements only**, while "After Cleaning" counts include a small number of interpolation artifacts. That's why showing both in the PPTX is useful: if you see a big difference, it indicates heavy gap-filling in that site's data.
