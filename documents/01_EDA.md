## Goal
- Understand Site 1 solar generation, irradiance, temperature, and load (June–December 2025) to clean, align, and inspect the data before anomaly modeling.

## Steps
1. Load the four CSV streams (`gen`, `irradiance`, `temperature`, `load`), rename columns, and log coverage.
2. Quantify nulls and time gaps per dataset.
3. Resample the 1‑minute streams to 15 minutes, outer-join all variables, and inspect missing data.
4. Clean the merged `df` by forward-filling/interpolating up to 1 hour, then drop remaining NaNs.
5. Compute summary statistics and plot long- and short-term time series.
6. Visualize distributions (histograms/box plots) to reveal skew/outliers.
7. Correlation matrix + scatter plots confirm irradiance dominates generation.
8. Aggregate hourly/monthly patterns via mean ± Std.
9. Estimate daytime efficiency (`generation / irradiance`) and its temperature dependence.
10. Highlight candidate anomalies: high irradiance/low generation and generation with no sunlight.
11. Derive net power (`generation - load`) and summarize surplus/deficit shares.
12. Save cleaned data (`datasets/site_1_cleaned.csv`) for modeling.

## Key Variables & Outputs
- `DATA_DIR`: `datasets/site_1/`.
- `gen`, `irr`, `temp`, `load`: raw DataFrames with `datetime`.
- `df`: merged resampled dataset; `df_clean`: filled/interpolated version used in analyses.
- `cols`, `colors`, `titles`: helper lists for looping over the four main variables.
- `daytime`: subset with irradiance > 10; used to compute `efficiency`.
- `high_irr_low_gen`, `gen_no_sun`: DataFrames of manually defined anomaly candidates.
- `net_power`: generation minus load for surplus/deficit analysis.
- `df_save`: cleaned DataFrame written to `datasets/site_1_cleaned.csv`.

## Execution
1. Ensure `datasets/site_1/` contains the CSV files named in the notebook.
2. Run via Jupyter: `jupyter nbconvert --execute 01_EDA.ipynb`.
3. The notebook prints dataset summaries, shows the plots listed above, and writes `datasets/site_1_cleaned.csv`.

## Suggestions
- Parameterize data paths/dates so the same notebook can be rerun for other sites or periods without edits.
- Persist the textual summary/statistics in JSON or Markdown (automated at runtime) so downstream reports can ingest them programmatically.
