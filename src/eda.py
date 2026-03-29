"""Data cleaning and Exploratory Data Analysis (EDA).

Loads raw CSVs by filename pattern, cleans/merges to 15-min intervals,
generates EDA plots, and saves cleaned CSV.
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns
from pathlib import Path

plt.style.use("seaborn-v0_8-whitegrid")
plt.rcParams["figure.figsize"] = (14, 5)
plt.rcParams["font.size"] = 12


# ---------------------------------------------------------------------------
# 1. Load raw data by filename pattern
# ---------------------------------------------------------------------------

def _load_csv_by_prefix(data_dir: Path, prefix: str, col_name: str) -> pd.DataFrame:
    """Find CSV matching prefix*, load first 2 columns only."""
    matches = sorted(data_dir.glob(f"{prefix}*.csv"))
    if not matches:
        raise FileNotFoundError(f"No CSV starting with '{prefix}' in {data_dir}")
    path = matches[0]
    df = pd.read_csv(path, parse_dates=[0], usecols=[0, 1])
    df.columns = ["datetime", col_name]
    return df, path.name


def load_raw_site(data_dir: str | Path) -> dict:
    """Load 4 raw CSVs from a site folder using filename patterns.

    Pattern rules:
        Temp*.csv       -> temperature_c
        Irradiance*.csv -> irradiance_wm2
        load*.csv       -> load_kw
        gen*.csv        -> generation_kw

    Returns dict of {col_name: (DataFrame, filename)}.
    """
    data_dir = Path(data_dir)
    raw = {}
    patterns = [
        ("Temp", "temperature_c"),
        ("Irradiance", "irradiance_wm2"),
        ("load", "load_kw"),
        ("gen", "generation_kw"),
    ]
    for prefix, col_name in patterns:
        df, fname = _load_csv_by_prefix(data_dir, prefix, col_name)
        raw[col_name] = (df, fname)
    return raw


# ---------------------------------------------------------------------------
# 2. Clean & merge
# ---------------------------------------------------------------------------

def clean_and_merge(raw: dict, resample: str = "15min",
                    ffill_limit: int = 4, interp_limit: int = 4) -> tuple:
    """Resample, merge, fill gaps, drop remaining NaN.

    Returns:
        df_merged: merged DataFrame BEFORE dropna (for before/after comparison)
        df_clean: final cleaned DataFrame
    """
    resampled = {}
    for col_name, (df, _) in raw.items():
        rs = df.set_index("datetime").resample(resample).mean()
        resampled[col_name] = rs

    df_merged = pd.concat(resampled.values(), axis=1, join="outer").sort_index()

    df_clean = df_merged.copy()
    df_clean = df_clean.ffill(limit=ffill_limit)
    df_clean = df_clean.interpolate(method="time", limit=interp_limit)
    df_clean = df_clean.dropna()

    return df_merged, df_clean


def detect_no_sun_generation(df: pd.DataFrame) -> pd.DataFrame:
    """Detect rows where generation > 0 but irradiance <= 0 (before cleaning)."""
    if "irradiance_wm2" in df.columns and "generation_kw" in df.columns:
        mask = (df["irradiance_wm2"] <= 0) & (df["generation_kw"] > 0)
        return df[mask].copy()
    return pd.DataFrame()


def zero_no_sun_generation(df: pd.DataFrame) -> pd.DataFrame:
    """Physical constraint: no sunlight = no solar generation."""
    df = df.copy()
    if "irradiance_wm2" in df.columns and "generation_kw" in df.columns:
        mask = (df["irradiance_wm2"] <= 0) & (df["generation_kw"] > 0)
        n_zeroed = mask.sum()
        df.loc[mask, "generation_kw"] = 0.0
        if n_zeroed > 0:
            print(f"  Zeroed {n_zeroed} generation readings with no sunlight "
                  f"(standby/sensor noise)")
    return df


# ---------------------------------------------------------------------------
# 3. Print EDA summary to terminal
# ---------------------------------------------------------------------------

def print_raw_summary(raw: dict):
    """Print summary of raw data."""
    print("\n" + "=" * 60)
    print("  Raw Data Summary")
    print("=" * 60)
    for col_name, (df, fname) in raw.items():
        n = len(df)
        nulls = df[col_name].isnull().sum()
        print(f"\n  {col_name} ({fname}):")
        print(f"    Rows: {n:,}")
        print(f"    Range: {df['datetime'].min()} to {df['datetime'].max()}")
        print(f"    Nulls: {nulls} ({nulls/n*100:.2f}%)")

        # Time interval analysis
        diffs = df["datetime"].diff().dropna()
        print(f"    Median interval: {diffs.median()}")
        print(f"    Max gap: {diffs.max()}")
        gaps = diffs[diffs > 2 * diffs.median()]
        print(f"    Gaps > 2x median: {len(gaps)}")


def print_merge_summary(df_merged: pd.DataFrame, df_clean: pd.DataFrame):
    """Print before/after cleaning summary."""
    print("\n" + "=" * 60)
    print("  Merge & Cleaning Summary")
    print("=" * 60)

    print(f"\n  After merge (before cleaning): {df_merged.shape[0]:,} rows")
    print("  Nulls per column:")
    for col in df_merged.columns:
        n = df_merged[col].isnull().sum()
        pct = n / len(df_merged) * 100
        print(f"    {col:20s}: {n:6d} ({pct:.2f}%)")

    print(f"\n  After cleaning: {df_clean.shape[0]:,} rows")
    print(f"  Dropped: {len(df_merged) - len(df_clean):,} rows")
    print("  Remaining nulls per column:")
    for col in df_clean.columns:
        n = df_clean[col].isnull().sum()
        print(f"    {col:20s}: {n}")


def print_statistics(df_clean: pd.DataFrame):
    """Print basic statistics."""
    print("\n" + "=" * 60)
    print("  Basic Statistics")
    print("=" * 60)
    print(df_clean.describe().round(2).to_string())


def print_correlation(df_clean: pd.DataFrame):
    """Print correlation with generation."""
    print("\n" + "=" * 60)
    print("  Correlation with generation_kw")
    print("=" * 60)
    if "generation_kw" in df_clean.columns:
        corr = df_clean.corr()["generation_kw"].sort_values(ascending=False)
        for k, v in corr.items():
            print(f"    {k:20s}: {v:.4f}")


def detect_anomaly_categories(df: pd.DataFrame, gen_no_sun: pd.DataFrame) -> dict:
    """Detect 5 categories of anomalies for EDA visualization.

    Categories:
        1. High irradiance, low generation - potential panel degradation/shading
        2. Sudden generation drop - abrupt drop between consecutive timesteps
        3. Gradual efficiency decline - rolling efficiency trend drops
        4. Unusual generation spikes - generation far above normal for given irradiance
        5. Generation with zero irradiance - sensor/logging error (pre-cleaning snapshot)

    Returns dict of {category_name: DataFrame}.
    """
    has_irr = "irradiance_wm2" in df.columns
    has_gen = "generation_kw" in df.columns
    anomalies = {}

    if not (has_irr and has_gen):
        return anomalies

    # 1. High irradiance but low generation (daytime context)
    #    Use daytime-only percentiles so the threshold is meaningful
    daytime_all = df[df["irradiance_wm2"] > 50]
    if len(daytime_all) > 0:
        irr_q75 = daytime_all["irradiance_wm2"].quantile(0.75)
        gen_q10 = daytime_all["generation_kw"].quantile(0.10)
        anomalies["High Irr / Low Gen"] = df[
            (df["irradiance_wm2"] > irr_q75) &
            (df["generation_kw"] < gen_q10)
        ]
    else:
        anomalies["High Irr / Low Gen"] = pd.DataFrame()

    # 2. Sudden generation drop (consecutive daytime steps only)
    #    Large negative change between 15-min consecutive intervals
    daytime = df[df["irradiance_wm2"] > 50].copy()
    if len(daytime) > 4:
        time_diff = daytime.index.to_series().diff()
        consecutive = time_diff <= pd.Timedelta(minutes=20)
        gen_diff = daytime["generation_kw"].diff()
        # Drop > 50% of current generation in one step
        pct_drop = gen_diff / daytime["generation_kw"].shift(1)
        drop_mask = consecutive & (pct_drop < -0.5) & (gen_diff.abs() > 50)
        anomalies["Sudden Drop"] = daytime[drop_mask.fillna(False)]
    else:
        anomalies["Sudden Drop"] = pd.DataFrame()

    # 3. Gradual efficiency decline
    #    Daytime efficiency (gen/irr) drops below rolling mean - 2*std
    daytime_eff = df[df["irradiance_wm2"] > 50].copy()
    if len(daytime_eff) > 20:
        daytime_eff["efficiency"] = (
            daytime_eff["generation_kw"] / daytime_eff["irradiance_wm2"]
        )
        roll_mean = daytime_eff["efficiency"].rolling(96, min_periods=20).mean()
        roll_std = daytime_eff["efficiency"].rolling(96, min_periods=20).std()
        decline_mask = daytime_eff["efficiency"] < (roll_mean - 2 * roll_std)
        anomalies["Efficiency Decline"] = daytime_eff[decline_mask.fillna(False)]
    else:
        anomalies["Efficiency Decline"] = pd.DataFrame()

    # 4. Unusual generation spikes
    #    Generation far above expected for given irradiance bin
    daytime_sp = df[df["irradiance_wm2"] > 50].copy()
    if len(daytime_sp) > 20:
        daytime_sp["irr_bin"] = pd.qcut(
            daytime_sp["irradiance_wm2"], q=10, duplicates="drop"
        )
        bin_stats = daytime_sp.groupby("irr_bin", observed=True)["generation_kw"].agg(
            ["mean", "std"]
        )
        daytime_sp = daytime_sp.join(bin_stats, on="irr_bin", rsuffix="_bin")
        spike_mask = daytime_sp["generation_kw"] > (
            daytime_sp["mean"] + 3 * daytime_sp["std"]
        )
        anomalies["Gen Spike"] = daytime_sp[spike_mask.fillna(False)]
    else:
        anomalies["Gen Spike"] = pd.DataFrame()

    # 5. Generation with zero irradiance (pre-cleaning snapshot)
    anomalies["Gen / Zero Irr"] = gen_no_sun

    return anomalies


def print_anomaly_indicators(anomalies: dict):
    """Print anomaly category counts."""
    print("\n" + "=" * 60)
    print("  Anomaly Indicators (Visual Inspection)")
    print("=" * 60)
    for name, adf in anomalies.items():
        print(f"    {name:30s}: {len(adf):,} samples")


def print_net_power(df_clean: pd.DataFrame):
    """Print net power summary."""
    if "generation_kw" in df_clean.columns and "load_kw" in df_clean.columns:
        net = df_clean["generation_kw"] - df_clean["load_kw"]
        surplus_pct = (net >= 0).sum() / len(net) * 100
        print("\n" + "=" * 60)
        print("  Net Power (Generation - Load)")
        print("=" * 60)
        print(f"  Surplus periods: {surplus_pct:.1f}%")
        print(f"  Deficit periods: {100 - surplus_pct:.1f}%")


# ---------------------------------------------------------------------------
# 4. EDA Plots
# ---------------------------------------------------------------------------

def plot_missing_data(df_merged: pd.DataFrame, save_path: Path):
    """Heatmap of missing data pattern."""
    fig, ax = plt.subplots(figsize=(14, 3))
    sns.heatmap(df_merged.isnull().T, cbar=False, yticklabels=True, cmap="Reds", ax=ax)
    ax.set_title("Missing Data Pattern (red = missing)")
    ax.set_xlabel("Time Index")
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_before_after_cleaning(df_merged: pd.DataFrame, df_clean: pd.DataFrame,
                                save_path: Path):
    """Compare before (with gaps) and after cleaning for each variable."""
    cols = [c for c in df_merged.columns if c in df_clean.columns]
    n_cols = len(cols)
    fig, axes = plt.subplots(n_cols, 2, figsize=(18, 4 * n_cols))
    if n_cols == 1:
        axes = axes.reshape(1, -1)

    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}

    for i, col in enumerate(cols):
        color = colors.get(col, "gray")

        # Before
        axes[i, 0].plot(df_merged.index, df_merged[col], color=color, lw=0.5)
        n_null = df_merged[col].isnull().sum()
        axes[i, 0].set_title(f"BEFORE - {col} (nulls: {n_null})")
        axes[i, 0].set_ylabel(col)

        # After
        axes[i, 1].plot(df_clean.index, df_clean[col], color=color, lw=0.5)
        axes[i, 1].set_title(f"AFTER - {col} (rows: {len(df_clean)})")
        axes[i, 1].set_ylabel(col)

    plt.suptitle("Before vs After Cleaning", fontsize=14, y=1.01)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_time_series(df_clean: pd.DataFrame, save_path: Path):
    """4-panel time series of all variables."""
    cols = ["generation_kw", "irradiance_wm2", "temperature_c", "load_kw"]
    cols = [c for c in cols if c in df_clean.columns]
    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}
    titles = {"generation_kw": "Solar Generation (kW)",
              "irradiance_wm2": "Solar Irradiance (W/m²)",
              "temperature_c": "Ambient Temperature (°C)",
              "load_kw": "Power Consumption (kW)"}

    fig, axes = plt.subplots(len(cols), 1, figsize=(16, 3.5 * len(cols)), sharex=True)
    if len(cols) == 1:
        axes = [axes]

    for ax, col in zip(axes, cols):
        ax.plot(df_clean.index, df_clean[col], color=colors.get(col, "gray"), lw=0.5)
        ax.set_ylabel(col)
        ax.set_title(titles.get(col, col))

    axes[-1].xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
    plt.xticks(rotation=45)
    plt.suptitle("All Variables Over Time", fontsize=14, y=1.01)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_one_week(df_clean: pd.DataFrame, save_path: Path):
    """Zoom into one week of data."""
    # Find a good week with data
    if len(df_clean) < 7 * 96:  # less than 7 days at 15-min
        return

    mid = df_clean.index[len(df_clean) // 4]  # early part of data
    start = mid.normalize()
    end = start + pd.Timedelta(days=7)
    week = df_clean[start:end]
    if len(week) < 10:
        return

    cols = ["generation_kw", "irradiance_wm2", "temperature_c", "load_kw"]
    cols = [c for c in cols if c in df_clean.columns]
    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}

    fig, axes = plt.subplots(len(cols), 1, figsize=(16, 3 * len(cols)), sharex=True)
    if len(cols) == 1:
        axes = [axes]
    for ax, col in zip(axes, cols):
        ax.plot(week.index, week[col], color=colors.get(col, "gray"), lw=1)
        ax.set_ylabel(col)

    plt.suptitle(f"One Week Detail ({start.date()} to {end.date()})", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_distributions(df_clean: pd.DataFrame, save_path: Path):
    """Histogram of each variable."""
    cols = ["generation_kw", "irradiance_wm2", "temperature_c", "load_kw"]
    cols = [c for c in cols if c in df_clean.columns]
    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}

    n = len(cols)
    rows = (n + 1) // 2
    fig, axes = plt.subplots(rows, 2, figsize=(14, 5 * rows))
    axes = axes.flat

    for ax, col in zip(axes, cols):
        ax.hist(df_clean[col], bins=80, color=colors.get(col, "gray"),
                alpha=0.7, edgecolor="black", linewidth=0.3)
        ax.set_title(col)
        ax.set_ylabel("Count")
        ax.axvline(df_clean[col].mean(), color="black", ls="--",
                    label=f"Mean={df_clean[col].mean():.1f}")
        ax.legend()

    # Hide unused axes
    for i in range(len(cols), len(axes)):
        axes[i].set_visible(False)

    plt.suptitle("Distribution of Variables", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_boxplots(df_clean: pd.DataFrame, save_path: Path):
    """Box plots for outlier detection."""
    cols = ["generation_kw", "irradiance_wm2", "temperature_c", "load_kw"]
    cols = [c for c in cols if c in df_clean.columns]
    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}

    fig, axes = plt.subplots(1, len(cols), figsize=(4 * len(cols), 5))
    if len(cols) == 1:
        axes = [axes]
    for ax, col in zip(axes, cols):
        ax.boxplot(df_clean[col].dropna(), patch_artist=True,
                   boxprops=dict(facecolor=colors.get(col, "gray"), alpha=0.6))
        ax.set_title(col)

    plt.suptitle("Box Plots - Outlier Detection", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_correlation(df_clean: pd.DataFrame, save_path: Path):
    """Correlation heatmap."""
    corr = df_clean.corr()
    fig, ax = plt.subplots(figsize=(8, 6))
    sns.heatmap(corr, annot=True, cmap="coolwarm", center=0, fmt=".3f",
                square=True, linewidths=1, ax=ax)
    ax.set_title("Correlation Matrix")
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_scatter_vs_generation(df_clean: pd.DataFrame, save_path: Path):
    """Scatter plots of generation vs each feature."""
    targets = ["irradiance_wm2", "temperature_c", "load_kw"]
    targets = [c for c in targets if c in df_clean.columns]
    if "generation_kw" not in df_clean.columns or not targets:
        return

    colors = {"irradiance_wm2": "orange", "temperature_c": "red", "load_kw": "blue"}
    labels = {"irradiance_wm2": "Irradiance (W/m²)",
              "temperature_c": "Temperature (°C)", "load_kw": "Load (kW)"}

    fig, axes = plt.subplots(1, len(targets), figsize=(6 * len(targets), 5))
    if len(targets) == 1:
        axes = [axes]

    for ax, col in zip(axes, targets):
        ax.scatter(df_clean[col], df_clean["generation_kw"],
                   alpha=0.1, s=1, color=colors.get(col, "gray"))
        ax.set_xlabel(labels.get(col, col))
        ax.set_ylabel("Generation (kW)")
        ax.set_title(f"Generation vs {labels.get(col, col)}")

    plt.suptitle("Generation vs Features", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_hourly_patterns(df_clean: pd.DataFrame, save_path: Path):
    """Hourly average patterns with std band."""
    df = df_clean.copy()
    df["hour"] = df.index.hour

    cols = ["generation_kw", "irradiance_wm2", "temperature_c", "load_kw"]
    cols = [c for c in cols if c in df.columns]
    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}

    n = len(cols)
    rows = (n + 1) // 2
    fig, axes = plt.subplots(rows, 2, figsize=(14, 5 * rows))
    axes = axes.flat

    for ax, col in zip(axes, cols):
        hourly = df.groupby("hour")[col].agg(["mean", "std"])
        ax.plot(hourly.index, hourly["mean"], color=colors.get(col, "gray"), lw=2)
        ax.fill_between(hourly.index, hourly["mean"] - hourly["std"],
                        hourly["mean"] + hourly["std"],
                        alpha=0.2, color=colors.get(col, "gray"))
        ax.set_xlabel("Hour of Day")
        ax.set_ylabel(col)
        ax.set_title(f"Average {col} by Hour")
        ax.set_xticks(range(0, 24))

    for i in range(len(cols), len(axes)):
        axes[i].set_visible(False)

    plt.suptitle("Hourly Patterns (mean ± 1 std)", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_monthly_trends(df_clean: pd.DataFrame, save_path: Path):
    """Monthly bar charts."""
    df = df_clean.copy()
    df["month"] = df.index.month

    cols = ["generation_kw", "irradiance_wm2", "temperature_c", "load_kw"]
    cols = [c for c in cols if c in df.columns]
    colors = {"generation_kw": "orange", "irradiance_wm2": "gold",
              "temperature_c": "red", "load_kw": "blue"}

    n = len(cols)
    rows = (n + 1) // 2
    fig, axes = plt.subplots(rows, 2, figsize=(14, 5 * rows))
    axes = axes.flat

    for ax, col in zip(axes, cols):
        monthly = df.groupby("month")[col].agg(["mean", "std"])
        ax.bar(monthly.index, monthly["mean"], color=colors.get(col, "gray"),
               alpha=0.7, yerr=monthly["std"], capsize=3)
        ax.set_xlabel("Month")
        ax.set_ylabel(col)
        ax.set_title(f"Monthly Average {col}")

    for i in range(len(cols), len(axes)):
        axes[i].set_visible(False)

    plt.suptitle("Monthly Trends", fontsize=14)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_efficiency(df_clean: pd.DataFrame, save_path: Path):
    """Generation efficiency analysis (daytime only)."""
    if "irradiance_wm2" not in df_clean.columns or "generation_kw" not in df_clean.columns:
        return

    daytime = df_clean[df_clean["irradiance_wm2"] > 10].copy()
    if len(daytime) == 0:
        return

    daytime["efficiency"] = daytime["generation_kw"] / daytime["irradiance_wm2"]

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))

    axes[0].hist(daytime["efficiency"], bins=80, color="green", alpha=0.7,
                 edgecolor="black", linewidth=0.3)
    axes[0].set_title("Generation Efficiency (Gen/Irradiance) during daytime")
    axes[0].set_xlabel("Efficiency (kW per W/m²)")

    axes[1].scatter(daytime["temperature_c"], daytime["efficiency"],
                    alpha=0.1, s=1, color="green")
    axes[1].set_title("Efficiency vs Temperature")
    axes[1].set_xlabel("Temperature (°C)")
    axes[1].set_ylabel("Efficiency")

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()

    print(f"\n  Daytime samples (irradiance > 10): {len(daytime)}")
    print(f"  Efficiency stats:")
    print(f"    Mean:   {daytime['efficiency'].mean():.4f}")
    print(f"    Std:    {daytime['efficiency'].std():.4f}")
    print(f"    Median: {daytime['efficiency'].median():.4f}")


def plot_anomaly_indicators(df_clean: pd.DataFrame, anomalies: dict,
                             save_path: Path):
    """Scatter plot highlighting 5 anomaly categories with distinct colors."""
    if "irradiance_wm2" not in df_clean.columns or "generation_kw" not in df_clean.columns:
        return

    CATEGORY_STYLE = {
        "High Irr / Low Gen":   {"color": "red",       "marker": "o"},
        "Sudden Drop":          {"color": "magenta",    "marker": "v"},
        "Efficiency Decline":   {"color": "orange",     "marker": "s"},
        "Gen Spike":            {"color": "lime",       "marker": "^"},
        "Gen / Zero Irr":       {"color": "dodgerblue", "marker": "D"},
    }

    fig, ax = plt.subplots(figsize=(16, 7))

    # Background: all normal points
    ax.scatter(df_clean["irradiance_wm2"], df_clean["generation_kw"],
               alpha=0.05, s=1, color="gray", label="Normal")

    # Overlay each anomaly category
    for name, adf in anomalies.items():
        if len(adf) == 0:
            continue
        style = CATEGORY_STYLE.get(name, {"color": "black", "marker": "x"})
        ax.scatter(
            adf["irradiance_wm2"], adf["generation_kw"],
            alpha=0.6, s=18, edgecolors="black", linewidths=0.3,
            color=style["color"], marker=style["marker"],
            label=f"{name} ({len(adf):,})",
        )

    ax.set_xlabel("Irradiance (W/m²)")
    ax.set_ylabel("Generation (kW)")
    ax.set_title("Potential Anomaly Points by Category")
    ax.legend(loc="upper left", fontsize=9, framealpha=0.9)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


def plot_net_power(df_clean: pd.DataFrame, save_path: Path):
    """Net power (generation - load) with surplus/deficit."""
    if "generation_kw" not in df_clean.columns or "load_kw" not in df_clean.columns:
        return

    net = df_clean["generation_kw"] - df_clean["load_kw"]

    fig, ax = plt.subplots(figsize=(16, 5))
    ax.fill_between(df_clean.index, net, 0,
                    where=net >= 0, color="green", alpha=0.5, label="Surplus")
    ax.fill_between(df_clean.index, net, 0,
                    where=net < 0, color="red", alpha=0.5, label="Deficit")
    ax.set_ylabel("Net Power (kW)")
    ax.set_title("Net Power (Generation - Load): Surplus vs Deficit")
    ax.legend()
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()


# ---------------------------------------------------------------------------
# 5. Main entry point
# ---------------------------------------------------------------------------

def run_eda(data_dir: str, output_base: str = "outputs") -> tuple:
    """Run full EDA pipeline: load, clean, print, plot, save.

    Args:
        data_dir: path to site folder (e.g., "datasets/site_1")
        output_base: base output directory (e.g., "outputs")

    Returns:
        (df_clean, cleaned_csv_path)
    """
    data_dir = Path(data_dir)
    site_name = data_dir.name  # e.g., "site_1"

    # Output directories
    eda_dir = Path(output_base) / "data_EDA" / site_name
    eda_dir.mkdir(parents=True, exist_ok=True)

    # Cleaned CSV path: datasets/site_1_cleaned.csv
    cleaned_csv = data_dir.parent / f"{site_name}_cleaned.csv"

    print("\n" + "=" * 60)
    print(f"  EDA Pipeline - {site_name}")
    print("=" * 60)
    print(f"  Data dir:    {data_dir}")
    print(f"  EDA output:  {eda_dir}")
    print(f"  Cleaned CSV: {cleaned_csv}")

    # --- Load raw ---
    raw = load_raw_site(data_dir)
    print_raw_summary(raw)

    # --- Clean & merge ---
    df_merged, df_clean = clean_and_merge(raw)
    print_merge_summary(df_merged, df_clean)

    # --- Detect anomalies BEFORE zeroing (for visualization) ---
    gen_no_sun = detect_no_sun_generation(df_clean)

    # --- Zero out no-sun generation (physical constraint) ---
    df_clean = zero_no_sun_generation(df_clean)

    # --- Statistics ---
    print_statistics(df_clean)
    print_correlation(df_clean)

    # --- Anomaly indicators (5 categories) ---
    anomalies = detect_anomaly_categories(df_clean, gen_no_sun)
    print_anomaly_indicators(anomalies)
    print_net_power(df_clean)

    # --- Plots ---
    print(f"\n  Generating EDA plots -> {eda_dir}")

    plot_missing_data(df_merged, eda_dir / "missing_data_pattern.png")
    print("    saved: missing_data_pattern.png")

    plot_before_after_cleaning(df_merged, df_clean, eda_dir / "before_after_cleaning.png")
    print("    saved: before_after_cleaning.png")

    plot_time_series(df_clean, eda_dir / "time_series.png")
    print("    saved: time_series.png")

    plot_one_week(df_clean, eda_dir / "one_week_detail.png")
    print("    saved: one_week_detail.png")

    plot_distributions(df_clean, eda_dir / "distributions.png")
    print("    saved: distributions.png")

    plot_boxplots(df_clean, eda_dir / "boxplots.png")
    print("    saved: boxplots.png")

    plot_correlation(df_clean, eda_dir / "correlation_matrix.png")
    print("    saved: correlation_matrix.png")

    plot_scatter_vs_generation(df_clean, eda_dir / "scatter_vs_generation.png")
    print("    saved: scatter_vs_generation.png")

    plot_hourly_patterns(df_clean, eda_dir / "hourly_patterns.png")
    print("    saved: hourly_patterns.png")

    plot_monthly_trends(df_clean, eda_dir / "monthly_trends.png")
    print("    saved: monthly_trends.png")

    plot_efficiency(df_clean, eda_dir / "efficiency_analysis.png")
    print("    saved: efficiency_analysis.png")

    plot_anomaly_indicators(df_clean, anomalies,
                            eda_dir / "anomaly_indicators.png")
    print("    saved: anomaly_indicators.png")

    plot_net_power(df_clean, eda_dir / "net_power.png")
    print("    saved: net_power.png")

    # --- Save cleaned CSV ---
    df_clean.to_csv(cleaned_csv)
    print(f"\n  Saved cleaned data: {df_clean.shape} -> {cleaned_csv}")
    print(f"  Columns: {list(df_clean.columns)}")

    return df_clean, cleaned_csv
