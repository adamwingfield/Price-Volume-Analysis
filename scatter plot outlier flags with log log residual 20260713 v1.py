from dataclasses import dataclass, field
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D


REQUIRED_COLUMNS = ["Procedure Units", "Procedure Price"]
VALID_FIRST_COLUMN_NAMES = ["Customer Group", "Customer Group L3", "Customer"]


@dataclass(frozen=True)
class PlotConfig:
    """
    Store application configuration for peer grouping, plotting,
    and residual-based outlier detection.
    """

    peer_group_bins: list = field(
        default_factory=lambda: [10, 100, 200, 500, float("inf")]
    )
    peer_group_labels: list = field(
        default_factory=lambda: ["10-100", "101-200", "201-500", "501+"]
    )
    peer_group_colors: dict = field(
        default_factory=lambda: {
            "10-100": "red",
            "101-200": "green",
            "201-500": "blue",
            "501+": "orange",
        }
    )
    robust_z_threshold: float = 3.5
    lower_percentile: int = 5
    upper_percentile: int = 95
    min_procedure_units: int = 10
    max_procedure_price: int = 4000
    figure_size: tuple = (13, 8)
    scatter_size: int = 80
    scatter_alpha: float = 0.6
    flagged_scatter_size: int = 180
    export_columns: list = field(
        default_factory=lambda: [
            "Customer Group",
            "Procedure Units",
            "Procedure Price",
            "Peer Group",
            "Log Procedure Units",
            "Log Procedure Price",
            "Expected Log Price",
            "Expected Price",
            "Residual",
            "Log Residual",
            "Residual Robust Z-Score",
            "Is Residual Robust Z Outlier",
            "Is Residual IQR Outlier",
            "Is Residual Percentile Extreme",
        ]
    )


CONFIG = PlotConfig()


def select_input_file():
    """
    Open a file dialog and return the selected file path.

    Returns:
        str | None: Selected file path, or None if no file was chosen.
    """
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select CSV or Excel file",
        filetypes=[
            ("Data files", "*.csv *.xlsx"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx"),
            ("All files", "*.*"),
        ],
    )

    if not file_path:
        messagebox.showwarning("No File Selected", "No file was selected.")
        return None

    return file_path


def read_input_file(file_path):
    """
    Read a CSV or Excel file into a DataFrame.

    Args:
        file_path (str): Path to the selected input file.

    Returns:
        pandas.DataFrame | None: Loaded DataFrame, or None on failure.
    """
    ext = os.path.splitext(file_path)[1].lower()

    try:
        if ext == ".csv":
            return pd.read_csv(file_path)
        if ext == ".xlsx":
            return pd.read_excel(file_path)

        messagebox.showerror(
            "Unsupported File Type",
            "Please select a .csv or .xlsx file.",
        )
        return None

    except FileNotFoundError:
        messagebox.showerror("File Error", "The selected file could not be found.")
    except pd.errors.EmptyDataError:
        messagebox.showerror("File Error", "The selected file is empty.")
    except pd.errors.ParserError:
        messagebox.showerror("File Error", "The file could not be parsed.")
    except ImportError:
        messagebox.showerror(
            "Missing Dependency",
            "Excel file support requires the 'openpyxl' package.\n\n"
            "Install it with:\npip install openpyxl",
        )
    except Exception as error:
        messagebox.showerror(
            "File Error",
            f"An unexpected error occurred while reading the file:\n{error}",
        )

    return None


def validate_and_standardize_columns(df):
    """
    Validate the input file structure and standardize column names.

    Args:
        df (pandas.DataFrame): Raw input DataFrame.

    Returns:
        pandas.DataFrame | None: Standardized DataFrame, or None on failure.
    """
    if df.empty or len(df.columns) < 3:
        messagebox.showerror(
            "Invalid File Format",
            "The file must contain at least 3 columns:\n"
            "1. Customer Group or Customer\n"
            "2. Procedure Units\n"
            "3. Procedure Price",
        )
        return None

    first_col = df.columns[0]
    if first_col not in VALID_FIRST_COLUMN_NAMES:
        messagebox.showerror(
            "Invalid First Column",
            "The first column must be either 'Customer Group' or 'Customer'.\n"
            f"Found: '{first_col}'",
        )
        return None

    df = df.rename(columns={first_col: "Customer Group"})

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        messagebox.showerror(
            "Missing Columns",
            f"The file is missing required columns:\n{', '.join(missing_cols)}",
        )
        return None

    return df


def clean_data(df, config):
    """
    Clean and filter the dataset for analysis.

    Args:
        df (pandas.DataFrame): Standardized input DataFrame.
        config (PlotConfig): Application configuration.

    Returns:
        pandas.DataFrame | None: Cleaned DataFrame, or None if no valid data remains.
    """
    df = df[["Customer Group", "Procedure Units", "Procedure Price"]].copy()

    df["Procedure Units"] = pd.to_numeric(df["Procedure Units"], errors="coerce")
    df["Procedure Price"] = pd.to_numeric(df["Procedure Price"], errors="coerce")

    df["Customer Group"] = df["Customer Group"].astype(str).str.strip()
    df.loc[
        df["Customer Group"].isin(["", "nan", "None"]),
        "Customer Group",
    ] = "Unknown"

    invalid_numeric_rows = (
        df["Procedure Units"].isna() | df["Procedure Price"].isna()
    )
    dropped_count = invalid_numeric_rows.sum()
    df = df[~invalid_numeric_rows].copy()

    df = df[df["Procedure Units"] > 0].copy()
    df = df[df["Procedure Price"] > 0].copy()
    df = df[df["Procedure Units"] >= config.min_procedure_units].copy()
    df = df[df["Procedure Price"] <= config.max_procedure_price].copy()

    if df.empty:
        msg = "No valid data remains after cleaning and filtering."
        if dropped_count > 0:
            msg += f"\n\nRows dropped due to invalid numeric values: {dropped_count}"
        messagebox.showwarning("No Valid Data", msg)
        return None

    if dropped_count > 0:
        messagebox.showinfo(
            "Data Cleaning Complete",
            "Loaded successfully.\n"
            f"Rows dropped due to invalid numeric values: {dropped_count}",
        )

    return df


def load_and_clean_file(config):
    """
    Select, load, validate, and clean an input data file.

    Args:
        config (PlotConfig): Application configuration.

    Returns:
        pandas.DataFrame | None: Cleaned DataFrame, or None on failure.
    """
    file_path = select_input_file()
    if not file_path:
        return None

    df = read_input_file(file_path)
    if df is None:
        return None

    df = validate_and_standardize_columns(df)
    if df is None:
        return None

    return clean_data(df, config)


def assign_peer_groups(df, config):
    """
    Assign peer group labels based on procedure volume.

    Args:
        df (pandas.DataFrame): Cleaned input DataFrame.
        config (PlotConfig): Application configuration.

    Returns:
        pandas.DataFrame | None: DataFrame with peer groups, or None if empty.
    """
    df = df.copy()
    df["Peer Group"] = pd.cut(
        df["Procedure Units"],
        bins=config.peer_group_bins,
        labels=config.peer_group_labels,
        include_lowest=True,
        right=True,
    )

    df = df.dropna(subset=["Peer Group"]).copy()

    if df.empty:
        messagebox.showwarning(
            "No Valid Data",
            "No records remain after assigning peer groups.",
        )
        return None

    return df


def calculate_expected_price_and_residuals(group):
    """
    Model log(procedure price) as a linear function of log(procedure units)
    within a peer group, then calculate expected price and residuals.

    Residual is defined on the original price scale as:
        actual price - expected price

    Log residual is defined as:
        log(actual price) - expected log(price)

    Args:
        group (pandas.DataFrame): Peer group subset.

    Returns:
        pandas.DataFrame: Group with modeled price and residual columns added.
    """
    group = group.copy()

    group["Log Procedure Units"] = np.log(group["Procedure Units"])
    group["Log Procedure Price"] = np.log(group["Procedure Price"])

    if len(group) < 2 or group["Log Procedure Units"].nunique() < 2:
        expected_log_price = group["Log Procedure Price"].median()
        group["Expected Log Price"] = expected_log_price
        group["Expected Price"] = np.exp(group["Expected Log Price"])
        group["Log Residual"] = (
            group["Log Procedure Price"] - group["Expected Log Price"]
        )
        group["Residual"] = group["Procedure Price"] - group["Expected Price"]
        return group

    slope, intercept = np.polyfit(
        group["Log Procedure Units"],
        group["Log Procedure Price"],
        1,
    )

    group["Expected Log Price"] = (
        intercept + slope * group["Log Procedure Units"]
    )
    group["Expected Price"] = np.exp(group["Expected Log Price"])
    group["Log Residual"] = (
        group["Log Procedure Price"] - group["Expected Log Price"]
    )
    group["Residual"] = group["Procedure Price"] - group["Expected Price"]

    return group


def calculate_robust_z_scores(series):
    """
    Calculate robust Z-scores using the median and median absolute deviation.

    Robust Z-score formula:
        0.6745 * (x - median) / MAD

    Args:
        series (pandas.Series): Numeric series to evaluate.

    Returns:
        pandas.Series: Robust Z-scores for each value.
    """
    median_value = series.median()
    absolute_deviation = (series - median_value).abs()
    mad = absolute_deviation.median()

    if pd.isna(mad) or mad == 0:
        return pd.Series(0.0, index=series.index)

    return 0.6745 * (series - median_value) / mad


def calculate_outlier_flags(group, config):
    """
    Calculate outlier flags based on log residuals using robust Z-score, IQR,
    and percentile-based methods.

    Args:
        group (pandas.DataFrame): Peer group subset.
        config (PlotConfig): Application configuration.

    Returns:
        pandas.DataFrame: Peer group with residual-based outlier flags added.
    """
    group = calculate_expected_price_and_residuals(group)

    group["Residual Robust Z-Score"] = calculate_robust_z_scores(
        group["Log Residual"]
    )
    group["Is Residual Robust Z Outlier"] = (
        group["Residual Robust Z-Score"].abs() >= config.robust_z_threshold
    )

    q1 = group["Log Residual"].quantile(0.25)
    q3 = group["Log Residual"].quantile(0.75)
    iqr = q3 - q1

    if pd.notna(iqr):
        lower_iqr = q1 - 1.5 * iqr
        upper_iqr = q3 + 1.5 * iqr
        group["Is Residual IQR Outlier"] = (
            (group["Log Residual"] < lower_iqr)
            | (group["Log Residual"] > upper_iqr)
        )
    else:
        group["Is Residual IQR Outlier"] = False

    p_low = group["Log Residual"].quantile(config.lower_percentile / 100)
    p_high = group["Log Residual"].quantile(config.upper_percentile / 100)

    if pd.notna(p_low) and pd.notna(p_high):
        group["Is Residual Percentile Extreme"] = (
            (group["Log Residual"] <= p_low)
            | (group["Log Residual"] >= p_high)
        )
    else:
        group["Is Residual Percentile Extreme"] = False

    return group


def analyze_residuals(df, config):
    """
    Apply log-log residual modeling and outlier detection to each peer group.

    Args:
        df (pandas.DataFrame): Input DataFrame with peer groups assigned.
        config (PlotConfig): Application configuration.

    Returns:
        pandas.DataFrame: Combined analyzed DataFrame.
    """
    analyzed_groups = []

    for bin_label, group in df.groupby("Peer Group", observed=True):
        if group.empty or bin_label not in config.peer_group_colors:
            continue

        analyzed_group = calculate_outlier_flags(group, config)
        analyzed_groups.append(analyzed_group)

    if not analyzed_groups:
        return pd.DataFrame()

    return pd.concat(analyzed_groups, ignore_index=True)


def get_flagged_rows(group):
    """
    Return rows flagged by any residual-based outlier detection method.

    Args:
        group (pandas.DataFrame): Peer group with outlier flags.

    Returns:
        pandas.DataFrame: Flagged rows only.
    """
    return group[
        group["Is Residual Robust Z Outlier"]
        | group["Is Residual IQR Outlier"]
        | group["Is Residual Percentile Extreme"]
    ].copy()


def build_flag_text(row):
    """
    Build a label describing which outlier methods flagged a row.

    Args:
        row (pandas.Series): A flagged row.

    Returns:
        str: Comma-separated flag labels.
    """
    flag_types = []

    if row["Is Residual Robust Z Outlier"]:
        flag_types.append("Log-Residual Robust Z-score")
    if row["Is Residual IQR Outlier"]:
        flag_types.append("Log-Residual IQR")
    if row["Is Residual Percentile Extreme"]:
        flag_types.append("Log-Residual Percentile")

    return ", ".join(flag_types)


def plot_peer_group(ax, group, bin_label, config):
    """
    Plot the base scatter points and expected-price trend line for one peer group.

    Args:
        ax (matplotlib.axes.Axes): Plot axes.
        group (pandas.DataFrame): Peer group subset.
        bin_label (str): Peer group label.
        config (PlotConfig): Application configuration.
    """
    color = config.peer_group_colors[bin_label]

    ax.scatter(
        group["Procedure Units"],
        group["Procedure Price"],
        label=bin_label,
        color=color,
        s=config.scatter_size,
        alpha=config.scatter_alpha,
    )

    sorted_group = group.sort_values("Procedure Units")
    ax.plot(
        sorted_group["Procedure Units"],
        sorted_group["Expected Price"],
        color=color,
        linestyle="--",
        linewidth=2,
    )


def highlight_flagged_points(ax, flagged, config):
    """
    Highlight flagged points with a black outline.

    Args:
        ax (matplotlib.axes.Axes): Plot axes.
        flagged (pandas.DataFrame): Flagged rows.
        config (PlotConfig): Application configuration.
    """
    if flagged.empty:
        return

    ax.scatter(
        flagged["Procedure Units"],
        flagged["Procedure Price"],
        s=config.flagged_scatter_size,
        facecolors="none",
        edgecolors="black",
        linewidths=1.5,
        zorder=4,
    )


def annotate_flagged_customers(ax, flagged, start_counter, numbered_customers):
    """
    Annotate flagged customers on the plot and collect legend entries.

    Args:
        ax (matplotlib.axes.Axes): Plot axes.
        flagged (pandas.DataFrame): Flagged rows.
        start_counter (int): Starting annotation number.
        numbered_customers (list): List to store legend metadata.

    Returns:
        int: Updated counter after processing flagged rows.
    """
    counter = start_counter

    for _, row in flagged.iterrows():
        flag_text = build_flag_text(row)

        ax.annotate(
            str(counter),
            xy=(row["Procedure Units"], row["Procedure Price"]),
            xytext=(6, 6),
            textcoords="offset points",
            ha="left",
            va="bottom",
            fontsize=9,
            fontweight="bold",
            color="black",
            bbox=dict(
                boxstyle="circle,pad=0.2",
                fc="white",
                ec="black",
                alpha=0.9,
            ),
        )

        numbered_customers.append(
            (
                counter,
                row["Customer Group"],
                row["Peer Group"],
                flag_text,
                row["Residual"],
                row["Log Residual"],
            )
        )
        counter += 1

    return counter


def format_plot(ax):
    """
    Apply titles, labels, scale, and grid formatting to the plot.

    Args:
        ax (matplotlib.axes.Axes): Plot axes.
    """
    ax.set_xlabel("Annual Procedure Volume per Customer (Log Scale)", fontsize=15)
    ax.set_ylabel("Average Procedure Price per Customer (USD, Log Model)", fontsize=15)
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_title(
        "Customer Price vs. Volume with Log-Log Residual-Based Outlier Flags",
        fontsize=15,
    )
    plt.xticks(fontsize=14)
    plt.yticks(fontsize=14)
    ax.grid(True, linestyle="--", alpha=0.4)


def add_legends(ax, numbered_customers):
    """
    Add the peer group legend and flagged customer legend.

    Args:
        ax (matplotlib.axes.Axes): Plot axes.
        numbered_customers (list): Annotated flagged customer metadata.
    """
    bin_legend = ax.legend(title="Volume Size Bin", loc="upper left")
    ax.add_artist(bin_legend)

    customer_handles = [
        Line2D(
            [0],
            [0],
            marker="o",
            color="w",
            markerfacecolor="white",
            markeredgecolor="black",
            markersize=10,
            label=(
                f"{num}: {cust} ({peer}) - {flags}, "
                f"Residual={residual:.2f}, "
                f"Log Residual={log_residual:.3f}"
            ),
        )
        for num, cust, peer, flags, residual, log_residual in numbered_customers
    ]

    if customer_handles:
        ax.legend(
            handles=customer_handles,
            title="Flagged Customers",
            loc="upper right",
            fontsize=8,
            title_fontsize=9,
        )


def create_scatter_plot(df, config):
    """
    Create the scatter plot with peer groups, expected-price trend lines,
    residual-based outlier highlighting, and customer annotations.

    Args:
        df (pandas.DataFrame): Analyzed DataFrame.
        config (PlotConfig): Application configuration.
    """
    _, ax = plt.subplots(figsize=config.figure_size)
    numbered_customers = []
    customer_counter = 1

    for bin_label, group in df.groupby("Peer Group", observed=True):
        if group.empty or bin_label not in config.peer_group_colors:
            continue

        flagged = get_flagged_rows(group)

        plot_peer_group(ax, group, bin_label, config)
        highlight_flagged_points(ax, flagged, config)

        customer_counter = annotate_flagged_customers(
            ax,
            flagged,
            customer_counter,
            numbered_customers,
        )

    format_plot(ax)
    add_legends(ax, numbered_customers)
    plt.tight_layout()
    plt.show()


def build_export_table(df, config):
    """
    Build an export-ready table containing modeled price and residual metrics.

    Args:
        df (pandas.DataFrame): Analyzed DataFrame.
        config (PlotConfig): Application configuration.

    Returns:
        pandas.DataFrame: Export-ready residual analysis table.
    """
    export_df = df[config.export_columns].copy()
    export_df = export_df.sort_values(
        by=["Peer Group", "Log Residual"],
        ascending=[True, False],
    )
    return export_df


def export_results_table(export_df):
    """
    Prompt the user to save the residual analysis table to CSV or Excel.

    Args:
        export_df (pandas.DataFrame): Export-ready results table.
    """
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.asksaveasfilename(
        title="Save residual analysis table",
        defaultextension=".xlsx",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("CSV files", "*.csv"),
            ("All files", "*.*"),
        ],
        initialfile="customer_loglog_residual_analysis.xlsx",
    )

    if not file_path:
        return

    ext = os.path.splitext(file_path)[1].lower()

    try:
        if ext == ".csv":
            export_df.to_csv(file_path, index=False)
        elif ext == ".xlsx":
            export_df.to_excel(file_path, index=False)
        else:
            messagebox.showerror(
                "Unsupported File Type",
                "Please save the file as .csv or .xlsx.",
            )
            return

        messagebox.showinfo(
            "Export Complete",
            f"Residual analysis table saved successfully:\n{file_path}",
        )
    except ImportError:
        messagebox.showerror(
            "Missing Dependency",
            "Excel export requires the 'openpyxl' package.\n\n"
            "Install it with:\npip install openpyxl",
        )
    except Exception as error:
        messagebox.showerror(
            "Export Error",
            f"An unexpected error occurred while saving the file:\n{error}",
        )


def main():
    """
    Run the full workflow: load data, clean it, assign peer groups,
    model expected price using a log-log relationship, calculate residuals,
    create the plot, and export the results table.
    """
    config = CONFIG

    df = load_and_clean_file(config)
    if df is None:
        raise SystemExit

    df = assign_peer_groups(df, config)
    if df is None:
        raise SystemExit

    analyzed_df = analyze_residuals(df, config)
    if analyzed_df.empty:
        messagebox.showwarning(
            "No Analysis Results",
            "No data was available for residual analysis.",
        )
        raise SystemExit

    create_scatter_plot(analyzed_df, config)

    export_df = build_export_table(analyzed_df, config)
    export_results_table(export_df)


if __name__ == "__main__":
    main()
