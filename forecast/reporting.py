"""
Reporting and visualization for CLO cashflow forecasts.

Adds the functionality from CLO_FHLB_CashflowForecast.ipynb and
CLO_Derivaties_CashflowForecast.ipynb:
  - Balance scenario chart (line plot across all scenarios)
  - Quarterly balance roll-up (snap to quarter-end dates)
  - Hedging output (first 2 payment dates and balances per CUSIP)
  - Read and parse multi-scenario cashflow summary workbooks
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Parse a multi-scenario cashflow summary workbook
# ---------------------------------------------------------------------------

def read_cashflow_summary(
    cashflow_path: str | Path,
    scenario_names: Optional[list[str]] = None,
) -> dict[str, pd.DataFrame]:
    """Read a cashflow summary workbook with one sheet per scenario.

    Each sheet has: Date, then groups of (Interest, Principal, Balance) per CUSIP,
    with CUSIPs in row 1 and headers (Interest/Principal/Balance) in row 3.

    Args:
        cashflow_path: Path to the cashflow summary Excel file (the output of
            Step 7 or the existing 'CLO Balance Forecast' workbooks).
        scenario_names: Optional list of sheet names to read. If None, reads all
            non-"Summary"/"Balance Summary" sheets.

    Returns:
        Dict mapping scenario name -> DataFrame with Date index and per-CUSIP
        Interest/Principal/Balance columns.
    """
    cashflow_path = Path(cashflow_path)
    xls = pd.ExcelFile(cashflow_path, engine="openpyxl")

    skip_sheets = {"Summary", "Balance Summary", "Sheet1"}
    if scenario_names is None:
        scenario_names = [s for s in xls.sheet_names if s not in skip_sheets]

    results = {}
    for name in scenario_names:
        if name not in xls.sheet_names:
            continue

        # Try reading with header row 2 (0-indexed), which is the format
        # from the existing CLO Balance Forecast workbooks.
        df = pd.read_excel(xls, sheet_name=name, header=2)
        df = df.fillna(0.0)

        # Rename the first unnamed column to "Date".
        first_col = df.columns[0]
        if "unnamed" in str(first_col).lower() or first_col == "Date":
            df = df.rename(columns={first_col: "Date"})

        results[name] = df

    return results


def extract_cusip_names(
    cashflow_path: str | Path,
    sheet_name: str = "Scenario 1",
) -> list[str]:
    """Extract CUSIP identifiers from row 1 of a cashflow summary sheet.

    The first row of each scenario sheet contains CUSIPs above each group
    of Interest/Principal/Balance columns.  CUSIPs of 'XXXXXXXXX' or blank
    are replaced by the column header.
    """
    df = pd.read_excel(cashflow_path, sheet_name=sheet_name, nrows=1, header=None)
    df = df.dropna(axis="columns")
    tranches = df.T
    tranches.columns = ["cusip"]
    tranches["cusip"] = tranches["cusip"].astype(str)

    names = []
    for idx, row in tranches.iterrows():
        cusip = row["cusip"]
        if cusip in ("XXXXXXXXX", "", "nan"):
            names.append(str(idx))
        else:
            names.append(cusip)

    return names[1:]  # Skip the first (Date column).


# ---------------------------------------------------------------------------
# Aggregate balances across scenarios
# ---------------------------------------------------------------------------

def build_balance_summary(
    scenario_data: dict[str, pd.DataFrame],
    cusip_names: Optional[list[str]] = None,
) -> pd.DataFrame:
    """Build a portfolio-level balance summary across all scenarios.

    Args:
        scenario_data: Dict from read_cashflow_summary().
        cusip_names: Optional CUSIP names for column renaming.

    Returns:
        DataFrame with Date index and one column per scenario (total balance).
    """
    # Find the scenario with the most rows (longest cashflow horizon).
    max_len = 0
    max_scenario = None
    for name, df in scenario_data.items():
        if len(df) > max_len:
            max_len = len(df)
            max_scenario = name

    index_values = scenario_data[max_scenario]["Date"].values
    balance_arrays = {}

    for name, df in scenario_data.items():
        balance_cols = [c for c in df.columns if "Balance" in str(c)]
        if balance_cols and balance_cols[0] in ("Balance", "Portfolio Balance"):
            balance_cols = balance_cols[1:]  # Skip portfolio-level if present

        if not balance_cols:
            continue

        total_balance = df[balance_cols].sum(axis=1).values

        # Pad shorter scenarios to match the longest.
        while len(total_balance) < max_len:
            total_balance = np.append(total_balance, total_balance[-1])

        balance_arrays[name] = total_balance

    return pd.DataFrame(balance_arrays, index=index_values)


# ---------------------------------------------------------------------------
# Quarterly roll-up
# ---------------------------------------------------------------------------

def quarterly_rollup(
    scenario_data: dict[str, pd.DataFrame],
    cusip_names: list[str],
    scenario_name: str = "Scenario 5",
) -> pd.DataFrame:
    """Snap per-CUSIP balances to quarter-end dates.

    Takes a single scenario's cashflow data and returns balances at
    each quarter-end (the last cashflow date before each quarter boundary).

    Args:
        scenario_data: Dict from read_cashflow_summary().
        cusip_names: CUSIP identifiers for column names.
        scenario_name: Which scenario to use (default "Scenario 5" = Flat).

    Returns:
        DataFrame with quarter-end dates and per-CUSIP balances.
    """
    df = scenario_data[scenario_name].copy()

    balance_cols = [c for c in df.columns if "Balance" in str(c)]
    if balance_cols and balance_cols[0] in ("Balance", "Portfolio Balance"):
        balance_cols = balance_cols[1:]

    # Build balance-only DataFrame with CUSIP names.
    bal_df = df[balance_cols].copy()
    if len(cusip_names) == len(bal_df.columns):
        bal_df.columns = cusip_names
    bal_df.index = pd.to_datetime(df["Date"])
    bal_df = bal_df.reset_index()
    bal_df = bal_df.rename(columns={"index": "Date"})

    # Compute quarter-end dates.
    bal_df["Qdate"] = [
        d - pd.tseries.offsets.DateOffset(days=1) + pd.tseries.offsets.QuarterEnd()
        for d in bal_df["Date"]
    ]
    quarters = bal_df["Qdate"].unique()

    # Find the last cashflow date before each quarter-end.
    qe_dates = []
    for q in quarters:
        before = bal_df[bal_df["Date"] < q]
        if len(before) > 0:
            qe_dates.append(before["Date"].values[-1])

    qe = bal_df[bal_df["Date"].isin(qe_dates)].copy()
    qe["Date"] = qe["Qdate"]
    qe = qe.drop(columns=["Qdate"])

    return qe


# ---------------------------------------------------------------------------
# Hedging output
# ---------------------------------------------------------------------------

def hedging_output(
    scenario_data: dict[str, pd.DataFrame],
    cusip_names: list[str],
    scenario_name: str = "Scenario 5",
) -> pd.DataFrame:
    """Generate a hedging table with first 2 payment dates and balances per CUSIP.

    For derivatives desk use — shows when each tranche starts paying and
    the outstanding balance at those dates.

    Args:
        scenario_data: Dict from read_cashflow_summary().
        cusip_names: CUSIP identifiers.
        scenario_name: Which scenario to use (default "Scenario 5" = Flat).

    Returns:
        DataFrame with columns: CUSIP, Payment Date 1, Balance 1, Payment Date 2, Balance 2
    """
    df = scenario_data[scenario_name].copy()

    interest_cols = [c for c in df.columns if "Interest" in str(c)]
    balance_cols = [c for c in df.columns if "Balance" in str(c)]

    if interest_cols and interest_cols[0] in ("Interest", "Portfolio Interest"):
        interest_cols = interest_cols[1:]
    if balance_cols and balance_cols[0] in ("Balance", "Portfolio Balance"):
        balance_cols = balance_cols[1:]

    interest_df = df[interest_cols].copy()
    balance_df = df[balance_cols].copy()

    if len(cusip_names) == len(interest_df.columns):
        interest_df.columns = cusip_names
        balance_df.columns = cusip_names

    dates = pd.to_datetime(df["Date"])

    rows = []
    for cusip in cusip_names:
        if cusip not in interest_df.columns:
            continue

        # Find rows where interest is non-zero (payment dates).
        paying = interest_df[interest_df[cusip] != 0].index
        pd1 = str(dates.iloc[paying[0]].date()) if len(paying) >= 1 else "NA"
        bal1 = balance_df[cusip].iloc[paying[0]] if len(paying) >= 1 else "NA"
        pd2 = str(dates.iloc[paying[1]].date()) if len(paying) >= 2 else "NA"
        bal2 = balance_df[cusip].iloc[paying[1]] if len(paying) >= 2 else "NA"

        rows.append({
            "CUSIP": cusip,
            "Payment Date 1": pd1,
            "Balance 1": bal1,
            "Payment Date 2": pd2,
            "Balance 2": bal2,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Visualization
# ---------------------------------------------------------------------------

def plot_balance_scenarios(
    balance_df: pd.DataFrame,
    title: str = "Forecasted Balances by Spread Scenario",
    figsize: tuple[int, int] = (12, 8),
):
    """Plot portfolio balance across all scenarios.

    Args:
        balance_df: DataFrame from build_balance_summary() with Date index
            and one column per scenario.
        title: Chart title.
        figsize: Figure size.

    Returns:
        matplotlib Figure.
    """
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    import seaborn as sns

    def billions(x, pos):
        return f"{x * 1e-9:.1f}B"

    fig, ax = plt.subplots(figsize=figsize)
    sns.lineplot(data=balance_df, palette="colorblind", ax=ax)
    ax.yaxis.set_major_formatter(FuncFormatter(billions))
    ax.set_title(title)
    ax.set_xlabel("")
    ax.set_ylabel("Portfolio Balance")
    plt.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# Full reporting output (mirrors the FHLB notebook)
# ---------------------------------------------------------------------------

def generate_report(
    cashflow_path: str | Path,
    output_path: str | Path,
    flat_scenario: str = "Scenario 5",
    scenario_labels: Optional[list[str]] = None,
) -> None:
    """Generate the full reporting output from a cashflow summary workbook.

    Produces an Excel file with:
      - Balance Summary: portfolio balance across all scenarios
      - Quarterly Forecasting: per-CUSIP balances at quarter-ends (flat scenario)
      - Hedging: first 2 payment dates/balances per CUSIP (flat scenario)

    Also displays a balance scenario chart.

    Args:
        cashflow_path: Path to the cashflow summary Excel file.
        output_path: Path to write the report workbook.
        flat_scenario: Name of the flat/base scenario sheet.
        scenario_labels: Optional friendly labels for scenarios.
    """
    cashflow_path = Path(cashflow_path)
    output_path = Path(output_path)

    if scenario_labels is None:
        scenario_labels = [
            "75 Tight", "50 Tight", "25 Tight", "10 Tight",
            "Flat", "10 Wide", "25 Wide", "50 Wide",
        ]

    print(f"Generating report from '{cashflow_path.name}'...")

    # Read data.
    cusip_names = extract_cusip_names(cashflow_path)
    scenario_data = read_cashflow_summary(cashflow_path)

    # Balance summary.
    balance_df = build_balance_summary(scenario_data, cusip_names)
    if len(scenario_labels) == len(balance_df.columns):
        balance_df.columns = scenario_labels

    # Quarterly roll-up (flat scenario).
    qe = quarterly_rollup(scenario_data, cusip_names, flat_scenario)

    # Hedging output (flat scenario).
    hedging = hedging_output(scenario_data, cusip_names, flat_scenario)

    # Write report.
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        balance_df.to_excel(writer, sheet_name="Balance Summary")
        qe.to_excel(writer, sheet_name="Quarterly Forecasting", index=False)
        hedging.to_excel(writer, sheet_name="Hedging", index=False)

    print(f"  Report written to '{output_path.name}'.")
    print(f"    Balance Summary: {len(balance_df)} periods × {len(balance_df.columns)} scenarios")
    print(f"    Quarterly Forecasting: {len(qe)} quarter-ends × {len(cusip_names)} CUSIPs")
    print(f"    Hedging: {len(hedging)} CUSIPs")

    # Plot.
    try:
        fig = plot_balance_scenarios(balance_df)
        fig.savefig(output_path.with_suffix(".png"), dpi=150, bbox_inches="tight")
        print(f"    Chart saved to '{output_path.with_suffix('.png').name}'.")
    except Exception as e:
        print(f"    Chart generation skipped: {e}")

    return balance_df
