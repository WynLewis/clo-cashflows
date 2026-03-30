"""
Reporting and visualization for CLO cashflow forecasts.

Builds reports directly from in-memory Python data — no Excel round-trip.

Features (from CLO_FHLB_CashflowForecast.ipynb and
CLO_Derivaties_CashflowForecast.ipynb):
  - Balance scenario chart (line plot across all scenarios)
  - Quarterly balance roll-up (snap to quarter-end dates)
  - Hedging output (first 2 payment dates and balances per CUSIP)
  - Per-scenario and per-CUSIP cashflow breakdowns
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from forecast.models import CashflowsReport, Scenario


# ---------------------------------------------------------------------------
# Build per-scenario DataFrames from in-memory CashflowsReport objects
# ---------------------------------------------------------------------------

def build_scenario_cashflows(
    reports: list[CashflowsReport],
) -> dict[str, dict[str, pd.DataFrame]]:
    """Organize CashflowsReport objects into per-scenario, per-CUSIP DataFrames.

    Returns:
        Dict[scenario_name -> Dict["dates"/"interest"/"principal"/"balance" -> DataFrame]]
        Each DataFrame has Date index and one column per CUSIP.
    """
    # Group reports by scenario.
    by_scenario: dict[str, list[CashflowsReport]] = {}
    for r in reports:
        by_scenario.setdefault(r.scenario_name, []).append(r)

    result = {}
    for scen_name, scen_reports in by_scenario.items():
        # Collect all dates across reports.
        all_dates: set[date] = set()
        for r in scen_reports:
            all_dates.update(r.cashflows.keys())
        sorted_dates = sorted(all_dates)
        if not sorted_dates:
            continue

        interest_data = {}
        principal_data = {}
        balance_data = {}

        for r in scen_reports:
            label = r.cusip or r.tranche or r.ws_name
            int_vals, prin_vals, bal_vals = [], [], []
            prev_bal = r.orig_principal

            for d in sorted_dates:
                cf = r.cashflows.get(d)
                if cf:
                    int_vals.append(cf.interest)
                    prin_vals.append(cf.principal)
                    bal_vals.append(cf.balance)
                    prev_bal = cf.balance
                else:
                    int_vals.append(0.0)
                    prin_vals.append(0.0)
                    bal_vals.append(prev_bal)

            interest_data[label] = int_vals
            principal_data[label] = prin_vals
            balance_data[label] = bal_vals

        dates_series = pd.to_datetime(sorted_dates)
        result[scen_name] = {
            "interest": pd.DataFrame(interest_data, index=dates_series),
            "principal": pd.DataFrame(principal_data, index=dates_series),
            "balance": pd.DataFrame(balance_data, index=dates_series),
        }

    return result


# ---------------------------------------------------------------------------
# Balance summary across scenarios
# ---------------------------------------------------------------------------

def build_balance_summary(
    reports: list[CashflowsReport],
    scenario_labels: Optional[dict[str, str]] = None,
) -> pd.DataFrame:
    """Build a portfolio-level balance summary across all scenarios.

    Args:
        reports: List of CashflowsReport objects.
        scenario_labels: Optional mapping from scenario_name to display label.

    Returns:
        DataFrame with Date index and one column per scenario (total balance).
    """
    scen_data = build_scenario_cashflows(reports)

    # Find the longest date range.
    all_dates: set = set()
    for sd in scen_data.values():
        all_dates.update(sd["balance"].index)
    sorted_dates = sorted(all_dates)

    balance_cols = {}
    for scen_name, sd in scen_data.items():
        total = sd["balance"].sum(axis=1)
        # Reindex to the full date range, forward-filling.
        total = total.reindex(sorted_dates, method="ffill")
        label = (scenario_labels or {}).get(scen_name, scen_name)
        balance_cols[label] = total.values

    return pd.DataFrame(balance_cols, index=sorted_dates)


# ---------------------------------------------------------------------------
# Quarterly roll-up
# ---------------------------------------------------------------------------

def quarterly_rollup(
    reports: list[CashflowsReport],
    scenario_name: str = "Scenario 5",
) -> pd.DataFrame:
    """Snap per-CUSIP balances to quarter-end dates for a given scenario.

    Args:
        reports: List of CashflowsReport objects.
        scenario_name: Which scenario to use (default "Scenario 5" = Flat).

    Returns:
        DataFrame with quarter-end Date column and per-CUSIP balance columns.
    """
    scen_data = build_scenario_cashflows(reports)
    if scenario_name not in scen_data:
        available = list(scen_data.keys())
        raise ValueError(f"Scenario '{scenario_name}' not found. Available: {available}")

    bal_df = scen_data[scenario_name]["balance"].copy()
    bal_df.index.name = "Date"
    bal_df = bal_df.reset_index()

    # Compute quarter-end for each date.
    bal_df["Qdate"] = [
        d - pd.tseries.offsets.DateOffset(days=1) + pd.tseries.offsets.QuarterEnd()
        for d in bal_df["Date"]
    ]
    quarters = sorted(bal_df["Qdate"].unique())

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
    reports: list[CashflowsReport],
    scenario_name: str = "Scenario 5",
) -> pd.DataFrame:
    """Generate a hedging table: first 2 payment dates and balances per CUSIP.

    Args:
        reports: List of CashflowsReport objects.
        scenario_name: Which scenario to use (default "Scenario 5" = Flat).

    Returns:
        DataFrame with CUSIP, Payment Date 1, Balance 1, Payment Date 2, Balance 2.
    """
    scen_data = build_scenario_cashflows(reports)
    if scenario_name not in scen_data:
        available = list(scen_data.keys())
        raise ValueError(f"Scenario '{scenario_name}' not found. Available: {available}")

    int_df = scen_data[scenario_name]["interest"]
    bal_df = scen_data[scenario_name]["balance"]

    rows = []
    for cusip in int_df.columns:
        # Find rows where interest is non-zero.
        paying = int_df.index[int_df[cusip] != 0]
        pd1 = str(paying[0].date()) if len(paying) >= 1 else "NA"
        bal1 = bal_df.loc[paying[0], cusip] if len(paying) >= 1 else "NA"
        pd2 = str(paying[1].date()) if len(paying) >= 2 else "NA"
        bal2 = bal_df.loc[paying[1], cusip] if len(paying) >= 2 else "NA"

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
        balance_df: DataFrame from build_balance_summary().
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
# Full report generation
# ---------------------------------------------------------------------------

def generate_report(
    reports: list[CashflowsReport],
    output_path: str | Path,
    flat_scenario: str = "Scenario 5",
    scenario_labels: Optional[dict[str, str]] = None,
) -> pd.DataFrame:
    """Generate the full reporting output from in-memory cashflow data.

    Produces an Excel file with:
      - Balance Summary: portfolio balance across all scenarios
      - Quarterly Forecasting: per-CUSIP balances at quarter-ends (flat scenario)
      - Hedging: first 2 payment dates/balances per CUSIP (flat scenario)

    Also saves a balance scenario chart as PNG.

    Args:
        reports: List of CashflowsReport objects (from load_cashflows_reports).
        output_path: Path to write the report workbook.
        flat_scenario: Name of the flat/base scenario.
        scenario_labels: Optional {scenario_name: display_label} mapping.

    Returns:
        The balance summary DataFrame.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"Generating report...")

    # Build outputs from in-memory data.
    balance_df = build_balance_summary(reports, scenario_labels)
    qe = quarterly_rollup(reports, flat_scenario)
    hedging = hedging_output(reports, flat_scenario)

    # Write Excel report.
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        balance_df.to_excel(writer, sheet_name="Balance Summary")
        qe.to_excel(writer, sheet_name="Quarterly Forecasting", index=False)
        hedging.to_excel(writer, sheet_name="Hedging", index=False)

    print(f"  Report written to '{output_path.name}'.")
    print(f"    Balance Summary: {len(balance_df)} periods x {len(balance_df.columns)} scenarios")
    print(f"    Quarterly Forecasting: {len(qe)} quarter-ends x {len(qe.columns) - 1} CUSIPs")
    print(f"    Hedging: {len(hedging)} CUSIPs")

    # Chart.
    try:
        fig = plot_balance_scenarios(balance_df)
        chart_path = output_path.with_suffix(".png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight")
        print(f"    Chart saved to '{chart_path.name}'.")
    except Exception as e:
        print(f"    Chart generation skipped: {e}")

    return balance_df
