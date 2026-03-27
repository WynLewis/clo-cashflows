"""
Cashflow summary generation.

Mirrors VBA module: cashflows_summary.bas

Step 7: Extract data from the final Intex Cashflows report and summarize
        interest, principal, and balance by scenario across all tranches.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd

from forecast.models import CashflowsReport, Scenario


def load_cashflows_reports(cashflows_path: str | Path) -> list[CashflowsReport]:
    """Load all worksheets from an Intex Cashflows Excel export into CashflowsReport objects.

    Skips sheets named "Sheet1" or starting with "Total".
    """
    cashflows_path = Path(cashflows_path)
    print(f"Loading cashflows from '{cashflows_path.name}'...")

    wb = openpyxl.load_workbook(cashflows_path, read_only=True, data_only=True)
    reports: list[CashflowsReport] = []

    for i, ws in enumerate(wb.worksheets):
        name = ws.title
        if name == "Sheet1" or name.startswith("Total"):
            print(f"  {i + 1}/{len(wb.worksheets)}: '{name}' --> SKIPPED")
            continue

        print(f"  {i + 1}/{len(wb.worksheets)}: '{name}' --> Extracting data...")
        report = CashflowsReport()
        report.load_data(ws)
        reports.append(report)

    wb.close()
    print(f"  {len(reports)} cashflows reports loaded.")
    return reports


def generate_cashflow_summary(
    reports: list[CashflowsReport],
    output_path: str | Path,
    scenarios: Optional[list[Scenario]] = None,
) -> None:
    """Generate the cashflow summary workbook with one sheet per scenario plus a Balance Summary.

    Mirrors VBA method: cashflows_summary.GenerateCashflowSummary()

    Args:
        reports: List of CashflowsReport objects (from load_cashflows_reports).
        output_path: Path to write the summary Excel workbook.
        scenarios: Optional list of Scenario objects for name-to-number mapping.
    """
    output_path = Path(output_path)

    # Build scenario name-to-number mapping if provided.
    scen_map: dict[str, str] = {}
    if scenarios:
        for s in scenarios:
            scen_map[s.name] = f"Scenario {s.number}"

    # Map scenario names for reports with raw names.
    for r in reports:
        if r.scenario_name in scen_map:
            pass  # Already in "Scenario N" format or mapped
        elif r.scenario_name not in [f"Scenario {i}" for i in range(1, 20)]:
            mapped = scen_map.get(r.scenario_name)
            if mapped:
                r.scenario_name = mapped

    # Extract unique scenario names (preserving order of first appearance).
    seen: dict[str, None] = {}
    for r in reports:
        if r.scenario_name not in seen:
            seen[r.scenario_name] = None
    scenario_names = list(seen.keys())

    print(f"Generating cashflow summary for {len(scenario_names)} scenarios...")

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        balance_summary_data: dict[str, dict[date, float]] = {}

        for scenario_name in scenario_names:
            # Get reports for this scenario.
            scen_reports = [r for r in reports if r.scenario_name == scenario_name]
            if not scen_reports:
                continue

            # Collect all unique dates across reports in this scenario.
            all_dates: set[date] = set()
            for r in scen_reports:
                all_dates.update(r.cashflows.keys())
            sorted_dates = sorted(all_dates)

            if not sorted_dates:
                continue

            # Build the scenario sheet data.
            # Columns: Date, Portfolio Interest, Portfolio Principal, Portfolio Balance,
            #          then for each tranche: Interest, Principal, Balance
            data: dict[str, list] = {"Date": sorted_dates}

            # Initialize portfolio totals.
            portfolio_interest = [0.0] * len(sorted_dates)
            portfolio_principal = [0.0] * len(sorted_dates)
            portfolio_balance = [0.0] * len(sorted_dates)

            for r in scen_reports:
                label = r.cusip or r.tranche or r.ws_name
                int_col = f"{label} Interest"
                prin_col = f"{label} Principal"
                bal_col = f"{label} Balance"
                data[int_col] = []
                data[prin_col] = []
                data[bal_col] = []

                prev_balance = r.orig_principal

                for i, d in enumerate(sorted_dates):
                    cf = r.cashflows.get(d)
                    if cf:
                        data[int_col].append(cf.interest)
                        data[prin_col].append(cf.principal)
                        data[bal_col].append(cf.balance)
                        portfolio_interest[i] += cf.interest
                        portfolio_principal[i] += cf.principal
                        prev_balance = cf.balance
                    else:
                        data[int_col].append(0)
                        data[prin_col].append(0)
                        data[bal_col].append(prev_balance)

                    # First row: use orig_principal for balance.
                    if i == 0:
                        data[bal_col][-1] = r.orig_principal
                        portfolio_balance[i] += r.orig_principal
                        prev_balance = r.orig_principal

            # Compute portfolio balance as declining from initial.
            for i in range(len(sorted_dates)):
                if i == 0:
                    pass  # Already summed above
                else:
                    portfolio_balance[i] = max(0, portfolio_balance[i - 1] - portfolio_principal[i])

            # Insert portfolio columns at the front.
            ordered_data: dict[str, list] = {"Date": sorted_dates}
            ordered_data["Portfolio Interest"] = portfolio_interest
            ordered_data["Portfolio Principal"] = portfolio_principal
            ordered_data["Portfolio Balance"] = portfolio_balance
            for k, v in data.items():
                if k != "Date":
                    ordered_data[k] = v

            df = pd.DataFrame(ordered_data)

            # Write to sheet (truncate name to 31 chars for Excel).
            sheet_name = scenario_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Save balance data for Balance Summary.
            balance_summary_data[scenario_name] = {
                d: b for d, b in zip(sorted_dates, portfolio_balance)
            }

        # Write Balance Summary sheet.
        if balance_summary_data:
            all_summary_dates: set[date] = set()
            for dates_dict in balance_summary_data.values():
                all_summary_dates.update(dates_dict.keys())
            sorted_summary_dates = sorted(all_summary_dates)

            summary_data: dict[str, list] = {"Date": sorted_summary_dates}
            for scen_name in scenario_names:
                col_data = []
                prev_val = 0.0
                for d in sorted_summary_dates:
                    val = balance_summary_data.get(scen_name, {}).get(d)
                    if val is not None:
                        col_data.append(val)
                        prev_val = val
                    else:
                        col_data.append(prev_val)
                summary_data[scen_name] = col_data

            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Balance Summary", index=False)

    print(f"  Cashflow summary written to '{output_path.name}'.")
