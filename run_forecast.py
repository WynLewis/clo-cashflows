#!/usr/bin/env python3
"""
CLO Cashflow Forecast — Main Orchestrator

Replaces the 7-step macro workflow from Forecast v16.3.xlsm with Python.

Usage:
    python run_forecast.py --forecast-wb "Forecast v16.3.xlsm"

The 7-step workflow:
    Step 1: Import CLO holdings from a Structured Products Data Packet report.
    Step 2: Enrich holdings with Intex/Bloomberg data (read from workbook).
    Step 3: Clear forecasted factors.
    Step 4: Generate preliminary Intex Portfolio Upload (with COLLAT tranches).
    Step 5: Extract factor call dates from preliminary Intex Cashflows report.
    Step 6: Generate final Intex Portfolio Upload (with factor call dates).
    Step 7: Summarize final cashflows by scenario.
"""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

from forecast.cashflows_summary import generate_cashflow_summary, load_cashflows_reports
from forecast.factors import ForecastedFactors
from forecast.forward_curve import load_from_cashflows_report, load_from_workbook
from forecast.holdings import (
    export_tranches,
    import_holdings,
    load_holdings_from_forecast_workbook,
)
from forecast.preprice import PrePriceDeals
from forecast.scenarios import (
    generate_portfolio_upload,
    load_scenarios_from_workbook,
    write_portfolio_upload,
)


def run_full_workflow(
    forecast_wb: str | Path,
    data_packet: str | Path | None = None,
    preliminary_cashflows: str | Path | None = None,
    final_cashflows: str | Path | None = None,
    output_dir: str | Path | None = None,
    fhlb_only: bool = False,
) -> None:
    """Execute the full 7-step forecast workflow.

    Args:
        forecast_wb: Path to the Forecast v16.3.xlsm workbook (source of scenarios,
                     pre-price deals, forward curve, and optionally holdings).
        data_packet: Path to Structured Products Data Packet .xlsx (Step 1).
                     If None, holdings are loaded from the forecast workbook.
        preliminary_cashflows: Path to preliminary Intex Cashflows export (Step 5).
                               If None, Step 5 is skipped and factor call dates are
                               loaded from the forecast workbook.
        final_cashflows: Path to final Intex Cashflows export (Step 7).
                         If None, Step 7 is skipped.
        output_dir: Directory for output files. Defaults to current directory.
        fhlb_only: If True, import only FHLB CLO holdings.
    """
    forecast_wb = Path(forecast_wb)
    output_dir = Path(output_dir) if output_dir else Path(".")
    output_dir.mkdir(parents=True, exist_ok=True)
    today_str = date.today().strftime("%Y%m%d")

    # ── Step 1: Import holdings ──────────────────────────────────────────
    print("\n" + "=" * 70)
    print("STEP 1: Import CLO Holdings")
    print("=" * 70)

    if data_packet:
        holdings_df = import_holdings(data_packet, fhlb_only=fhlb_only)
    else:
        holdings_df = load_holdings_from_forecast_workbook(forecast_wb)

    # ── Step 2: Enrich with Intex/Bloomberg data ─────────────────────────
    print("\n" + "=" * 70)
    print("STEP 2: Enrich Holdings Data")
    print("=" * 70)
    print("  Holdings data read from workbook (Intex/Bloomberg fields pre-populated).")

    # Load pre-price deals and export tranches.
    preprice = PrePriceDeals.from_forecast_workbook(forecast_wb)
    tranches = export_tranches(holdings_df, preprice)

    # Load scenarios.
    scenarios = load_scenarios_from_workbook(forecast_wb)

    # ── Step 3: Clear forecasted factors ─────────────────────────────────
    print("\n" + "=" * 70)
    print("STEP 3: Clear Forecasted Factors")
    print("=" * 70)
    factors = ForecastedFactors()
    factors.clear()
    print("  Forecasted factors cleared.")

    # ── Step 4: Generate preliminary portfolio upload ─────────────────────
    print("\n" + "=" * 70)
    print("STEP 4: Generate Preliminary Intex Portfolio Upload")
    print("=" * 70)
    prelim_upload = generate_portfolio_upload(
        scenarios=scenarios,
        tranches=tranches,
        collateral_cashflows=True,
    )
    prelim_path = output_dir / f"{today_str}_PreliminaryPortfolioUpload.xlsx"
    write_portfolio_upload(prelim_upload, prelim_path)
    print(f"  --> Upload file to IntexCalc and export the Cashflows report.")

    # ── Step 5: Extract factor call dates ────────────────────────────────
    print("\n" + "=" * 70)
    print("STEP 5: Extract Forecasted Factor Call Dates")
    print("=" * 70)

    if preliminary_cashflows:
        factors.import_from_cashflows_report(preliminary_cashflows, tranches)

        # Also load forward curve from preliminary cashflows.
        curve_df = load_from_cashflows_report(preliminary_cashflows)
    else:
        # Load pre-existing factor data and curve from the forecast workbook.
        factors.load_from_workbook(forecast_wb)
        curve_df = load_from_workbook(forecast_wb)

    print(f"  Forward curve: {len(curve_df)} periods loaded.")
    print(f"  Factor call dates: {factors.count} entries.")

    # ── Step 6: Generate final portfolio upload ──────────────────────────
    print("\n" + "=" * 70)
    print("STEP 6: Generate Final Intex Portfolio Upload")
    print("=" * 70)
    final_upload = generate_portfolio_upload(
        scenarios=scenarios,
        tranches=tranches,
        collateral_cashflows=False,
        forecasted_factors=factors,
    )
    final_path = output_dir / f"{today_str}_FinalPortfolioUpload.xlsx"
    write_portfolio_upload(final_upload, final_path)
    print(f"  --> Upload file to IntexCalc to generate analytics and cashflows.")

    # ── Step 7: Generate cashflow summary ────────────────────────────────
    print("\n" + "=" * 70)
    print("STEP 7: Generate Cashflow Summary")
    print("=" * 70)

    if final_cashflows:
        reports = load_cashflows_reports(final_cashflows)
        summary_path = output_dir / f"{today_str}_CashflowSummary.xlsx"
        generate_cashflow_summary(reports, summary_path, scenarios)
    else:
        print("  No final cashflows file provided — skipping summary generation.")
        print("  After running the final upload in IntexCalc, re-run with --final-cashflows.")

    # ── Done ─────────────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("WORKFLOW COMPLETE")
    print("=" * 70)
    print(f"  Output directory: {output_dir.resolve()}")


def main():
    parser = argparse.ArgumentParser(
        description="CLO Cashflow Forecast — Python replacement for Forecast v16.3.xlsm",
    )
    parser.add_argument(
        "--forecast-wb",
        required=True,
        help="Path to the Forecast v16.3.xlsm workbook.",
    )
    parser.add_argument(
        "--data-packet",
        default=None,
        help="Path to Structured Products Data Packet .xlsx (Step 1). "
             "If omitted, holdings are loaded from the forecast workbook.",
    )
    parser.add_argument(
        "--preliminary-cashflows",
        default=None,
        help="Path to preliminary Intex Cashflows export (Step 5). "
             "If omitted, factor data is loaded from the forecast workbook.",
    )
    parser.add_argument(
        "--final-cashflows",
        default=None,
        help="Path to final Intex Cashflows export (Step 7). "
             "If omitted, Step 7 is skipped.",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="Output directory for generated files.",
    )
    parser.add_argument(
        "--fhlb-only",
        action="store_true",
        help="Import only FHLB CLO holdings (PAM Portfolio 13091).",
    )
    args = parser.parse_args()

    run_full_workflow(
        forecast_wb=args.forecast_wb,
        data_packet=args.data_packet,
        preliminary_cashflows=args.preliminary_cashflows,
        final_cashflows=args.final_cashflows,
        output_dir=args.output_dir,
        fhlb_only=args.fhlb_only,
    )


if __name__ == "__main__":
    main()
