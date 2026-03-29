"""
Scenario setup and Intex Portfolio Upload file generation.

Mirrors VBA worksheet module: scenario_setup.cls

Step 4: Generate preliminary Intex Portfolio Upload (with COLLAT tranches).
Step 6: Generate final Intex Portfolio Upload (with factor call dates).
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd

from forecast.config import CONFIG
from forecast.factors import ForecastedFactors
from forecast.models import Scenario, Tranche


# ---------------------------------------------------------------------------
# Scenario loading
# ---------------------------------------------------------------------------

def load_scenarios_from_workbook(wb_path: str | Path) -> list[Scenario]:
    """Load scenarios from the Forecast workbook's 'Scenario Setup' sheet.

    Only scenarios where 'Use?' is True are returned.
    Mirrors VBA method: scenario_setup.GetScenarios()
    """
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    ws = wb["Scenario Setup"]

    # Headers at row 4: Use?, No., Scenario Name, Settle Date, Initial AAA Margin,
    #                    AAA Margin Shock, Prepay Speed, Horizon Given Type
    scenarios: list[Scenario] = []
    row_num = 5
    while True:
        use = ws.cell(row=row_num, column=1).value
        if use is None:
            break
        if use:
            settle_val = ws.cell(row=row_num, column=4).value
            if isinstance(settle_val, datetime):
                settle_date = settle_val.date()
            elif isinstance(settle_val, date):
                settle_date = settle_val
            else:
                settle_date = date.today()

            s = Scenario(
                number=int(ws.cell(row=row_num, column=2).value),
                name=str(ws.cell(row=row_num, column=3).value),
                settle_date=settle_date,
                initial_aaa_margin=float(ws.cell(row=row_num, column=5).value),
                aaa_margin_shock=float(ws.cell(row=row_num, column=6).value),
                prepay_speed=int(ws.cell(row=row_num, column=7).value or 15),
                horizon_given_type=str(ws.cell(row=row_num, column=8).value or "DISC_MARGIN"),
            )
            scenarios.append(s)
        row_num += 1

    wb.close()
    print(f"  {len(scenarios)} scenarios loaded from Scenario Setup.")
    return scenarios


def load_scenarios_from_dataframe(df: pd.DataFrame) -> list[Scenario]:
    """Load scenarios from a DataFrame (alternative to reading from workbook)."""
    scenarios = []
    for _, row in df.iterrows():
        if not row.get("Use?", True):
            continue
        scenarios.append(Scenario(
            number=int(row["No."]),
            name=str(row["Scenario Name"]),
            settle_date=pd.Timestamp(row["Settle Date"]).date(),
            initial_aaa_margin=float(row["Initial AAA Margin"]),
            aaa_margin_shock=float(row["AAA Margin Shock"]),
            prepay_speed=int(row.get("Prepay Speed", 15)),
            horizon_given_type=str(row.get("Horizon Given Type", "DISC_MARGIN")),
        ))
    return scenarios


def load_scenarios_from_config() -> list[Scenario]:
    """Build Scenario objects from the ScenarioConfig in forecast/config.py.

    This lets you define the current spread level and scenario shocks in Python
    without needing the Excel workbook's Scenario Setup sheet.

    Usage:
        from forecast.config import CONFIG

        # Set current AAA spread and settle date
        CONFIG.scenario.current_aaa_margin_bps = 120
        CONFIG.scenario.settle_date = date(2026, 3, 18)
        CONFIG.scenario.prepay_speed = 15

        # Define scenarios (or keep the defaults)
        CONFIG.scenario.scenarios = [
            ScenarioDefinition("100 bps Tightening", -100),
            ScenarioDefinition("Flat", 0),
            ScenarioDefinition("100 bps Widening", 100),
        ]

        scenarios = load_scenarios_from_config()
    """
    sc = CONFIG.scenario
    scenarios = []
    for i, defn in enumerate(sc.scenarios, start=1):
        scenarios.append(Scenario(
            number=i,
            name=defn.name,
            settle_date=sc.settle_date,
            initial_aaa_margin=sc.current_aaa_margin_bps,
            aaa_margin_shock=defn.aaa_margin_shock_bps,
            prepay_speed=sc.prepay_speed,
            horizon_given_type=sc.horizon_given_type,
        ))
    print(f"  {len(scenarios)} scenarios loaded from config "
          f"(current AAA margin = {sc.current_aaa_margin_bps} bps, "
          f"settle = {sc.settle_date}).")
    return scenarios


# ---------------------------------------------------------------------------
# Optional Redemption logic
# ---------------------------------------------------------------------------

def optional_redemption(
    deal_aaa_margin: float,
    initial_aaa_margin: float,
    aaa_shock: float,
    middle_market: bool,
    reset_index: str,
) -> bool:
    """Determine if a deal is in-the-money to be called based on coupon vs. market levels.

    Mirrors VBA function: scenario_setup.OptionalRedemption()

    Args:
        deal_aaa_margin: The deal's AAA margin in bps (already multiplied by 100 in caller).
        initial_aaa_margin: Base AAA margin assumption in bps.
        aaa_shock: Spread shock in bps.
        middle_market: Whether the deal is a middle market CLO.
        reset_index: The index the tranche resets against (e.g. "US0003M" for LIBOR).

    Returns:
        True if the deal is expected to be called (refinanced).
    """
    strike = initial_aaa_margin + aaa_shock + CONFIG.call.refi_costs_bps

    if middle_market:
        strike += CONFIG.call.middle_market_bsl_basis_bps

    # Adjust for LIBOR-SOFR basis.
    if reset_index == CONFIG.call.libor_reset_index:
        strike -= CONFIG.call.libor_sofr_basis_bps

    return deal_aaa_margin > strike


# ---------------------------------------------------------------------------
# Portfolio Upload file generation
# ---------------------------------------------------------------------------

def generate_portfolio_upload(
    scenarios: list[Scenario],
    tranches: list[Tranche],
    collateral_cashflows: bool,
    forecasted_factors: Optional[ForecastedFactors] = None,
) -> pd.DataFrame:
    """Generate the Intex Portfolio Upload file as a DataFrame.

    This is the core function that mirrors VBA scenario_setup.GeneratePortfolioUploadFile().

    Args:
        scenarios: List of Scenario objects (from Scenario Setup).
        tranches: List of Tranche objects (from Initial Holdings + Pre-Price Deals).
        collateral_cashflows: If True, include COLLAT tranches (Step 4 / preliminary run).
                              If False, generate final upload with factor call dates (Step 6).
        forecasted_factors: Factor call date lookup (only used when collateral_cashflows=False).

    Returns:
        DataFrame in Intex Portfolio Upload format, ready to write to the 'Scenarios' sheet.
    """
    if forecasted_factors is None:
        forecasted_factors = ForecastedFactors()

    columns = [
        "Scenario", "Scenario Name", "User Comment2", "Settle Date", "Deal",
        "Preprice Password", "Orig Face", "Given Type", "Given Amount",
        "Enable Reinvestment Overrides", "Reinvestment Profile", "Run Calls",
        "Call Units", "Call Value", "User Comment", "Prepay Units", "Prepay",
        "Default Units", "Default", "Severity Units", "Severity", "Recovery Lag",
    ]
    rows: list[dict] = []

    for s in scenarios:
        # Track unique deals to write COLLAT tranche once per deal.
        deals_seen: set[str] = set()

        for t in tranches:
            settle_str = s.settle_date.strftime("%Y%m%d")

            if collateral_cashflows and t.intex_deal_name not in deals_seen:
                # Write COLLAT tranche (one per deal per scenario).
                row = _build_collat_row(s, t, settle_str)
                rows.append(row)
                deals_seen.add(t.intex_deal_name)

            if not collateral_cashflows:
                # For final upload: write one row per deal (first tranche seen) then per tranche.
                if t.intex_deal_name not in deals_seen:
                    deals_seen.add(t.intex_deal_name)

                row = _build_tranche_row(s, t, settle_str, forecasted_factors)
                rows.append(row)

    df = pd.DataFrame(rows, columns=columns)
    action = "preliminary (with COLLAT)" if collateral_cashflows else "final (with factor calls)"
    print(f"  Portfolio upload generated: {len(df)} rows ({action}).")
    return df


def write_portfolio_upload(df: pd.DataFrame, output_path: str | Path) -> None:
    """Write the portfolio upload DataFrame to an Excel workbook."""
    output_path = Path(output_path)
    df.to_excel(output_path, sheet_name="Scenarios", index=False, engine="openpyxl")
    print(f"  Portfolio upload written to '{output_path.name}'.")


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _build_collat_row(s: Scenario, t: Tranche, settle_str: str) -> dict:
    """Build one row for a COLLAT tranche in the portfolio upload."""
    row: dict = {
        "Scenario": s.number,
        "Scenario Name": s.name,
        "User Comment2": f"Scenario {s.number}",
        "Settle Date": settle_str,
        "Deal": f"{t.intex_deal_name},COLLAT",
        "Orig Face": "FULL",
        "Given Type": "PRICE100",
        "Given Amount": 100,
        "Enable Reinvestment Overrides": 1,
        "Reinvestment Profile": CONFIG.reinvestment.model_name,
        "Run Calls": 1,
        "Default Units": CONFIG.defaults.default_units,
        "Default": CONFIG.defaults.default_rate,
        "Severity Units": CONFIG.defaults.severity_units,
        "Severity": CONFIG.defaults.severity_pct,
        "Recovery Lag": CONFIG.defaults.recovery_lag_months,
        "Prepay Units": "CPR",
    }

    if t.preprice:
        row["Preprice Password"] = t.preprice_password

    # Determine call date based on refi economics.
    if optional_redemption(t.aaa_margin * 100, s.initial_aaa_margin, s.aaa_margin_shock, t.middle_market, t.reset_index):
        if t.non_call_end is not None and s.settle_date > t.non_call_end:
            row["Call Units"] = "ASAP + (mos)"
            row["Call Value"] = 0
            row["User Comment"] = "Refi Call"
        elif t.non_call_end is not None:
            row["Call Units"] = "Date (Exact)"
            row["Call Value"] = t.non_call_end.strftime("%Y%m%d")
            row["User Comment"] = "Refi Call"
        else:
            row["Call Units"] = "Never"
            row["Call Value"] = "NEVER"
    else:
        row["Call Units"] = "Never"
        row["Call Value"] = "NEVER"

    # Collateral prepayments.
    reinvest_months = t.reinvest_months(s.settle_date)
    if reinvest_months > 0:
        row["Prepay"] = f"{s.prepay_speed} FOR {reinvest_months} 0"
    else:
        row["Prepay"] = "0"

    return row


def _build_tranche_row(
    s: Scenario,
    t: Tranche,
    settle_str: str,
    forecasted_factors: ForecastedFactors,
) -> dict:
    """Build one row for a tranche in the final portfolio upload."""
    row: dict = {
        "Scenario": s.number,
        "Scenario Name": s.name,
        "User Comment2": f"Scenario {s.number}",
        "Settle Date": settle_str,
        "Deal": t.intex_name,
        "Orig Face": t.orig_face,
        "Given Type": "PRICE100",
        "Given Amount": t.price,
        "Enable Reinvestment Overrides": 1,
        "Reinvestment Profile": CONFIG.reinvestment.model_name,
        "Run Calls": 1,
        "Default Units": CONFIG.defaults.default_units,
        "Default": CONFIG.defaults.default_rate,
        "Severity Units": CONFIG.defaults.severity_units,
        "Severity": CONFIG.defaults.severity_pct,
        "Recovery Lag": CONFIG.defaults.recovery_lag_months,
        "Prepay Units": "CPR",
    }

    if t.preprice:
        row["Preprice Password"] = t.preprice_password

    # Determine call date based on refi economics.
    if optional_redemption(t.aaa_margin * 100, s.initial_aaa_margin, s.aaa_margin_shock, t.middle_market, t.reset_index):
        if t.non_call_end is not None and s.settle_date > t.non_call_end:
            row["Call Units"] = "ASAP + (mos)"
            row["Call Value"] = 0
            row["User Comment"] = "Refi Call"
        elif t.non_call_end is not None:
            row["Call Units"] = "Date (Exact)"
            row["Call Value"] = t.non_call_end.strftime("%Y%m%d")
            row["User Comment"] = "Refi Call"
        else:
            row["Call Units"] = "Never"
            row["Call Value"] = "NEVER"
    else:
        row["Call Units"] = "Never"
        row["Call Value"] = "NEVER"

    # Check for factor call date that might be earlier than refi call.
    factor_call = forecasted_factors.get_factor_call_date(s.name, t.intex_deal_name)
    if factor_call is not None:
        if row["Call Units"] == "Date (Exact)" and t.non_call_end is not None:
            # Override with factor call if it's earlier.
            if t.non_call_end > factor_call:
                row["Call Value"] = factor_call.strftime("%Y%m%d")
                row["User Comment"] = "Factor Call"
        elif row["Call Units"] != "ASAP + (mos)":
            # No refi call was set — use factor call.
            row["Call Units"] = "Date (Exact)"
            row["Call Value"] = factor_call.strftime("%Y%m%d")
            row["User Comment"] = "Factor Call"

    # Collateral prepayments.
    reinvest_months = t.reinvest_months(s.settle_date)
    if reinvest_months > 0:
        row["Prepay"] = f"{s.prepay_speed} FOR {reinvest_months} 0"
    else:
        row["Prepay"] = "0"

    return row
