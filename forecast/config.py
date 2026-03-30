"""
Configuration for CLO cashflow forecast parameters.

All tunable business-logic constants are defined here so they can be
adjusted without editing the core modules.  Import this module and
override values before running the workflow, or edit this file directly.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


@dataclass
class CallAssumptions:
    """Parameters governing optional-redemption (call) logic."""

    # Basis points of friction cost assumed for refinancing.
    refi_costs_bps: float = 10

    # Basis-point premium added to the strike for middle-market CLOs
    # (middle-market AAA spreads are wider than BSL).
    middle_market_bsl_basis_bps: float = 50

    # LIBOR-to-SOFR basis adjustment in bps (applied when the tranche
    # resets against 3-month LIBOR rather than SOFR).
    libor_sofr_basis_bps: float = 26.161

    # Reset index identifier that triggers the LIBOR-SOFR adjustment.
    libor_reset_index: str = "US0003M"

    # Bloomberg collateral type string that identifies middle-market CLOs.
    middle_market_collat_type: str = "CF-CLO-MML"


@dataclass
class ScenarioDefinition:
    """Definition for a single spread scenario.

    Attributes:
        name: Human-readable label (e.g. "75 bps Tightening").
        aaa_margin_shock_bps: Spread change from current level in bps.
            Negative = tightening, positive = widening, 0 = flat.
    """

    name: str
    aaa_margin_shock_bps: float


@dataclass
class DefaultAssumptions:
    """Default assumptions written to Intex uploads.

    Everything except current_aaa_margin_bps lives here.
    """

    # Settlement date for the scenarios.
    settle_date: date = field(default_factory=date.today)

    # Prepay speed assumption (CPR) applied to all scenarios.
    prepay_speed: int = 15

    # Constant default rate (CDR), in percent.
    default_rate: float = 5

    # Default rate units.
    default_units: str = "CDR"

    # Loss severity, in percent.
    severity_pct: float = 50

    # Severity units.
    severity_units: str = "Percent"

    # Recovery lag in months.
    recovery_lag_months: int = 12

    # Horizon given type for Intex analytics.
    horizon_given_type: str = "DISC_MARGIN"

    # List of spread scenarios to run.  Scenarios are numbered automatically
    # in the order they appear here.
    scenarios: list[ScenarioDefinition] = field(default_factory=lambda: [
        ScenarioDefinition("75 bps Tightening", -75),
        ScenarioDefinition("50 bps Tightening", -50),
        ScenarioDefinition("25 bps Tightening", -25),
        ScenarioDefinition("10 bps Tightening", -10),
        ScenarioDefinition("Flat", 0),
        ScenarioDefinition("10 bps Widening", 10),
        ScenarioDefinition("25 bps Widening", 25),
        ScenarioDefinition("50 bps Widening", 50),
    ])


@dataclass
class FactorCallAssumptions:
    """Parameters for factor-based call date extraction."""

    # Deal factor threshold below which a deal is considered called.
    # When deal_balance / orig_deal_balance < this value, the deal is called.
    factor_threshold: float = 0.35


@dataclass
class ReinvestmentDefaults:
    """Default reinvestment model parameters (mirrors the Reinvestment Models sheet)."""

    model_name: str = "#DefaultReinvestAsset"
    reinvest_type: str = "Float"
    percent: float = 100
    index: str = "SOFR (3mo)"
    coupon_spread: str = "WAVG"
    maturity_months: int = 60
    price: float = 99.5
    life_floor: float = 0
    asset_type: str = "Loan"
    asset_subtype: str = "TL"
    amortization: str = "Bullet"
    daycount: str = "Actual360"
    frequency: str = "Quarterly"


@dataclass
class ExcelAddIns:
    """Excel add-in configuration for IntexLINK and Bloomberg.

    The add-in names are used to toggle COM add-ins on/off via xlwings.
    Check yours in Excel via:
        File → Options → Add-ins → Manage: COM Add-ins → Go

    If INTEX() formulas show #NAME? errors, set intex_prefix to the
    XLL add-in filename.  Check via:
        File → Options → Add-ins → Manage: Excel Add-ins → Go
    """

    # COM Add-in ProgIDs (used for toggling on/off).
    # These are the names that appear in Excel's COM Add-ins dialog.
    # Common values — update if yours differ:
    intex_com_addin: str = "IntexLINK.Connect"
    bloomberg_com_addin: str = "Bloomberg Excel COM Add-In"

    # XLL add-in paths (used for toggling on/off).
    # Set to "" to skip XLL toggling.  Check yours in Excel via:
    #   File → Options → Add-ins → Manage: Excel Add-ins → Go
    intex_xll_path: str = ""  # e.g. "C:\\IntexLINK\\IntexLINK.xll"
    bloomberg_xll_path: str = ""  # e.g. "C:\\blp\\API\\Office Tools\\BloombergUI.xll"

    # Prefix for INTEX() formulas.  Set to "" for no prefix, or
    # e.g. "_xll.IntexLINK" if you get #NAME? errors.
    intex_prefix: str = ""

    # Prefix for BDP() formulas (Bloomberg).  Usually "" works.
    bloomberg_prefix: str = ""

    def intex_func(self, func_name: str = "INTEX") -> str:
        """Return the full function call string, with prefix if needed."""
        if self.intex_prefix:
            return f"{self.intex_prefix}.{func_name}"
        return func_name

    def bbg_func(self, func_name: str = "BDP") -> str:
        """Return the full Bloomberg function call string."""
        if self.bloomberg_prefix:
            return f"{self.bloomberg_prefix}.{func_name}"
        return func_name


@dataclass
class ScenarioConfig:
    """Scenario setup — only the value that changes run to run.

    The call decision for each deal compares:
        deal_aaa_margin  vs  current_aaa_margin + shock + refi_costs

    Everything else (settle date, prepay, scenarios list, etc.)
    lives in DefaultAssumptions.
    """

    # Your view of where BSL AAA CLO new-issue spreads are today (in bps).
    # This is the ONLY value you need to set each run.
    current_aaa_margin_bps: float = 115


@dataclass
class ForecastConfig:
    """Top-level configuration combining all assumption sets."""

    call: CallAssumptions = field(default_factory=CallAssumptions)
    defaults: DefaultAssumptions = field(default_factory=DefaultAssumptions)
    factor_call: FactorCallAssumptions = field(default_factory=FactorCallAssumptions)
    reinvestment: ReinvestmentDefaults = field(default_factory=ReinvestmentDefaults)
    scenario: ScenarioConfig = field(default_factory=ScenarioConfig)
    excel: ExcelAddIns = field(default_factory=ExcelAddIns)


# Module-level default instance.  Import and modify this, or create your own.
CONFIG = ForecastConfig()
