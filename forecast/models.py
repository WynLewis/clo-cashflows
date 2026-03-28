"""
Data models for CLO cashflow forecasting.

Mirrors the VBA class modules: tranche.cls, scenario.cls, cashflow.cls, cashflows_report.cls
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional


@dataclass
class Tranche:
    """Represents a CLO tranche (one per unique CUSIP, with face/par summed across positions).

    Mirrors VBA class module: tranche.cls
    """

    cusip: str
    orig_face: float = 0.0
    cur_par: float = 0.0
    price: float = 0.0
    intex_name: str = ""  # Deal,Tranche (e.g. "CAPC37C4,BR")
    intex_deal_name: str = ""  # Deal only (e.g. "CAPC37C4")
    preprice: bool = False
    preprice_password: str = ""
    aaa_margin: float = 0.0  # in decimal (e.g. 1.09 = 109 bps)
    non_call_end: Optional[date] = None
    reinvest_end: Optional[date] = None
    middle_market: bool = False
    oas: float = 0.0
    orig_deal_balance: float = 0.0
    reset_index: str = ""
    floater: bool = False

    def clean_intex_name(self) -> str:
        """Remove any pre-price password from the intex_name."""
        if "|" in self.intex_name:
            return self.intex_name[: self.intex_name.index("|")]
        return self.intex_name

    def current_factor(self) -> float:
        """Compute current factor = cur_par / orig_face."""
        if self.orig_face == 0:
            return 0.0
        return self.cur_par / self.orig_face

    def reinvest_months(self, start_date: date) -> int:
        """Compute months remaining in the reinvestment period from start_date (rounded down)."""
        if self.reinvest_end is None:
            return 0
        delta_days = (self.reinvest_end - start_date).days
        return max(0, math.floor(delta_days / 30))


@dataclass
class Scenario:
    """Represents a scenario from the Scenario Setup worksheet.

    Mirrors VBA class module: scenario.cls
    """

    number: int
    name: str
    settle_date: date
    initial_aaa_margin: float  # in bps (e.g. 115)
    aaa_margin_shock: float  # in bps (e.g. -75)
    prepay_speed: int = 15  # CPR — overridden at runtime from config or scenario setup
    horizon_given_type: str = "DISC_MARGIN"


@dataclass
class Cashflow:
    """Represents one period of cashflows (one row from a Cashflows report).

    Mirrors VBA class module: cashflow.cls
    """

    period: int
    cf_date: date
    principal: float = 0.0
    interest: float = 0.0
    balance: float = 0.0


@dataclass
class CashflowsReport:
    """Encapsulates an Intex Cashflows report exported to an Excel worksheet.

    Mirrors VBA class module: cashflows_report.cls
    """

    ws_name: str = ""
    deal_name: str = ""
    tranche: str = ""
    cusip: str = ""
    bloomberg_ticker: str = ""
    header_row: int = 0
    call_type: str = ""  # User Comment
    scenario_name: str = ""  # User Comment 2
    orig_face: float = 0.0
    factor: float = 0.0
    orig_principal: float = 0.0
    cashflows: dict[date, Cashflow] = field(default_factory=dict)

    def load_data(self, ws) -> None:
        """Load data from an openpyxl worksheet.

        Reads descriptive fields from the assumptions section, then reads the
        cashflow data rows into the cashflows dictionary.
        """
        row_num = 1
        all_fields_loaded = False

        while not all_fields_loaded and row_num < 1000:
            cell_val = ws.cell(row=row_num, column=1).value
            if cell_val == "Orig Face:":
                self.orig_face = float(ws.cell(row=row_num, column=2).value)
            elif cell_val == "Factor":
                self.factor = float(ws.cell(row=row_num, column=2).value)
            elif cell_val == "Deal Name:":
                self.deal_name = str(ws.cell(row=row_num, column=2).value)
            elif cell_val == "Tranche":
                self.tranche = str(ws.cell(row=row_num, column=2).value)
            elif cell_val == "CUSIP":
                self.cusip = str(ws.cell(row=row_num, column=2).value)
            elif cell_val == "Bloomberg Ticker":
                self.bloomberg_ticker = str(ws.cell(row=row_num, column=2).value or "")
            elif cell_val == "Period":
                self.header_row = row_num
            elif cell_val == "User Comment":
                self.call_type = str(ws.cell(row=row_num, column=4).value or "")
            elif cell_val == "User Comment 2":
                self.scenario_name = str(ws.cell(row=row_num, column=4).value or "")
                all_fields_loaded = True
            row_num += 1

        if not all_fields_loaded:
            print(f"  WARNING: Not all fields loaded from '{ws.title}'.")

        # Attempt scenario name from worksheet name if blank.
        if not self.scenario_name:
            print(f"  Scenario name blank for '{ws.title}', attempting extraction from worksheet name...")
            self.scenario_name = _map_ws_scenario_name(
                _get_scenario_name_from_ws_name(ws.title)
            )

        self.ws_name = ws.title
        self.orig_principal = self.orig_face * self.factor

        # Build column index from header row.
        col_map = {}
        col = 1
        while ws.cell(row=self.header_row, column=col).value is not None:
            col_map[ws.cell(row=self.header_row, column=col).value] = col
            col += 1

        # Load cashflows starting 3 rows below header.
        row_num = self.header_row + 3
        while ws.cell(row=row_num, column=1).value is not None:
            cf = Cashflow(
                period=int(ws.cell(row=row_num, column=col_map["Period"]).value),
                cf_date=_to_date(ws.cell(row=row_num, column=col_map["Date"]).value),
                interest=float(ws.cell(row=row_num, column=col_map["Interest"]).value),
                principal=float(ws.cell(row=row_num, column=col_map["Principal"]).value),
                balance=float(ws.cell(row=row_num, column=col_map["Balance"]).value),
            )
            self.cashflows[cf.cf_date] = cf
            row_num += 1


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _to_date(val) -> date:
    """Convert a value to a date object."""
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y%m%d"):
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                continue
    raise ValueError(f"Cannot convert {val!r} to date")


def _get_scenario_name_from_ws_name(ws_name: str) -> str:
    """Deduce scenario name from an Intex Cashflows worksheet name.

    Handles formats like: "DEAL,TR, 75 bps Tightening" or "DEAL-Scenario 1-75 bps Tightening"
    """
    # Try comma-delimited format first.
    if "," in ws_name:
        raw = ws_name.rsplit(",", 1)[-1].strip()
    elif "-" in ws_name:
        raw = ws_name.rsplit("-", 1)[-1].strip()
    else:
        return ws_name

    if "Flat" in raw:
        return "Flat"
    if raw.startswith("CCC"):
        return "CCC_Stress"
    if raw.startswith("High"):
        return "High_Prepay"

    parts = raw.split()
    if len(parts) >= 3:
        num_bps = f"{parts[0]} {parts[1]}"
        direction_letter = parts[2][0].upper()
        if direction_letter == "T":
            return f"{num_bps} Tightening"
        elif direction_letter == "W":
            return f"{num_bps} Widening"

    return raw


def _map_ws_scenario_name(ws_scen_name: str) -> str:
    """Map a scenario name like '75 bps Tightening' to 'Scenario N'.

    This requires access to the scenario setup, which is handled at the
    orchestrator level. Here we return the name as-is; the caller can map it.
    """
    return ws_scen_name
