"""
Holdings import and export.

Mirrors VBA worksheet module: initial_holdings.cls

Step 1: Import CLO holdings from a Structured Products Data Packet report.
Step 2: Enrich with Intex/Bloomberg data (in VBA this wrote formulas; here we
        read pre-populated values from the workbook).
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from forecast.models import Tranche
from forecast.preprice import PrePriceDeals


def _to_date_or_none(val) -> Optional[date]:
    """Safely convert a value to date, returning None on failure."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return pd.Timestamp(val).date()
    except Exception:
        return None


def import_holdings(
    data_packet_path: str | Path,
    fhlb_only: bool = False,
) -> pd.DataFrame:
    """Import CLO holdings from a Structured Products Data Packet report.

    This replaces the VBA ImportHoldings() method. It reads the Data Packet
    workbook and filters for CLO positions, optionally restricting to FHLB
    (PAM Portfolio 13091).

    Args:
        data_packet_path: Path to the Structured Products Data Packet .xlsx file.
        fhlb_only: If True, import only FHLB CLO holdings (PAM Portfolio 13091).

    Returns:
        DataFrame with one row per position.
    """
    data_packet_path = Path(data_packet_path)
    print(f"Importing holdings from '{data_packet_path.name}'...")

    # Read the Holding Detail sheet from the Data Packet.
    # The Data Packet has a "Holding Detail" sheet with headers in row 1.
    df = pd.read_excel(data_packet_path, sheet_name="Holding Detail", engine="openpyxl")

    # Standardize column names — the Data Packet may use slightly different headers.
    # Map known Data Packet columns to our internal names.
    col_map = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if col_lower in ("primary security id", "primary_security_id", "cusip"):
            col_map[col] = "CUSIP"
        elif col_lower in ("security description", "security_description", "description"):
            col_map[col] = "Description"
        elif col_lower in ("client level 2", "client_level_2"):
            col_map[col] = "Client Level 2"
        elif col_lower in ("client level 3", "client_level_3"):
            col_map[col] = "Client Level 3"
        elif col_lower in ("entity name", "entity_name"):
            col_map[col] = "Entity Name"
        elif col_lower in ("pam portfolio", "pam_portfolio"):
            col_map[col] = "PAM Portfolio"
        elif col_lower in ("original face", "original_face"):
            col_map[col] = "Original Face"
        elif col_lower in ("par", "current par"):
            col_map[col] = "Current Par"
        elif col_lower in ("gaap book value", "gaap_bv", "book value"):
            col_map[col] = "Book Value"
        elif col_lower in ("market value", "market_value"):
            col_map[col] = "Market Value"
        elif col_lower == "price":
            col_map[col] = "Price"
        elif col_lower == "oas":
            col_map[col] = "OAS"
        elif col_lower in ("s&p", "sp_rating"):
            col_map[col] = "S&P"
        elif col_lower in ("moody's", "moodys_rating"):
            col_map[col] = "Moody's"
        elif col_lower in ("fitch", "fitch_rating"):
            col_map[col] = "Fitch"
        elif col_lower in ("internal", "internal_rating"):
            col_map[col] = "Internal"
        elif col_lower == "floater":
            col_map[col] = "Floater"
        elif col_lower in ("portfolio view level 4", "portfolio_view_level_4"):
            col_map[col] = "Portfolio View Level 4"

    if col_map:
        df = df.rename(columns=col_map)

    # Filter for CLO positions.
    if "Portfolio View Level 4" in df.columns:
        df = df[df["Portfolio View Level 4"] == "CLO"].copy()

    if fhlb_only and "PAM Portfolio" in df.columns:
        df = df[df["PAM Portfolio"] == 13091].copy()

    print(f"  {len(df)} CLO positions imported.")
    return df.reset_index(drop=True)


def load_holdings_from_forecast_workbook(
    forecast_wb_path: str | Path,
) -> pd.DataFrame:
    """Load holdings directly from the Forecast workbook's 'Initial Holdings' sheet.

    This is an alternative to importing from a Data Packet — useful when the
    Forecast workbook already has populated holdings data.
    """
    forecast_wb_path = Path(forecast_wb_path)
    print(f"Loading holdings from '{forecast_wb_path.name}' Initial Holdings sheet...")

    # Use pandas read_excel which is much faster than cell-by-cell openpyxl reads.
    # Header row is row 7 (0-indexed: skiprows=6), data starts at row 8.
    df = pd.read_excel(
        forecast_wb_path,
        sheet_name="Initial Holdings",
        header=6,  # Row 7 is the header (0-indexed row 6)
        engine="openpyxl",
    )

    # Drop rows where CUSIP is NaN (empty rows at end of table).
    if "CUSIP" in df.columns:
        df = df.dropna(subset=["CUSIP"])

    print(f"  {len(df)} holdings rows loaded.")
    return df


def export_tranches(
    holdings_df: pd.DataFrame,
    preprice_deals: PrePriceDeals,
) -> list[Tranche]:
    """Convert holdings DataFrame to a list of Tranche objects (one per unique CUSIP).

    Sums Original Face and Current Par across positions sharing the same CUSIP.
    Looks up pre-price deal info when available.

    Mirrors VBA method: initial_holdings.ExportTranches()
    """
    print("Exporting tranches from holdings...")
    tranches: dict[str, Tranche] = {}

    for _, row in holdings_df.iterrows():
        cusip = str(row.get("CUSIP", ""))
        if not cusip:
            continue

        if cusip in tranches:
            # Sum face and par for duplicate CUSIPs.
            tranches[cusip].orig_face += float(row.get("Original Face", 0) or 0)
            tranches[cusip].cur_par += float(row.get("Current Par", 0) or 0)
            continue

        t = Tranche(cusip=cusip)
        t.orig_face = float(row.get("Original Face", 0) or 0)
        t.cur_par = float(row.get("Current Par", 0) or 0)
        t.price = float(row.get("Price", 0) or 0)
        t.oas = float(row.get("OAS", 0) or 0)
        t.aaa_margin = float(row.get("AAA Margin", 0) or 0)
        t.non_call_end = _to_date_or_none(row.get("BBG_NC_END") or row.get("Non-Call End"))
        t.reinvest_end = _to_date_or_none(row.get("BBG_REINVEST_END") or row.get("Reinvest End"))
        t.orig_deal_balance = float(row.get("Orig Deal Balance", 0) or 0)
        t.reset_index = str(row.get("RESET_IDX", "") or "")
        t.floater = str(row.get("Floater", "")).upper() == "Y"
        t.middle_market = str(row.get("COLLAT_TYP", "")).upper() == "CF-CLO-MML"

        # Check pre-price deals for Intex info.
        if preprice_deals.get_identifier(cusip) != "N/A":
            t.intex_name = preprice_deals.get_name(cusip)
            t.intex_deal_name = preprice_deals.get_deal_name(cusip)
            t.preprice_password = preprice_deals.get_password(cusip)
            t.preprice = True
        else:
            t.intex_name = str(row.get("Intex Name", "") or "")
            t.intex_deal_name = str(row.get("Intex Deal Name", "") or "")
            t.preprice = False

        tranches[cusip] = t

    result = list(tranches.values())
    print(f"  {len(result)} unique tranches exported.")
    return result
