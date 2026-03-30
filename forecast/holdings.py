"""
Holdings import and export.

Mirrors VBA worksheet module: initial_holdings.cls

Step 1: Import CLO holdings from the structured_products library (default),
        a Structured Products Data Packet report, or the Forecast workbook.
Step 2: Enrich with Intex/Bloomberg data (in VBA this wrote formulas; here we
        read pre-populated values from the workbook).
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from forecast.config import CONFIG
from forecast.models import Tranche
from forecast.preprice import PrePriceDeals


def load_holdings_from_clo_library() -> pd.DataFrame:
    """Load CLO holdings via the structured_products.clo library.

    This is the **default** method — it pulls live portfolio holdings from
    the structured_products library (same source as the holdings_pricing notebook).

    Requires the structured_products package to be on PYTHONPATH.
    If not available, falls back gracefully with a clear error.

    Returns:
        DataFrame with one row per position.
    """
    try:
        from structured_products import clo
    except ImportError:
        raise ImportError(
            "The 'structured_products' package is not available. "
            "Either add it to your PYTHONPATH (e.g. sys.path.append('C:/ActData/Python/files/OOI/structured-products-main/')) "
            "or use import_holdings() with a Data Packet file, "
            "or load_holdings_from_forecast_workbook() instead."
        )

    print("Loading CLO holdings from structured_products.clo library...")
    holdings = clo.holdings(type='portfolio')

    as_of_date = pd.to_datetime(holdings.iloc[0]["As_at_Date"]).date()
    print(f"  Data as of: {as_of_date.strftime('%B %d, %Y')}")
    print(f"  {len(holdings)} positions loaded")
    print(f"  {holdings['Primary Security ID'].nunique()} unique securities")
    print(f"  ${holdings['GAAP BV'].sum():,.2f} total GAAP book value")

    # Standardize column names to match the Forecast workbook convention.
    col_map = {
        "Primary Security ID": "CUSIP",
        "Security Description": "Description",
        "Client Level 2": "Client Level 2",
        "Client Level 3": "Client Level 3",
        "Entity Name": "Entity Name",
        "PAM Portfolio": "PAM Portfolio",
        "Original Face": "Original Face",
        "Par": "Current Par",
        "GAAP BV": "Book Value",
        "Market Value": "Market Value",
        "Price": "Price",
        "OAS": "OAS",
        "S&P Rating": "S&P",
        "Moodys Rating": "Moody's",
        "Fitch Rating": "Fitch",
        "Internal Rating": "Internal",
        "Floater": "Floater",
        "Portfolio View Level 4": "Portfolio View Level 4",
    }
    # Only rename columns that exist in the DataFrame.
    rename_map = {k: v for k, v in col_map.items() if k in holdings.columns}
    holdings = holdings.rename(columns=rename_map)

    # Filter for CLO positions only.
    if "Portfolio View Level 4" in holdings.columns:
        holdings = holdings[holdings["Portfolio View Level 4"] == "CLO"].copy()
        print(f"  {len(holdings)} CLO positions after filtering")

    return holdings.reset_index(drop=True)


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
) -> pd.DataFrame:
    """Import CLO holdings from a Structured Products Data Packet report.

    This replaces the VBA ImportHoldings() method. It reads the Data Packet
    workbook and filters for CLO positions.

    Args:
        data_packet_path: Path to the Structured Products Data Packet .xlsx file.

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


def enrich_holdings(
    holdings_df: pd.DataFrame,
    preprice_deals: PrePriceDeals,
    output_path: str | Path,
) -> None:
    """Write holdings to an Excel file with IntexLINK and Bloomberg BDP formulas.

    This bridges the gap between Python and the Intex/Bloomberg Excel add-ins.
    The workflow is:
        1. Python writes this file with formulas in the Intex/BBG columns.
        2. You open it in Excel (with IntexLINK and Bloomberg add-ins active).
        3. Formulas calculate and populate the data.
        4. Save the file.
        5. Python reads it back with load_enriched_holdings().

    For pre-price CUSIPs, the Intex Name and Deal Name are written as static
    values (from the Pre-Price Deals table) instead of formulas.

    Args:
        holdings_df: Raw holdings DataFrame (from any loading method).
        preprice_deals: PrePriceDeals lookup for overriding Intex identifiers.
        output_path: Path to write the enrichment workbook.
    """
    import openpyxl
    from openpyxl.utils import get_column_letter

    output_path = Path(output_path)
    print(f"Writing enrichment workbook to '{output_path.name}'...")

    # Deduplicate to unique CUSIPs (sum face/par).
    group_cols = ["CUSIP"]
    agg_cols = {}
    for col in holdings_df.columns:
        if col == "CUSIP":
            continue
        elif col in ("Original Face", "Current Par", "Book Value", "Market Value"):
            agg_cols[col] = "sum"
        else:
            agg_cols[col] = "first"
    unique_df = holdings_df.groupby("CUSIP", as_index=False).agg(agg_cols)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Holdings"

    # Define columns: existing data + Intex/BBG formula columns.
    base_cols = ["CUSIP", "Description", "Entity Name", "PAM Portfolio",
                 "Original Face", "Current Par", "Price", "OAS", "Floater"]
    intex_cols = ["Intex Name", "Intex Deal Name", "AAA Margin",
                  "Non-Call End", "Reinvest End", "Orig Deal Balance"]
    bbg_cols = ["BBG_NC_END", "BBG_REINVEST_END", "RESET_IDX", "COLLAT_TYP"]
    all_cols = base_cols + intex_cols + bbg_cols

    # Write headers.
    for col_idx, col_name in enumerate(all_cols, 1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Write data rows with formulas.
    for row_idx, (_, row) in enumerate(unique_df.iterrows(), 2):
        cusip = str(row.get("CUSIP", ""))

        # Write base data columns.
        for col_idx, col_name in enumerate(base_cols, 1):
            val = row.get(col_name)
            if pd.notna(val):
                ws.cell(row=row_idx, column=col_idx, value=val)

        cusip_col = "A"  # CUSIP is always column A
        cusip_ref = f"${cusip_col}${row_idx}"

        # -- Intex columns --
        intex_name_col_idx = base_cols.index("CUSIP") + len(base_cols) + 1  # First intex col
        intex_name_col = get_column_letter(intex_name_col_idx)
        intex_deal_col = get_column_letter(intex_name_col_idx + 1)

        if preprice_deals.get_identifier(cusip) != "N/A":
            # Pre-price: write static values.
            ws.cell(row=row_idx, column=intex_name_col_idx,
                    value=preprice_deals.get_name(cusip))
            ws.cell(row=row_idx, column=intex_name_col_idx + 1,
                    value=preprice_deals.get_deal_name(cusip))
        else:
            # Regular: write IntexLINK formulas.
            ws.cell(row=row_idx, column=intex_name_col_idx,
                    value=f'=INTEX({cusip_ref},"INTEX_DEAL")')
            ws.cell(row=row_idx, column=intex_name_col_idx + 1,
                    value=f'=INTEX({cusip_ref},"DEAL_DEALNAME")')

        # AAA Margin — references the Intex Name cell.
        intex_name_ref = f"${intex_name_col}${row_idx}"
        ws.cell(row=row_idx, column=intex_name_col_idx + 2,
                value=f'=INTEX({intex_name_ref},"INTXDA_TRBLOCK_FLOAT_MARGIN[AAA]")')

        # Non-Call End, Reinvest End, Orig Deal Balance.
        ws.cell(row=row_idx, column=intex_name_col_idx + 3,
                value=f'=INTEX({intex_name_ref},"DEAL_CALLABLE_AS_OF_DATE")')
        ws.cell(row=row_idx, column=intex_name_col_idx + 4,
                value=f'=INTEX({intex_name_ref},"DEAL_REINV_END_DATE")')
        ws.cell(row=row_idx, column=intex_name_col_idx + 5,
                value=f'=INTEX({intex_name_ref},"DEAL_ORIGBAL")')

        # -- Bloomberg columns --
        bbg_start_col_idx = intex_name_col_idx + len(intex_cols)
        bbg_cusip_ref = f'{cusip_ref}&" CUSIP"'

        ws.cell(row=row_idx, column=bbg_start_col_idx,
                value=f'=BDP({bbg_cusip_ref},"MTG_DEAL_CALL_DT")')
        ws.cell(row=row_idx, column=bbg_start_col_idx + 1,
                value=f'=BDP({bbg_cusip_ref},"REINVEST_END_DATE")')
        ws.cell(row=row_idx, column=bbg_start_col_idx + 2,
                value=f'=BDP({bbg_cusip_ref},"RESET_IDX")')
        ws.cell(row=row_idx, column=bbg_start_col_idx + 3,
                value=f'=BDP({bbg_cusip_ref},"COLLAT_TYP")')

    wb.save(output_path)
    print(f"  {len(unique_df)} CUSIPs written with IntexLINK and Bloomberg formulas.")
    print(f"  Next steps:")
    print(f"    1. Open '{output_path.name}' in Excel (with IntexLINK + Bloomberg add-ins)")
    print(f"    2. Wait for all formulas to calculate")
    print(f"    3. Fix any #N/A errors (especially pre-price deals)")
    print(f"    4. Save the file")
    print(f"    5. Come back here and run load_enriched_holdings('{output_path.name}')")


def load_enriched_holdings(
    enriched_path: str | Path,
    original_holdings_df: pd.DataFrame,
) -> pd.DataFrame:
    """Read back an enriched holdings file (after formulas have calculated in Excel).

    Merges the Intex/BBG columns from the enrichment workbook back onto the
    full (non-deduplicated) holdings DataFrame so all positions get the data.

    Args:
        enriched_path: Path to the enrichment workbook (saved after formulas calculated).
        original_holdings_df: The original holdings DataFrame (with all positions).

    Returns:
        Holdings DataFrame with Intex/BBG columns merged in.
    """
    enriched_path = Path(enriched_path)
    print(f"Loading enriched data from '{enriched_path.name}'...")

    enriched_df = pd.read_excel(enriched_path, sheet_name="Holdings", engine="openpyxl")

    # The enriched file has one row per CUSIP.  Merge the Intex/BBG columns
    # onto the original (multi-position) holdings by CUSIP.
    merge_cols = ["Intex Name", "Intex Deal Name", "AAA Margin",
                  "Non-Call End", "Reinvest End", "Orig Deal Balance",
                  "BBG_NC_END", "BBG_REINVEST_END", "RESET_IDX", "COLLAT_TYP"]
    present_cols = [c for c in merge_cols if c in enriched_df.columns]

    # Drop these columns from original if they exist (will be replaced).
    drop_cols = [c for c in present_cols if c in original_holdings_df.columns]
    result = original_holdings_df.drop(columns=drop_cols, errors="ignore")

    # Merge.
    enriched_subset = enriched_df[["CUSIP"] + present_cols].drop_duplicates(subset=["CUSIP"])
    result = result.merge(enriched_subset, on="CUSIP", how="left")

    populated = result[present_cols].notna().all(axis=1).sum()
    total = len(result)
    print(f"  {populated}/{total} positions fully enriched.")

    return result


def clean_holdings(holdings_df: pd.DataFrame) -> pd.DataFrame:
    """Clean holdings data for Intex runs.

    1. Cross-fills missing date columns between BBG and Intex sources:
       - BBG_NC_END <-> Non-Call End (whichever has a value fills the other)
       - BBG_REINVEST_END <-> Reinvest End
    2. Drops any CUSIP that still has NA in any required column.
       For dates, a CUSIP is only dropped if BOTH sources are missing
       (i.e. as long as at least one of BBG or Intex has the date, it's kept).

    Returns a cleaned copy of the DataFrame.
    """
    df = holdings_df.copy()
    initial_cusips = df["CUSIP"].nunique()

    # --- Cross-fill dates between BBG and Intex sources ---
    date_pairs = [
        ("BBG_NC_END", "Non-Call End"),
        ("BBG_REINVEST_END", "Reinvest End"),
    ]
    for bbg_col, intex_col in date_pairs:
        if bbg_col in df.columns and intex_col in df.columns:
            # Fill BBG from Intex where BBG is missing.
            mask = df[bbg_col].isna() & df[intex_col].notna()
            if mask.any():
                df.loc[mask, bbg_col] = df.loc[mask, intex_col]
                print(f"  Filled {mask.sum()} missing {bbg_col} values from {intex_col}.")
            # Fill Intex from BBG where Intex is missing.
            mask = df[intex_col].isna() & df[bbg_col].notna()
            if mask.any():
                df.loc[mask, intex_col] = df.loc[mask, bbg_col]
                print(f"  Filled {mask.sum()} missing {intex_col} values from {bbg_col}.")

    # --- Drop CUSIPs with NA in any required column ---
    # Non-date columns: always required.
    required_cols = ["Intex Name", "Intex Deal Name", "AAA Margin",
                     "Orig Deal Balance", "RESET_IDX", "COLLAT_TYP"]
    # Date columns: only required if BOTH BBG and Intex are missing
    # (after cross-fill above, if one had a value both now do).
    date_cols = ["BBG_NC_END", "BBG_REINVEST_END"]

    check_cols = [c for c in required_cols + date_cols if c in df.columns]

    if check_cols:
        na_mask = df[check_cols].isna().any(axis=1)
        bad_cusips = df.loc[na_mask, "CUSIP"].unique()

        if len(bad_cusips) > 0:
            df = df[~df["CUSIP"].isin(bad_cusips)].copy()
            final_cusips = df["CUSIP"].nunique()
            print(f"  Dropped {len(bad_cusips)} CUSIPs with missing data "
                  f"({initial_cusips} -> {final_cusips} unique CUSIPs).")
            for cusip in sorted(bad_cusips)[:10]:
                cusip_rows = holdings_df[holdings_df["CUSIP"] == cusip]
                missing = [c for c in check_cols if cusip_rows[c].isna().any()]
                print(f"    {cusip}: missing {missing}")
            if len(bad_cusips) > 10:
                print(f"    ... and {len(bad_cusips) - 10} more.")
        else:
            print(f"  All {initial_cusips} CUSIPs have complete Intex/BBG data.")

    return df.reset_index(drop=True)


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
        t.middle_market = str(row.get("COLLAT_TYP", "")).upper() == CONFIG.call.middle_market_collat_type

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
