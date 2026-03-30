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
    wait_seconds: int = 120,
) -> pd.DataFrame:
    """Enrich holdings with Intex/BBG data via xlwings (fully automated).

    Uses xlwings to:
      1. Create an Excel workbook with IntexLINK and Bloomberg BDP formulas.
      2. Open it in Excel (with add-ins active) and wait for formulas to calculate.
      3. Read the calculated values back into Python.
      4. Close Excel.

    No manual intervention required — as long as Excel, IntexLINK, and Bloomberg
    Terminal are running on the machine.

    For pre-price CUSIPs, Intex Name and Deal Name are written as static values
    from the Pre-Price Deals table.

    Args:
        holdings_df: Raw holdings DataFrame (from any loading method).
        preprice_deals: PrePriceDeals lookup for overriding Intex identifiers.
        output_path: Path to save the enrichment workbook.
        wait_seconds: Max seconds to wait for formulas to finish calculating.

    Returns:
        The original holdings DataFrame with Intex/BBG columns merged in.
    """
    import time
    import xlwings as xw

    output_path = Path(output_path).resolve()

    # Deduplicate to unique CUSIPs (sum face/par).
    agg_cols = {}
    for col in holdings_df.columns:
        if col == "CUSIP":
            continue
        elif col in ("Original Face", "Current Par", "Book Value", "Market Value"):
            agg_cols[col] = "sum"
        else:
            agg_cols[col] = "first"
    unique_df = holdings_df.groupby("CUSIP", as_index=False).agg(agg_cols)

    print(f"Enriching {len(unique_df)} CUSIPs via xlwings...")

    # Define column layout.
    base_cols = ["CUSIP", "Description", "Entity Name", "PAM Portfolio",
                 "Original Face", "Current Par", "Price", "OAS", "Floater"]
    intex_cols = ["Intex Name", "Intex Deal Name", "AAA Margin",
                  "Non-Call End", "Reinvest End", "Orig Deal Balance"]
    bbg_cols = ["BBG_NC_END", "BBG_REINVEST_END", "RESET_IDX", "COLLAT_TYP"]
    all_cols = base_cols + intex_cols + bbg_cols

    intex_name_col_idx = len(base_cols) + 1  # 1-indexed column for "Intex Name"

    # Open Excel via xlwings (visible so add-ins load properly).
    app = xw.App(visible=True)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Holdings"

        # Write headers.
        ws.range("A1").value = all_cols

        # Write base data.
        for row_idx, (_, row) in enumerate(unique_df.iterrows(), 2):
            for col_idx, col_name in enumerate(base_cols):
                val = row.get(col_name)
                if pd.notna(val):
                    ws.cells(row_idx, col_idx + 1).value = val

        # Write formulas.
        for row_idx, (_, row) in enumerate(unique_df.iterrows(), 2):
            cusip = str(row.get("CUSIP", ""))
            cusip_ref = f"A{row_idx}"
            intex_name_ref = f"{_col_letter(intex_name_col_idx)}{row_idx}"

            # Intex Name and Deal Name.
            if preprice_deals.get_identifier(cusip) != "N/A":
                ws.cells(row_idx, intex_name_col_idx).value = preprice_deals.get_name(cusip)
                ws.cells(row_idx, intex_name_col_idx + 1).value = preprice_deals.get_deal_name(cusip)
            else:
                ws.cells(row_idx, intex_name_col_idx).formula = f'=INTEX({cusip_ref},"INTEX_DEAL")'
                ws.cells(row_idx, intex_name_col_idx + 1).formula = f'=INTEX({cusip_ref},"DEAL_DEALNAME")'

            # AAA Margin, dates, deal balance (reference Intex Name cell).
            ws.cells(row_idx, intex_name_col_idx + 2).formula = f'=INTEX({intex_name_ref},"INTXDA_TRBLOCK_FLOAT_MARGIN[AAA]")'
            ws.cells(row_idx, intex_name_col_idx + 3).formula = f'=INTEX({intex_name_ref},"DEAL_CALLABLE_AS_OF_DATE")'
            ws.cells(row_idx, intex_name_col_idx + 4).formula = f'=INTEX({intex_name_ref},"DEAL_REINV_END_DATE")'
            ws.cells(row_idx, intex_name_col_idx + 5).formula = f'=INTEX({intex_name_ref},"DEAL_ORIGBAL")'

            # Bloomberg fields.
            bbg_start = intex_name_col_idx + len(intex_cols)
            bbg_cusip = f'{cusip_ref}&" CUSIP"'
            ws.cells(row_idx, bbg_start).formula = f'=BDP({bbg_cusip},"MTG_DEAL_CALL_DT")'
            ws.cells(row_idx, bbg_start + 1).formula = f'=BDP({bbg_cusip},"REINVEST_END_DATE")'
            ws.cells(row_idx, bbg_start + 2).formula = f'=BDP({bbg_cusip},"RESET_IDX")'
            ws.cells(row_idx, bbg_start + 3).formula = f'=BDP({bbg_cusip},"COLLAT_TYP")'

        print(f"  Formulas written. Waiting for calculation (up to {wait_seconds}s)...")

        # Wait for Excel to finish calculating.
        app.calculation = "automatic"
        elapsed = 0
        poll_interval = 5
        while elapsed < wait_seconds:
            time.sleep(poll_interval)
            elapsed += poll_interval
            # Check if any cells still show the calculating placeholder.
            try:
                app.calculate()
                # Test a formula cell for a non-pre-price CUSIP.
                test_val = ws.cells(2, intex_name_col_idx + 2).value  # AAA Margin
                if test_val is not None and not isinstance(test_val, str):
                    # Spot-check the last row too.
                    last_row = len(unique_df) + 1
                    last_val = ws.cells(last_row, intex_name_col_idx + 2).value
                    if last_val is not None and not isinstance(last_val, str):
                        print(f"  Formulas calculated ({elapsed}s elapsed).")
                        break
            except Exception:
                pass
            if elapsed % 30 == 0:
                print(f"  Still waiting... ({elapsed}s)")
        else:
            print(f"  WARNING: Timed out after {wait_seconds}s. Some formulas may not have calculated.")

        # Save the workbook.
        wb.save(str(output_path))
        print(f"  Saved to '{output_path.name}'.")

        # Read the calculated values back.
        data = ws.range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value

    finally:
        wb.close()
        app.quit()

    # Merge enriched columns back onto the full (multi-position) holdings.
    enriched_df = data
    merge_cols = intex_cols + bbg_cols
    present_cols = [c for c in merge_cols if c in enriched_df.columns]

    drop_cols = [c for c in present_cols if c in holdings_df.columns]
    result = holdings_df.drop(columns=drop_cols, errors="ignore")

    enriched_subset = enriched_df[["CUSIP"] + present_cols].drop_duplicates(subset=["CUSIP"])
    result = result.merge(enriched_subset, on="CUSIP", how="left")

    populated = result[present_cols].notna().all(axis=1).sum()
    print(f"  {populated}/{len(result)} positions fully enriched.")

    return result


def load_enriched_holdings(
    enriched_path: str | Path,
    original_holdings_df: pd.DataFrame,
) -> pd.DataFrame:
    """Read back a previously saved enrichment file.

    Use this if you already have an enrichment file from a prior run
    (or if you manually opened and saved one). Merges the Intex/BBG columns
    back onto the full (non-deduplicated) holdings DataFrame.

    Args:
        enriched_path: Path to the enrichment workbook (with calculated values).
        original_holdings_df: The original holdings DataFrame (with all positions).

    Returns:
        Holdings DataFrame with Intex/BBG columns merged in.
    """
    enriched_path = Path(enriched_path)
    print(f"Loading enriched data from '{enriched_path.name}'...")

    enriched_df = pd.read_excel(enriched_path, sheet_name="Holdings", engine="openpyxl")

    merge_cols = ["Intex Name", "Intex Deal Name", "AAA Margin",
                  "Non-Call End", "Reinvest End", "Orig Deal Balance",
                  "BBG_NC_END", "BBG_REINVEST_END", "RESET_IDX", "COLLAT_TYP"]
    present_cols = [c for c in merge_cols if c in enriched_df.columns]

    drop_cols = [c for c in present_cols if c in original_holdings_df.columns]
    result = original_holdings_df.drop(columns=drop_cols, errors="ignore")

    enriched_subset = enriched_df[["CUSIP"] + present_cols].drop_duplicates(subset=["CUSIP"])
    result = result.merge(enriched_subset, on="CUSIP", how="left")

    populated = result[present_cols].notna().all(axis=1).sum()
    print(f"  {populated}/{len(result)} positions fully enriched.")

    return result


def _col_letter(col_num: int) -> str:
    """Convert a 1-indexed column number to an Excel column letter."""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
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
