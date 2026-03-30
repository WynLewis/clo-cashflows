"""
Bloomberg data retrieval via xbbg.

Replaces the Excel BDP() formulas with direct Python API calls.
Requires Bloomberg Terminal running and the xbbg package installed:
    pip install xbbg
"""

from __future__ import annotations

import pandas as pd


# Bloomberg fields we need for the forecast.
# Keys = Bloomberg field names, Values = our column names.
BBG_FIELDS = {
    "MTG_DEAL_CALL_DT": "BBG_NC_END",        # Non-call end date
    "REINVEST_END_DATE": "BBG_REINVEST_END",  # Reinvestment end date
    "RESET_IDX": "RESET_IDX",                 # Rate reset index (SOFR vs LIBOR)
    "COLLAT_TYP": "COLLAT_TYP",              # Collateral type (BSL vs MML)
}


def fetch_bbg_data(cusips: list[str]) -> pd.DataFrame:
    """Fetch Bloomberg reference data for a list of CUSIPs using xbbg.

    Makes a single bulk BDP call for all CUSIPs and fields.

    Args:
        cusips: List of CUSIP strings.

    Returns:
        DataFrame with columns: CUSIP, BBG_NC_END, BBG_REINVEST_END, RESET_IDX, COLLAT_TYP
    """
    from xbbg import blp

    # Build security identifiers (Bloomberg expects "/cusip/XXXXXXXXX" or "XXXXXXXXX CUSIP").
    tickers = [f"/cusip/{cusip}" for cusip in cusips]
    fields = list(BBG_FIELDS.keys())

    print(f"  Requesting {len(fields)} fields for {len(cusips)} CUSIPs from Bloomberg...")

    # xbbg.blp.bdp returns a DataFrame with tickers as index and fields as columns.
    raw = blp.bdp(tickers=tickers, flds=fields)

    print(f"  Received {len(raw)} rows from Bloomberg.")

    # Map the index back to CUSIPs and rename columns.
    # xbbg index is the ticker string; extract the CUSIP from it.
    rows = []
    for ticker, data in raw.iterrows():
        # Extract CUSIP from "/cusip/XXXXXXXXX" or "XXXXXXXXX CUSIP" format.
        if "/cusip/" in str(ticker).lower():
            cusip = str(ticker).split("/")[-1].strip()
        else:
            cusip = str(ticker).split()[0].strip()

        row = {"CUSIP": cusip}
        for bbg_field, our_col in BBG_FIELDS.items():
            # xbbg lowercases field names in the DataFrame columns.
            val = data.get(bbg_field) or data.get(bbg_field.lower())
            if pd.notna(val):
                row[our_col] = val
        rows.append(row)

    df = pd.DataFrame(rows)

    # Summary.
    for col in BBG_FIELDS.values():
        if col in df.columns:
            populated = df[col].notna().sum()
            print(f"    {col}: {populated}/{len(cusips)} populated")
        else:
            print(f"    {col}: column missing from response")

    return df


def enrich_with_bbg(holdings_df: pd.DataFrame) -> pd.DataFrame:
    """Add Bloomberg fields to a holdings DataFrame.

    Fetches BBG_NC_END, BBG_REINVEST_END, RESET_IDX, COLLAT_TYP for
    each unique CUSIP and merges them onto the holdings.

    Args:
        holdings_df: Holdings DataFrame with a CUSIP column.

    Returns:
        Holdings DataFrame with Bloomberg columns added/updated.
    """
    cusips = holdings_df["CUSIP"].dropna().unique().tolist()
    if not cusips:
        print("  No CUSIPs to fetch Bloomberg data for.")
        return holdings_df

    bbg_df = fetch_bbg_data(cusips)

    # Drop existing BBG columns if present (will be replaced).
    bbg_cols = list(BBG_FIELDS.values())
    drop_cols = [c for c in bbg_cols if c in holdings_df.columns]
    result = holdings_df.drop(columns=drop_cols, errors="ignore")

    # Merge by CUSIP.
    merge_cols = [c for c in bbg_cols if c in bbg_df.columns]
    if merge_cols:
        result = result.merge(
            bbg_df[["CUSIP"] + merge_cols].drop_duplicates(subset=["CUSIP"]),
            on="CUSIP",
            how="left",
        )

    populated = result[merge_cols].notna().all(axis=1).sum() if merge_cols else 0
    print(f"  {populated}/{len(result)} positions have complete Bloomberg data.")

    return result
