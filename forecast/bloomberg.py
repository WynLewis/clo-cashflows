"""
Bloomberg data retrieval via the blpapi Python SDK.

Replaces the Excel BDP() formulas with direct Python API calls.
Requires Bloomberg Terminal running and the blpapi package installed:
    pip install blpapi
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Optional

import pandas as pd


# Bloomberg fields we need for the forecast.
BBG_FIELDS = {
    "MTG_DEAL_CALL_DT": "BBG_NC_END",        # Non-call end date
    "REINVEST_END_DATE": "BBG_REINVEST_END",  # Reinvestment end date
    "RESET_IDX": "RESET_IDX",                 # Rate reset index (SOFR vs LIBOR)
    "COLLAT_TYP": "COLLAT_TYP",              # Collateral type (BSL vs MML)
}


def fetch_bbg_data(cusips: list[str], timeout_ms: int = 60000) -> pd.DataFrame:
    """Fetch Bloomberg reference data for a list of CUSIPs.

    Makes a single bulk ReferenceDataRequest for all CUSIPs and fields,
    rather than one request per CUSIP.

    Args:
        cusips: List of CUSIP strings.
        timeout_ms: Timeout in milliseconds for the Bloomberg request.

    Returns:
        DataFrame with columns: CUSIP, BBG_NC_END, BBG_REINVEST_END, RESET_IDX, COLLAT_TYP
    """
    import blpapi

    # Build security identifiers (Bloomberg expects "CUSIP CUSIP" format).
    securities = [f"{cusip} CUSIP" for cusip in cusips]
    fields = list(BBG_FIELDS.keys())

    print(f"  Requesting {len(fields)} fields for {len(cusips)} CUSIPs from Bloomberg...")

    # Connect to Bloomberg.
    session_options = blpapi.SessionOptions()
    session_options.setServerHost("localhost")
    session_options.setServerPort(8194)

    session = blpapi.Session(session_options)
    if not session.start():
        raise ConnectionError(
            "Could not start Bloomberg session. "
            "Ensure Bloomberg Terminal is running and blpapi is installed."
        )

    if not session.openService("//blp/refdata"):
        session.stop()
        raise ConnectionError("Could not open //blp/refdata service.")

    try:
        service = session.getService("//blp/refdata")
        request = service.createRequest("ReferenceDataRequest")

        for sec in securities:
            request.append("securities", sec)
        for fld in fields:
            request.append("fields", fld)

        session.sendRequest(request)

        # Collect responses.
        results: dict[str, dict[str, object]] = {cusip: {} for cusip in cusips}
        done = False

        while not done:
            event = session.nextEvent(timeout_ms)

            if event.eventType() in (blpapi.Event.RESPONSE, blpapi.Event.PARTIAL_RESPONSE):
                for msg in event:
                    security_data = msg.getElement("securityData")
                    for i in range(security_data.numValues()):
                        sec_element = security_data.getValueAsElement(i)
                        sec_name = sec_element.getElementAsString("security")

                        # Extract CUSIP from "CUSIP CUSIP" format.
                        cusip = sec_name.split()[0]

                        # Check for security-level errors.
                        if sec_element.hasElement("securityError"):
                            err = sec_element.getElement("securityError")
                            print(f"    WARNING: {cusip}: {err.getElementAsString('message')}")
                            continue

                        field_data = sec_element.getElement("fieldData")
                        for bbg_field, our_col in BBG_FIELDS.items():
                            if field_data.hasElement(bbg_field):
                                val = _extract_value(field_data, bbg_field)
                                results[cusip][our_col] = val

            if event.eventType() == blpapi.Event.RESPONSE:
                done = True

    finally:
        session.stop()

    # Convert to DataFrame.
    rows = []
    for cusip in cusips:
        row = {"CUSIP": cusip}
        row.update(results.get(cusip, {}))
        rows.append(row)

    df = pd.DataFrame(rows)

    # Summary.
    for col in BBG_FIELDS.values():
        if col in df.columns:
            populated = df[col].notna().sum()
            print(f"    {col}: {populated}/{len(cusips)} populated")

    return df


def _extract_value(field_data, field_name) -> object:
    """Extract a typed value from a Bloomberg field element."""
    try:
        element = field_data.getElement(field_name)
        if element.isNull():
            return None

        # Bloomberg returns dates as blpapi.Datetime objects.
        dtype = element.datatype()

        # Date type
        if dtype == 10:  # BLPAPI_DATATYPE_DATE
            val = element.getValueAsDatetime()
            if hasattr(val, "year"):
                return date(val.year, val.month, val.day)
            return val

        # String type
        if dtype in (8, 9):  # STRING, BYTEARRAY
            return element.getValueAsString()

        # Numeric types
        if dtype in (1, 2, 3, 4, 5, 6, 7):  # various int/float types
            return element.getValueAsFloat()

        # Fallback
        return element.getValueAsString()

    except Exception:
        return None


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

    populated = result[merge_cols].notna().all(axis=1).sum()
    print(f"  {populated}/{len(result)} positions have complete Bloomberg data.")

    return result
