"""
Bloomberg data retrieval via the blpapi Python SDK.

Replaces the Excel BDP() formulas with direct Python API calls.
Requires Bloomberg Terminal running and the blpapi package installed
with the C++ SDK.
"""

from __future__ import annotations

from datetime import date

import pandas as pd


# Bloomberg fields we need for the forecast.
# Keys = Bloomberg field names, Values = our column names.
BBG_FIELDS = {
    "MTG_DEAL_CALL_DT": "BBG_NC_END",        # Non-call end date
    "REINVEST_END_DATE": "BBG_REINVEST_END",  # Reinvestment end date
    "RESET_IDX": "RESET_IDX",                 # Rate reset index (SOFR vs LIBOR)
    "COLLAT_TYP": "COLLAT_TYP",              # Collateral type (BSL vs MML)
}


def fetch_bbg_data(cusips: list[str], timeout_ms: int = 60000) -> pd.DataFrame:
    """Fetch Bloomberg reference data for a list of CUSIPs.

    Uses blpapi directly (not xbbg) for reliability.

    Args:
        cusips: List of CUSIP strings.
        timeout_ms: Timeout in milliseconds for the Bloomberg request.

    Returns:
        DataFrame with columns: CUSIP, BBG_NC_END, BBG_REINVEST_END, RESET_IDX, COLLAT_TYP
    """
    import blpapi

    tickers = [f"/cusip/{cusip}" for cusip in cusips]
    fields = list(BBG_FIELDS.keys())

    print(f"  Requesting {len(fields)} fields for {len(cusips)} CUSIPs from Bloomberg...")

    # Connect.
    session_options = blpapi.SessionOptions()
    session_options.setServerHost("localhost")
    session_options.setServerPort(8194)

    session = blpapi.Session(session_options)
    if not session.start():
        raise ConnectionError("Could not start Bloomberg session.")

    if not session.openService("//blp/refdata"):
        session.stop()
        raise ConnectionError("Could not open //blp/refdata service.")

    try:
        service = session.getService("//blp/refdata")
        request = service.createRequest("ReferenceDataRequest")

        for t in tickers:
            request.append("securities", t)
        for f in fields:
            request.append("fields", f)

        session.sendRequest(request)

        # Collect responses.
        results: dict[str, dict[str, object]] = {c: {} for c in cusips}
        done = False

        while not done:
            event = session.nextEvent(timeout_ms)
            event_type = event.eventType()

            if event_type in (blpapi.Event.RESPONSE, blpapi.Event.PARTIAL_RESPONSE):
                for msg in event:
                    if not msg.hasElement("securityData"):
                        continue
                    security_data = msg.getElement("securityData")
                    for i in range(security_data.numValues()):
                        sec = security_data.getValueAsElement(i)
                        sec_name = sec.getElementAsString("security")

                        # Extract CUSIP from "/cusip/XXXXXXXXX".
                        cusip = sec_name.rsplit("/", 1)[-1].strip()

                        if sec.hasElement("securityError"):
                            continue

                        if not sec.hasElement("fieldData"):
                            continue

                        fd = sec.getElement("fieldData")
                        for bbg_field, our_col in BBG_FIELDS.items():
                            if fd.hasElement(bbg_field):
                                el = fd.getElement(bbg_field)
                                if not el.isNull():
                                    results[cusip][our_col] = _to_python(el)

            if event_type == blpapi.Event.RESPONSE:
                done = True

    finally:
        session.stop()

    # Build DataFrame.
    rows = [{"CUSIP": c, **results.get(c, {})} for c in cusips]
    df = pd.DataFrame(rows)

    for col in BBG_FIELDS.values():
        if col in df.columns:
            n = df[col].notna().sum()
            print(f"    {col}: {n}/{len(cusips)} populated")
        else:
            print(f"    {col}: no data returned")

    return df


def _to_python(element) -> object:
    """Convert a blpapi Element to a Python value."""
    import blpapi

    dtype = element.datatype()

    # Date
    if dtype == blpapi.DataType.DATE:
        v = element.getValueAsDatetime()
        try:
            return date(v.year, v.month, v.day)
        except Exception:
            return None

    # String
    if dtype in (blpapi.DataType.STRING, blpapi.DataType.BYTEARRAY):
        return element.getValueAsString()

    # Numeric
    if dtype in (blpapi.DataType.BOOL,):
        return element.getValueAsBool()
    if dtype in (blpapi.DataType.INT32, blpapi.DataType.INT64):
        return element.getValueAsInteger()
    if dtype in (blpapi.DataType.FLOAT32, blpapi.DataType.FLOAT64):
        return element.getValueAsFloat()

    # Fallback
    try:
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

    populated = result[merge_cols].notna().all(axis=1).sum() if merge_cols else 0
    print(f"  {populated}/{len(result)} positions have complete Bloomberg data.")

    return result
