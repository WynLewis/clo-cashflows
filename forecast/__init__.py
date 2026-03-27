"""
CLO Cashflow Forecast - Python migration of Forecast v16.3.xlsm

This package replaces the VBA macro workbook with Python modules that implement
the same 7-step CLO cashflow forecasting workflow:

  Step 1: Import CLO holdings from a Structured Products Data Packet report.
  Step 2: Enrich holdings with Intex/Bloomberg data (manual or API-based).
  Step 3: Clear forecasted factors.
  Step 4: Generate preliminary Intex Portfolio Upload file (with COLLAT tranches).
  Step 5: Extract forecasted factor call dates from preliminary cashflows.
  Step 6: Generate final Intex Portfolio Upload file (with factor call dates).
  Step 7: Summarize final cashflows by scenario.
"""
