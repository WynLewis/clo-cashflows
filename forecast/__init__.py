"""
CLO Cashflow Forecast - Python migration of Forecast v16.3.xlsm

This package models how CLO portfolio cashflows and balances evolve
under different AAA spread scenarios.  Tighter spreads cause more deals
to be called sooner; wider spreads extend weighted average life.

Modules:
    config          — All tunable parameters (CONFIG singleton)
    models          — Tranche, Scenario, Cashflow, CashflowsReport dataclasses
    holdings        — Import, enrich (Intex/BBG), and clean CLO holdings
    bloomberg       — Bloomberg data via blpapi (replaces Excel BDP formulas)
    preprice        — Pre-price deal lookup (optional)
    scenarios       — Scenario setup, call logic, Intex upload generation
    factors         — Factor call date extraction from COLLAT cashflows
    forward_curve   — Forward rate curve loading
    cashflows_summary — Parse Intex exports, summarize by scenario
    portfolio       — Sub-portfolio allocation (par-weighted by CUSIP)
    reporting       — Balance charts, quarterly rollup, hedging output
    excel_addins    — COM/XLL add-in toggling for xlwings
    file_picker     — Native file browser dialog for selecting files

Workflow (9 steps):
    1. Import CLO holdings (clo library, Data Packet, or Forecast workbook)
    2. Enrich with Bloomberg (blpapi) + IntexLINK (Excel formulas) data
    3. Clear forecasted factors
    4. Generate preliminary Intex upload (COLLAT tranches)
    4b. Run in IntexCalc (manual)
    5. Extract factor call dates from COLLAT cashflows
    6. Generate final Intex upload (actual tranches, call = min(refi, factor))
    6b. Run in IntexCalc (manual)
    7. Summarize cashflows by scenario
    8. Allocate cashflows to sub-portfolios
    9. Generate reports (balance chart, quarterly rollup, hedging output)
"""
