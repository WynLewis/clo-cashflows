# CLO Cashflow Forecast

Python replacement for the `Forecast v16.3.xlsm` VBA macro workbook. Models how CLO portfolio cashflows and balances evolve under different AAA spread scenarios, where tighter spreads cause more deals to be called (refinanced) sooner.

## Quick Start

```bash
pip install -r requirements.txt
```

Open `CLO_Forecast_Workflow.ipynb` and run cells top to bottom. The only value you **must** set each run:

```python
CONFIG.scenario.current_aaa_margin_bps = 115  # where AAA CLO spreads are today
```

## Architecture

```
forecast/
  config.py           — All tunable parameters (spreads, assumptions, scenarios)
  models.py           — Data classes: Tranche, Scenario, Cashflow, CashflowsReport
  holdings.py         — Import holdings, enrich with Intex/BBG data, clean
  bloomberg.py        — Bloomberg data via blpapi (replaces Excel BDP formulas)
  preprice.py         — Pre-price deal lookup (currently unused)
  scenarios.py        — Scenario setup, call logic, Intex upload generation
  factors.py          — Factor call date extraction from COLLAT cashflows
  forward_curve.py    — Forward rate curve loading
  cashflows_summary.py— Parse Intex cashflow exports, summarize by scenario
  portfolio.py        — Sub-portfolio allocation (par-weighted)
  reporting.py        — Balance charts, quarterly rollup, hedging output
  excel_addins.py     — COM/XLL add-in toggling for xlwings
  file_picker.py      — Native file browser dialog

CLO_Forecast_Workflow.ipynb  — Main interactive notebook (run this)
run_forecast.py              — CLI alternative to the notebook
```

## Workflow

### Step 1: Import Holdings

Loads CLO positions. Three sources (set `HOLDINGS_SOURCE` in the notebook):

| Source | Method | When to use |
|--------|--------|-------------|
| `"clo_library"` (default) | `structured_products.clo.holdings()` | Normal runs — pulls live data |
| `"data_packet"` | Excel Data Packet file | When clo library unavailable |
| `"forecast_wb"` | Forecast workbook | When Intex/BBG data already populated |

### Step 2: Enrich with Intex/Bloomberg Data

Holdings need additional fields from two sources:

**Bloomberg** (automated via `blpapi`):
- `BBG_NC_END` — non-call end date (`MTG_DEAL_CALL_DT`)
- `BBG_REINVEST_END` — reinvestment end date (`REINVEST_END_DATE`)
- `RESET_IDX` — rate reset index, e.g. SOFR vs LIBOR
- `COLLAT_TYP` — collateral type, BSL vs middle market

**IntexLINK** (via Excel formulas):
- `Intex Name` — deal,tranche identifier (e.g. `CAPC37C4,BR`)
- `Intex Deal Name` — deal only (e.g. `CAPC37C4`)
- `AAA Margin` — AAA tranche coupon spread
- `Non-Call End` / `Reinvest End` — deal dates
- `Orig Deal Balance` — original collateral balance

The enrichment function:
1. Fetches Bloomberg data via `blpapi` (no Excel needed)
2. Writes an Excel file with IntexLINK formulas + Bloomberg values
3. You open it in Excel, let IntexLINK calculate, save
4. On re-run, the saved file is loaded automatically

Set `AUTO_CALC_INTEX = True` in the notebook to try fully automated xlwings calculation (experimental).

**Cleaning** (`clean_holdings()`):
- Cross-fills dates between BBG and Intex sources (bidirectional)
- Drops CUSIPs missing required fields (only if **both** sources are empty for dates)

### Step 3: Clear Forecasted Factors

Resets factor call dates for a fresh run.

### Step 4: Generate Preliminary Intex Upload

Creates `PreliminaryPortfolioUpload.xlsx` with COLLAT tranches (one per deal per scenario). Purpose: get deal-level collateral cashflows to determine factor-based call dates.

Each row includes call timing based on spread economics:
```
strike = current_aaa_margin + shock + refi_costs (10 bps)
       + mm_basis (50 bps if middle market)
       - libor_adj (26 bps if LIBOR-based)

if deal_aaa_margin > strike → deal called at non-call end date
```

### Step 4b: Run in IntexCalc (Manual)

Upload the file to IntexCalc, run cashflows, export to Excel.

### Step 5: Extract Factor Call Dates

Parses the COLLAT cashflows to find when each deal's balance drops below 35% of original — that's the factor-based call date. A file browser prompts you to select the IntexCalc export.

### Step 6: Generate Final Intex Upload

Creates `FinalPortfolioUpload.xlsx` with actual tranches. Call date for each = **earlier** of:
1. Spread-based refi call (from Step 4 logic)
2. Factor call (from Step 5)

This is the key mechanism: different spread scenarios produce different call dates, which produce different cashflow profiles.

### Step 6b: Run in IntexCalc (Manual)

Upload the final file, run cashflows, export to Excel.

### Step 7: Summarize Cashflows

Parses the final Intex export into per-scenario summaries with Interest, Principal, and Balance for each tranche.

### Step 8: Sub-Portfolio Allocation

Allocates deal-level cashflows to sub-portfolios using par-weighted ownership:

```
weight = entity's Current Par for CUSIP / total Current Par for CUSIP
entity's cashflow = deal cashflow × weight
```

FHLB is identified dynamically by matching Entity Names containing "FHLB". Sub-portfolios are customizable:

```python
sub_portfolios = [
    SubPortfolio("FHLB", pam_portfolios=[13091]),
    SubPortfolio("Fixed Annuity", pam_portfolios=[13015, 13035]),
    SubPortfolio("Total"),  # everything
]
```

### Step 9: Reports & Visualization

Generates from in-memory data (no Excel round-trip):
- **Balance Summary** — portfolio balance across all scenarios
- **Quarterly Forecasting** — per-CUSIP balances at quarter-ends (FHLB reporting)
- **Hedging** — first 2 payment dates/balances per CUSIP (derivatives desk)
- **Balance chart** — line plot of portfolio balance under each scenario (PNG)

## Configuration Reference

All values in `forecast/config.py`, accessible via `CONFIG`:

### `CONFIG.scenario` — Set Each Run

| Parameter | Default | Description |
|-----------|---------|-------------|
| `current_aaa_margin_bps` | 115 | Your view of current BSL AAA spreads (bps) |

### `CONFIG.defaults` — Rarely Changed

| Parameter | Default | Description |
|-----------|---------|-------------|
| `settle_date` | today | Settlement date |
| `prepay_speed` | 15 | CPR assumption |
| `scenarios` | 8 shocks | List of `ScenarioDefinition(name, shock_bps)` |
| `default_rate` | 5 | CDR (%) |
| `severity_pct` | 50 | Loss severity (%) |
| `recovery_lag_months` | 12 | Recovery lag (months) |
| `horizon_given_type` | DISC_MARGIN | Intex analytics type |

### `CONFIG.call` — Call Logic

| Parameter | Default | Description |
|-----------|---------|-------------|
| `refi_costs_bps` | 10 | Refi friction cost |
| `middle_market_bsl_basis_bps` | 50 | MM spread premium over BSL |
| `libor_sofr_basis_bps` | 26.161 | LIBOR-to-SOFR adjustment |
| `libor_reset_index` | US0003M | Index triggering LIBOR adjustment |
| `middle_market_collat_type` | CF-CLO-MML | Bloomberg collat type for MM CLOs |

### `CONFIG.factor_call`

| Parameter | Default | Description |
|-----------|---------|-------------|
| `factor_threshold` | 0.35 | Deal factor below which a call is triggered |

### `CONFIG.excel` — Add-in Configuration

| Parameter | Default | Description |
|-----------|---------|-------------|
| `intex_com_addin` | IntexLINK.Connect | COM add-in name |
| `intex_prefix` | (empty) | XLL prefix for formulas if `#NAME?` errors |

## Output Files

| File | Description |
|------|-------------|
| `{DATE}_HoldingsEnrichment.xlsx` | Holdings with Intex formulas + BBG data |
| `{DATE}_PreliminaryPortfolioUpload.xlsx` | COLLAT upload for IntexCalc |
| `{DATE}_FinalPortfolioUpload.xlsx` | Final tranche upload for IntexCalc |
| `{DATE}_CashflowSummary.xlsx` | Per-scenario cashflow summary |
| `{DATE}_AllocatedCashflows.xlsx` | Cashflows by sub-portfolio |
| `{DATE}_Report.xlsx` | Balance Summary + Quarterly + Hedging |
| `{DATE}_Report.png` | Balance scenario chart |

## Dependencies

- `pandas`, `numpy` — data manipulation
- `openpyxl` — reading/writing Excel files
- `xlsxwriter` — writing formatted Excel output
- `xlwings` — Excel automation (optional, for auto-calc)
- `blpapi` — Bloomberg data (requires Terminal + C++ SDK)
- `matplotlib`, `seaborn` — visualization
- `structured_products` — internal holdings library (optional)
