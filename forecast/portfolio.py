"""
Portfolio allocation and sub-portfolio cashflow distribution.

Computes par-weighted allocation of deal-level Intex cashflows across
user-defined sub-portfolios, which can group any combination of
PAM Portfolios, Entity Names, or Client Level 3 values.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------------
# Sub-portfolio definition
# ---------------------------------------------------------------------------

@dataclass
class SubPortfolio:
    """A named grouping of holdings positions defined by filter criteria.

    Filters are combined with AND logic within a SubPortfolio.
    Multiple values within a single filter field use OR logic.

    Examples:
        # Single PAM portfolio
        SubPortfolio("FHLB", pam_portfolios=[13091])

        # Multiple PAMs combined
        SubPortfolio("Fixed Annuity", pam_portfolios=[13015, 13035])

        # By entity name pattern
        SubPortfolio("All NLIC", client_level_3=["NLIC"])

        # Everything (no filters = include all)
        SubPortfolio("Total")

        # Mix of criteria (AND logic: must match ALL specified filters)
        SubPortfolio("NLIC FHLB Only", client_level_3=["NLIC"], pam_portfolios=[13091])
    """

    name: str
    pam_portfolios: list[int] = field(default_factory=list)
    entity_names: list[str] = field(default_factory=list)
    client_level_3: list[str] = field(default_factory=list)

    def matches(self, row: pd.Series) -> bool:
        """Check if a holdings row matches this sub-portfolio's filters."""
        if self.pam_portfolios:
            if row.get("PAM Portfolio") not in self.pam_portfolios:
                return False
        if self.entity_names:
            if row.get("Entity Name") not in self.entity_names:
                return False
        if self.client_level_3:
            if row.get("Client Level 3") not in self.client_level_3:
                return False
        return True

    def filter_holdings(self, holdings_df: pd.DataFrame) -> pd.DataFrame:
        """Return the subset of holdings that match this sub-portfolio."""
        if not self.pam_portfolios and not self.entity_names and not self.client_level_3:
            return holdings_df.copy()

        mask = pd.Series(True, index=holdings_df.index)
        if self.pam_portfolios:
            mask &= holdings_df["PAM Portfolio"].isin(self.pam_portfolios)
        if self.entity_names:
            mask &= holdings_df["Entity Name"].isin(self.entity_names)
        if self.client_level_3:
            mask &= holdings_df["Client Level 3"].isin(self.client_level_3)

        return holdings_df[mask].copy()


# ---------------------------------------------------------------------------
# Weight computation
# ---------------------------------------------------------------------------

def compute_weights(
    holdings_df: pd.DataFrame,
    sub_portfolios: list[SubPortfolio],
    weight_column: str = "Current Par",
) -> pd.DataFrame:
    """Compute ownership weights by CUSIP for each sub-portfolio.

    For each (sub_portfolio, CUSIP) pair, the weight is:
        sum of sub_portfolio's positions in that CUSIP / total across ALL positions

    Args:
        holdings_df: Full holdings DataFrame with CUSIP, PAM Portfolio,
                     Entity Name, Client Level 3, Current Par columns.
        sub_portfolios: List of SubPortfolio definitions.
        weight_column: Column to use for weighting (default: "Current Par").

    Returns:
        DataFrame with columns: SubPortfolio, CUSIP, Weight, SubPortfolio Par, Total Par
    """
    # Compute total par per CUSIP across ALL holdings.
    total_par = holdings_df.groupby("CUSIP")[weight_column].sum().rename("Total Par")

    rows = []
    for sp in sub_portfolios:
        filtered = sp.filter_holdings(holdings_df)
        if filtered.empty:
            continue

        sp_par = filtered.groupby("CUSIP")[weight_column].sum().rename("SubPortfolio Par")

        # Join with total par and compute weight.
        merged = pd.DataFrame({"SubPortfolio Par": sp_par, "Total Par": total_par}).dropna()
        merged["Weight"] = merged["SubPortfolio Par"] / merged["Total Par"]
        merged["SubPortfolio"] = sp.name

        for cusip, r in merged.iterrows():
            rows.append({
                "SubPortfolio": sp.name,
                "CUSIP": cusip,
                "Weight": r["Weight"],
                "SubPortfolio Par": r["SubPortfolio Par"],
                "Total Par": r["Total Par"],
            })

    result = pd.DataFrame(rows)
    if not result.empty:
        print(f"  Weights computed: {len(result)} (sub-portfolio, CUSIP) pairs "
              f"across {len(sub_portfolios)} sub-portfolios.")
    return result


# ---------------------------------------------------------------------------
# Cashflow allocation
# ---------------------------------------------------------------------------

def allocate_cashflows(
    cashflow_summary_path: str | Path,
    weights_df: pd.DataFrame,
    output_path: str | Path,
    cusip_row: int = 1,
) -> None:
    """Allocate deal-level cashflows to sub-portfolios using ownership weights.

    Reads a cashflow summary workbook (one sheet per scenario, with per-CUSIP
    Interest/Principal/Balance columns) and produces a new workbook with
    per-sub-portfolio sheets.

    Args:
        cashflow_summary_path: Path to the cashflow summary Excel file
            (output of Step 7 / generate_cashflow_summary).
        weights_df: Weights DataFrame from compute_weights().
        output_path: Path to write the allocated cashflows workbook.
        cusip_row: Row index (0-based) in the summary sheet containing CUSIPs
            (default 1, i.e. second row of the header area).
    """
    cashflow_summary_path = Path(cashflow_summary_path)
    output_path = Path(output_path)

    print(f"Allocating cashflows from '{cashflow_summary_path.name}'...")

    # Read all scenario sheets.
    xls = pd.ExcelFile(cashflow_summary_path, engine="openpyxl")
    sub_portfolio_names = weights_df["SubPortfolio"].unique()

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for sp_name in sub_portfolio_names:
            sp_weights = weights_df[weights_df["SubPortfolio"] == sp_name].set_index("CUSIP")

            for sheet_name in xls.sheet_names:
                if sheet_name == "Balance Summary":
                    continue

                df = pd.read_excel(xls, sheet_name=sheet_name)

                # Build allocated cashflows for this sub-portfolio + scenario.
                result = {"Date": df["Date"]}

                sp_interest = pd.Series(0.0, index=df.index)
                sp_principal = pd.Series(0.0, index=df.index)
                sp_balance = pd.Series(0.0, index=df.index)

                # Find CUSIP columns (pattern: "CUSIP Interest", "CUSIP Principal", "CUSIP Balance")
                cusip_cols = set()
                for col in df.columns:
                    for suffix in (" Interest", " Principal", " Balance"):
                        if col.endswith(suffix):
                            cusip_cols.add(col.replace(suffix, ""))

                for cusip in cusip_cols:
                    weight = sp_weights.loc[cusip, "Weight"] if cusip in sp_weights.index else 0.0
                    if weight == 0:
                        continue

                    int_col = f"{cusip} Interest"
                    prin_col = f"{cusip} Principal"
                    bal_col = f"{cusip} Balance"

                    if int_col in df.columns:
                        weighted_int = df[int_col] * weight
                        result[int_col] = weighted_int
                        sp_interest += weighted_int

                    if prin_col in df.columns:
                        weighted_prin = df[prin_col] * weight
                        result[prin_col] = weighted_prin
                        sp_principal += weighted_prin

                    if bal_col in df.columns:
                        weighted_bal = df[bal_col] * weight
                        result[bal_col] = weighted_bal
                        sp_balance += weighted_bal

                # Insert portfolio totals at the front.
                ordered = {"Date": result["Date"]}
                ordered[f"{sp_name} Interest"] = sp_interest
                ordered[f"{sp_name} Principal"] = sp_principal
                ordered[f"{sp_name} Balance"] = sp_balance
                for k, v in result.items():
                    if k != "Date":
                        ordered[k] = v

                out_sheet = f"{sp_name} - {sheet_name}"[:31]
                pd.DataFrame(ordered).to_excel(writer, sheet_name=out_sheet, index=False)

        # Write a Balance Summary across sub-portfolios.
        _write_balance_summary(xls, weights_df, sub_portfolio_names, writer)

    print(f"  Allocated cashflows written to '{output_path.name}'.")


def allocate_from_holdings_and_cashflows(
    holdings_df: pd.DataFrame,
    cashflow_summary_path: str | Path,
    sub_portfolios: list[SubPortfolio],
    output_path: str | Path,
) -> pd.DataFrame:
    """Convenience function: compute weights and allocate in one call.

    Returns the weights DataFrame for inspection.
    """
    weights_df = compute_weights(holdings_df, sub_portfolios)
    allocate_cashflows(cashflow_summary_path, weights_df, output_path)
    return weights_df


# ---------------------------------------------------------------------------
# Sub-portfolio definition helpers
# ---------------------------------------------------------------------------

def sub_portfolios_from_config(config: list[dict]) -> list[SubPortfolio]:
    """Create SubPortfolio objects from a list of config dicts.

    Example config:
        [
            {"name": "FHLB", "pam_portfolios": [13091]},
            {"name": "Fixed Annuity", "pam_portfolios": [13015, 13035]},
            {"name": "All Gen Acct", "client_level_3": ["NLIC"], "exclude_pam": [13091]},
            {"name": "Total"},
        ]
    """
    result = []
    for c in config:
        result.append(SubPortfolio(
            name=c["name"],
            pam_portfolios=c.get("pam_portfolios", []),
            entity_names=c.get("entity_names", []),
            client_level_3=c.get("client_level_3", []),
        ))
    return result


def one_per_entity(holdings_df: pd.DataFrame) -> list[SubPortfolio]:
    """Auto-generate one SubPortfolio per unique Entity Name in the holdings."""
    entities = holdings_df["Entity Name"].dropna().unique()
    return [SubPortfolio(name=e, entity_names=[e]) for e in sorted(entities)]


def one_per_pam(holdings_df: pd.DataFrame) -> list[SubPortfolio]:
    """Auto-generate one SubPortfolio per unique PAM Portfolio in the holdings."""
    pams = holdings_df["PAM Portfolio"].dropna().unique()
    result = []
    for pam in sorted(pams):
        # Use entity name as the label if there's a 1:1 mapping.
        entities = holdings_df[holdings_df["PAM Portfolio"] == pam]["Entity Name"].unique()
        label = entities[0] if len(entities) == 1 else f"PAM {int(pam)}"
        result.append(SubPortfolio(name=label, pam_portfolios=[int(pam)]))
    return result


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _write_balance_summary(
    xls: pd.ExcelFile,
    weights_df: pd.DataFrame,
    sub_portfolio_names,
    writer: pd.ExcelWriter,
) -> None:
    """Write a cross-sub-portfolio balance summary sheet."""
    all_data: dict[str, dict[str, pd.Series]] = {}

    for sheet_name in xls.sheet_names:
        if sheet_name == "Balance Summary":
            continue
        df = pd.read_excel(xls, sheet_name=sheet_name)

        if "Date" not in df.columns or "Portfolio Balance" not in df.columns:
            continue

        for sp_name in sub_portfolio_names:
            sp_weights = weights_df[weights_df["SubPortfolio"] == sp_name].set_index("CUSIP")

            # Find CUSIP balance columns and compute weighted portfolio balance.
            sp_balance = pd.Series(0.0, index=df.index)
            for col in df.columns:
                if col.endswith(" Balance") and col != "Portfolio Balance":
                    cusip = col.replace(" Balance", "")
                    weight = sp_weights.loc[cusip, "Weight"] if cusip in sp_weights.index else 0.0
                    if weight > 0:
                        sp_balance += df[col] * weight

            key = f"{sp_name} - {sheet_name}"
            all_data[key] = {"Date": df["Date"], "Balance": sp_balance}

    if not all_data:
        return

    # Pivot into a summary table.
    first_key = next(iter(all_data))
    summary = {"Date": all_data[first_key]["Date"]}
    for key, data in all_data.items():
        summary[key] = data["Balance"]

    pd.DataFrame(summary).to_excel(writer, sheet_name="Balance Summary", index=False)
