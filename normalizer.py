"""
Financial Data Normalizer
Cleans, standardizes, and computes derived KPIs from raw financial statements.
"""

import pandas as pd
import numpy as np
from typing import Optional
from src.due_diligence.extractor import FinancialStatements


class FinancialNormalizer:
    """
    Normalizes raw financial statements into analysis-ready DataFrames.
    Computes margins, growth rates, and key financial ratios.
    """

    def __init__(self, statements: FinancialStatements):
        self.statements = statements
        self.ticker = statements.ticker

    # ── helpers ──────────────────────────────────────────────────────────────

    @staticmethod
    def _safe_div(a: pd.Series, b: pd.Series) -> pd.Series:
        return a.divide(b).replace([np.inf, -np.inf], np.nan)

    @staticmethod
    def _yoy_growth(series: pd.Series) -> pd.Series:
        return series.pct_change().replace([np.inf, -np.inf], np.nan)

    # ── income statement ─────────────────────────────────────────────────────

    def income_kpis(self) -> pd.DataFrame:
        """Returns normalized income statement with margins and growth."""
        is_ = self.statements.income_statement.copy()

        # Standardize column names (Yahoo Finance uses different names)
        col_map = {
            "Total Revenue": "Revenue",
            "Gross Profit": "Gross Profit",
            "EBITDA": "EBITDA",
            "EBIT": "EBIT",
            "Net Income": "Net Income",
        }
        is_.rename(columns={k: v for k, v in col_map.items() if k in is_.columns}, inplace=True)

        out = pd.DataFrame(index=is_.index)

        for col in ["Revenue", "Gross Profit", "EBITDA", "EBIT", "Net Income"]:
            if col in is_.columns:
                out[col] = is_[col]

        # Margins
        if "Revenue" in out.columns and out["Revenue"].notna().any():
            rev = out["Revenue"]
            for col in ["Gross Profit", "EBITDA", "EBIT", "Net Income"]:
                if col in out.columns:
                    out[f"{col} Margin"] = self._safe_div(out[col], rev).map(
                        lambda x: f"{x:.1%}" if pd.notna(x) else "n/a"
                    )

            # Revenue growth
            out["Revenue Growth"] = self._yoy_growth(rev).map(
                lambda x: f"{x:.1%}" if pd.notna(x) else "n/a"
            )

        return out.sort_index()

    # ── balance sheet ─────────────────────────────────────────────────────────

    def balance_sheet_kpis(self) -> pd.DataFrame:
        """Returns normalized balance sheet with leverage ratios."""
        bs = self.statements.balance_sheet.copy()

        out = pd.DataFrame(index=bs.index)

        cash_col = next((c for c in bs.columns if "cash" in c.lower()), None)
        debt_col = next((c for c in bs.columns if "total debt" in c.lower()), None)
        equity_col = next((c for c in bs.columns if "stockholder" in c.lower() or "equity" in c.lower()), None)

        if cash_col:
            out["Cash ($M)"] = bs[cash_col]
        if debt_col:
            out["Total Debt ($M)"] = bs[debt_col]
        if equity_col:
            out["Equity ($M)"] = bs[equity_col]

        # Net Debt
        if "Cash ($M)" in out.columns and "Total Debt ($M)" in out.columns:
            out["Net Debt ($M)"] = out["Total Debt ($M)"] - out["Cash ($M)"]

        # Debt/Equity
        if "Total Debt ($M)" in out.columns and "Equity ($M)" in out.columns:
            out["D/E Ratio"] = self._safe_div(out["Total Debt ($M)"], out["Equity ($M)"]).round(2)

        return out.sort_index()

    # ── cash flow ─────────────────────────────────────────────────────────────

    def cashflow_kpis(self) -> pd.DataFrame:
        """Returns normalized cash flow statement with FCF metrics."""
        cf = self.statements.cash_flow.copy()

        out = pd.DataFrame(index=cf.index)

        ocf_col = next((c for c in cf.columns if "operating" in c.lower()), None)
        capex_col = next((c for c in cf.columns if "capital expenditure" in c.lower() or "capex" in c.lower()), None)
        fcf_col = next((c for c in cf.columns if "free cash" in c.lower()), None)

        if ocf_col:
            out["Operating CF ($M)"] = cf[ocf_col]
        if capex_col:
            out["CapEx ($M)"] = cf[capex_col].abs()
        if fcf_col:
            out["Free Cash Flow ($M)"] = cf[fcf_col]
        elif ocf_col and capex_col:
            out["Free Cash Flow ($M)"] = cf[ocf_col] + cf[capex_col]  # capex is negative

        # FCF Conversion (FCF / Net Income)
        ni = self.statements.income_statement.get("Net Income")
        if ni is not None and "Free Cash Flow ($M)" in out.columns:
            out["FCF Conversion"] = self._safe_div(out["Free Cash Flow ($M)"], ni).map(
                lambda x: f"{x:.1%}" if pd.notna(x) else "n/a"
            )

        return out.sort_index()

    # ── summary ───────────────────────────────────────────────────────────────

    def full_summary(self) -> dict:
        """Returns dict of all normalized tables."""
        return {
            "income_kpis": self.income_kpis(),
            "balance_sheet_kpis": self.balance_sheet_kpis(),
            "cashflow_kpis": self.cashflow_kpis(),
        }

    def ltm_metrics(self) -> dict:
        """Returns most recent year (LTM proxy) key metrics as a flat dict."""
        is_ = self.statements.income_statement
        bs = self.statements.balance_sheet
        cf = self.statements.cash_flow

        def last(df: pd.DataFrame, col_fragment: str) -> Optional[float]:
            col = next((c for c in df.columns if col_fragment.lower() in c.lower()), None)
            if col and len(df) > 0:
                val = df[col].dropna()
                return float(val.iloc[-1]) if len(val) else None
            return None

        return {
            "revenue": last(is_, "revenue"),
            "ebitda": last(is_, "ebitda"),
            "ebit": last(is_, "ebit"),
            "net_income": last(is_, "net income"),
            "total_debt": last(bs, "total debt"),
            "cash": last(bs, "cash"),
            "free_cash_flow": last(cf, "free cash"),
        }
