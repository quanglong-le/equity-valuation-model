"""
Market Multiples — Comparable Company Analysis (CCA)
Computes trading multiples and derives implied valuation ranges.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional
import pandas as pd
import numpy as np


@dataclass
class CompanyData:
    """Holds financial data for one comparable company."""
    name: str
    ticker: str
    market_cap: float        # $M
    net_debt: float          # $M (positive = more debt than cash)
    revenue_ltm: float       # $M
    ebitda_ltm: float        # $M
    ebit_ltm: float          # $M
    net_income_ltm: float    # $M
    book_value_equity: float # $M
    shares_outstanding: float  # M

    @property
    def enterprise_value(self) -> float:
        return self.market_cap + self.net_debt

    @property
    def ev_revenue(self) -> float:
        return self.enterprise_value / self.revenue_ltm if self.revenue_ltm else np.nan

    @property
    def ev_ebitda(self) -> float:
        return self.enterprise_value / self.ebitda_ltm if self.ebitda_ltm else np.nan

    @property
    def ev_ebit(self) -> float:
        return self.enterprise_value / self.ebit_ltm if self.ebit_ltm else np.nan

    @property
    def pe_ratio(self) -> float:
        eps = self.net_income_ltm / self.shares_outstanding
        price_per_share = self.market_cap / self.shares_outstanding
        return price_per_share / eps if eps > 0 else np.nan

    @property
    def price_to_book(self) -> float:
        return self.market_cap / self.book_value_equity if self.book_value_equity else np.nan


class MultiplesAnalyzer:
    """
    Comparable Company Analysis engine.
    Derives implied valuation range for a target company.
    """

    def __init__(self, comps: List[CompanyData]):
        self.comps = comps

    def multiples_table(self) -> pd.DataFrame:
        """Returns DataFrame of all trading multiples for each comp."""
        rows = []
        for c in self.comps:
            rows.append({
                "Company": c.name,
                "Ticker": c.ticker,
                "Market Cap ($M)": round(c.market_cap, 0),
                "EV ($M)": round(c.enterprise_value, 0),
                "EV/Revenue": round(c.ev_revenue, 1),
                "EV/EBITDA": round(c.ev_ebitda, 1),
                "EV/EBIT": round(c.ev_ebit, 1),
                "P/E": round(c.pe_ratio, 1),
                "P/B": round(c.price_to_book, 1),
            })
        df = pd.DataFrame(rows)

        # Summary stats
        numeric_cols = ["EV/Revenue", "EV/EBITDA", "EV/EBIT", "P/E", "P/B"]
        stats = {}
        for col in numeric_cols:
            vals = df[col].replace([np.inf, -np.inf], np.nan).dropna()
            stats[col] = vals
        
        summary = pd.DataFrame({
            "Company": ["— Mean —", "— Median —", "— 25th Pct —", "— 75th Pct —"],
            "Ticker": ["", "", "", ""],
            "Market Cap ($M)": [np.nan]*4,
            "EV ($M)": [np.nan]*4,
        })
        for col in numeric_cols:
            vals = df[col].replace([np.inf, -np.inf], np.nan).dropna()
            summary[col] = [
                round(vals.mean(), 1),
                round(vals.median(), 1),
                round(vals.quantile(0.25), 1),
                round(vals.quantile(0.75), 1),
            ]

        return pd.concat([df, summary], ignore_index=True)

    def implied_value(
        self,
        target_revenue: float,
        target_ebitda: float,
        target_ebit: float,
        target_net_income: float,
        target_book_equity: float,
        net_debt_target: float = 0.0,
        method: str = "median",
    ) -> Dict[str, float]:
        """
        Derives implied EV and equity value for target using median (or mean) multiples.

        Returns dict with implied EV and equity value per multiple.
        """
        df = self.multiples_table()
        numeric_rows = df[~df["Ticker"].isin(["", "—"])].copy()

        agg = getattr(numeric_rows[["EV/Revenue","EV/EBITDA","EV/EBIT","P/E","P/B"]].replace([np.inf,-np.inf], np.nan), method)()

        results = {}

        if not np.isnan(agg["EV/Revenue"]):
            ev = target_revenue * agg["EV/Revenue"]
            results["EV/Revenue"] = {
                "multiple": round(agg["EV/Revenue"], 2),
                "implied_ev": round(ev, 0),
                "implied_equity": round(ev - net_debt_target, 0),
            }

        if not np.isnan(agg["EV/EBITDA"]):
            ev = target_ebitda * agg["EV/EBITDA"]
            results["EV/EBITDA"] = {
                "multiple": round(agg["EV/EBITDA"], 2),
                "implied_ev": round(ev, 0),
                "implied_equity": round(ev - net_debt_target, 0),
            }

        if not np.isnan(agg["EV/EBIT"]):
            ev = target_ebit * agg["EV/EBIT"]
            results["EV/EBIT"] = {
                "multiple": round(agg["EV/EBIT"], 2),
                "implied_ev": round(ev, 0),
                "implied_equity": round(ev - net_debt_target, 0),
            }

        return results

    def football_field_data(
        self,
        target_financials: dict,
        net_debt: float,
        dcf_low: float,
        dcf_high: float,
    ) -> pd.DataFrame:
        """
        Generates football field chart data showing valuation range by method.
        Returns DataFrame with columns: Method, Low, High (equity values $M).
        """
        df = self.multiples_table()
        numeric_rows = df[~df["Ticker"].isin(["", "—"])].copy()

        rows = []

        # DCF range
        rows.append({"Method": "DCF (WACC Sensitivity)", "Low": dcf_low, "High": dcf_high})

        # EV/EBITDA range (25th–75th pct)
        ebitda_multiples = numeric_rows["EV/EBITDA"].replace([np.inf,-np.inf], np.nan).dropna()
        if len(ebitda_multiples):
            low_ev = target_financials["ebitda"] * ebitda_multiples.quantile(0.25)
            high_ev = target_financials["ebitda"] * ebitda_multiples.quantile(0.75)
            rows.append({"Method": "EV/EBITDA Comps", "Low": low_ev - net_debt, "High": high_ev - net_debt})

        # EV/Revenue range
        rev_multiples = numeric_rows["EV/Revenue"].replace([np.inf,-np.inf], np.nan).dropna()
        if len(rev_multiples):
            low_ev = target_financials["revenue"] * rev_multiples.quantile(0.25)
            high_ev = target_financials["revenue"] * rev_multiples.quantile(0.75)
            rows.append({"Method": "EV/Revenue Comps", "Low": low_ev - net_debt, "High": high_ev - net_debt})

        return pd.DataFrame(rows)
