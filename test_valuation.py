"""
Unit Tests — Equity Valuation Model
Run with: pytest tests/ -v
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest
import numpy as np
import pandas as pd

from src.dcf.wacc import WACCCalculator
from src.dcf.free_cash_flow import FCFProjector
from src.dcf.terminal_value import TerminalValueCalculator
from src.multiples.market_multiples import MultiplesAnalyzer, CompanyData
from src.utils.financial_calcs import (
    capm, gordon_growth_value, unlevered_beta,
    levered_beta, enterprise_to_equity, cagr, sensitivity_2d
)


# ── WACC Tests ────────────────────────────────────────────────────────────────

class TestWACC:
    def test_basic_wacc(self):
        calc = WACCCalculator(
            risk_free_rate=0.04, equity_risk_premium=0.055,
            beta=1.2, cost_of_debt=0.045, tax_rate=0.21, debt_weight=0.3
        )
        wacc = calc.compute()
        assert 0.07 < wacc < 0.12, f"WACC {wacc:.2%} outside expected range"

    def test_weights_sum_to_one(self):
        calc = WACCCalculator(0.04, 0.055, 1.0, 0.05, 0.25, 0.35)
        assert abs(calc.equity_weight + calc.debt_weight - 1.0) < 1e-6

    def test_invalid_tax_rate(self):
        with pytest.raises(ValueError):
            WACCCalculator(0.04, 0.055, 1.0, 0.05, 1.5, 0.3).compute()

    def test_cost_of_equity_capm(self):
        calc = WACCCalculator(0.04, 0.055, 1.0, 0.05, 0.21, 0.3)
        ke = calc.cost_of_equity
        expected = 0.04 + 1.0 * 0.055
        assert abs(ke - expected) < 1e-8

    def test_after_tax_debt(self):
        calc = WACCCalculator(0.04, 0.055, 1.0, 0.05, 0.20, 0.3)
        assert abs(calc.after_tax_cost_of_debt - 0.05 * 0.80) < 1e-8

    def test_sensitivity_table(self):
        calc = WACCCalculator(0.04, 0.055, 1.0, 0.05, 0.21, 0.3)
        table = calc.sensitivity_table(steps=3)
        assert len(table) == 9


# ── FCF Projection Tests ──────────────────────────────────────────────────────

class TestFCFProjector:
    def _make_projector(self):
        return FCFProjector(
            base_revenue=10_000,
            revenue_growth=[0.10, 0.09, 0.08, 0.07, 0.06],
            ebitda_margin=0.25,
            da_pct_revenue=0.05,
            capex_pct=0.05,
            nwc_change_pct=0.02,
            tax_rate=0.21,
        )

    def test_project_returns_dataframe(self):
        proj = self._make_projector()
        df = proj.project()
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 5

    def test_revenue_grows(self):
        proj = self._make_projector()
        df = proj.project()
        revenues = df["Revenue ($M)"].tolist()
        assert all(revenues[i] < revenues[i+1] for i in range(len(revenues)-1))

    def test_fcf_positive(self):
        proj = self._make_projector()
        fcfs = proj.get_fcf_list()
        assert all(f > 0 for f in fcfs), "All FCFs should be positive"

    def test_pv_less_than_undiscounted(self):
        proj = self._make_projector()
        fcfs = proj.get_fcf_list()
        pv = proj.discount_fcfs(fcfs, wacc=0.09)
        assert pv < sum(fcfs), "PV must be less than sum of undiscounted FCFs"


# ── Terminal Value Tests ──────────────────────────────────────────────────────

class TestTerminalValue:
    def test_gordon_growth(self):
        calc = TerminalValueCalculator(method="gordon_growth", terminal_growth=0.025)
        tv = calc.compute(last_fcf=1000, wacc=0.09)
        expected = 1000 * 1.025 / (0.09 - 0.025)
        assert abs(tv - expected) < 0.01

    def test_wacc_must_exceed_growth(self):
        calc = TerminalValueCalculator(method="gordon_growth", terminal_growth=0.10)
        with pytest.raises(ValueError):
            calc.compute(last_fcf=1000, wacc=0.09)

    def test_exit_multiple(self):
        calc = TerminalValueCalculator(method="exit_multiple", exit_multiple=12.0)
        tv = calc.compute(last_fcf=500, wacc=0.09, last_ebitda=800)
        assert tv == 800 * 12.0

    def test_pv_discounting(self):
        calc = TerminalValueCalculator()
        pv = calc.pv(terminal_value=10_000, wacc=0.09, years=5)
        expected = 10_000 / (1.09 ** 5)
        assert abs(pv - expected) < 0.01

    def test_tv_pct(self):
        calc = TerminalValueCalculator()
        pct = calc.tv_as_pct_of_ev(7000, 3000)
        assert abs(pct - 0.70) < 1e-6


# ── Market Multiples Tests ────────────────────────────────────────────────────

class TestMultiplesAnalyzer:
    def _make_analyzer(self):
        comps = [
            CompanyData("A Corp", "A", 20_000, 1_000, 6_000, 1_500, 1_200, 1_000, 8_000, 200),
            CompanyData("B Corp", "B", 15_000, 800, 4_000, 900, 720, 600, 5_000, 150),
        ]
        return MultiplesAnalyzer(comps)

    def test_multiples_table(self):
        analyzer = self._make_analyzer()
        df = analyzer.multiples_table()
        assert "EV/EBITDA" in df.columns
        assert len(df) > 2  # includes summary rows

    def test_ev_revenue_positive(self):
        comp = CompanyData("X", "X", 10_000, 500, 3_000, 800, 600, 500, 4_000, 100)
        assert comp.ev_revenue > 0

    def test_implied_value(self):
        analyzer = self._make_analyzer()
        implied = analyzer.implied_value(5_000, 1_200, 960, 800, 6_000, 900)
        assert "EV/EBITDA" in implied or "EV/Revenue" in implied


# ── Utility Function Tests ────────────────────────────────────────────────────

class TestFinancialUtils:
    def test_capm(self):
        ke = capm(risk_free=0.04, beta=1.0, erp=0.055)
        assert abs(ke - 0.095) < 1e-8

    def test_gordon_growth(self):
        tv = gordon_growth_value(fcf=1000, wacc=0.09, g=0.025)
        assert tv > 0

    def test_unlevered_beta(self):
        bu = unlevered_beta(levered_beta=1.2, tax_rate=0.21, debt_equity_ratio=0.5)
        assert bu < 1.2

    def test_relever_beta_roundtrip(self):
        bl_original = 1.2
        bu = unlevered_beta(bl_original, 0.21, 0.5)
        bl_new = levered_beta(bu, 0.21, 0.5)
        assert abs(bl_original - bl_new) < 1e-8

    def test_enterprise_to_equity(self):
        eq = enterprise_to_equity(ev=50_000, net_debt=5_000, minorities=500)
        assert eq == 44_500

    def test_cagr(self):
        rate = cagr(start=100, end=161.05, periods=5)
        assert abs(rate - 0.10) < 0.001

    def test_sensitivity_2d(self):
        df = sensitivity_2d([0.08, 0.09], [0.02, 0.03], lambda r, c: r + c)
        assert df.shape == (2, 2)
