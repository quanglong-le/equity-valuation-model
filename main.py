#!/usr/bin/env python3
"""
Equity Valuation Model — Main Entry Point
Usage:
    python main.py --ticker AAPL --report
    python main.py --demo
"""

import argparse
import sys
import os

# Add src to path
sys.path.insert(0, os.path.dirname(__file__))

from src.dcf.wacc import WACCCalculator
from src.dcf.free_cash_flow import FCFProjector
from src.dcf.terminal_value import TerminalValueCalculator
from src.multiples.market_multiples import MultiplesAnalyzer, CompanyData
from src.due_diligence.extractor import FinancialExtractor
from src.due_diligence.normalizer import FinancialNormalizer
from src.due_diligence.report_generator import ValuationReportGenerator
from src.utils.financial_calcs import sensitivity_2d, enterprise_to_equity

import pandas as pd
import numpy as np


# ── Demo Comparable Companies ─────────────────────────────────────────────────

DEMO_COMPS = [
    CompanyData("Alpha Tech", "ATEC", 42_000, 2_500, 12_600, 3_150, 2_520, 2_100, 14_200, 500),
    CompanyData("Beta Systems", "BSYS", 28_000, 1_200, 7_800, 1_872, 1_404, 1_100, 9_500, 320),
    CompanyData("Gamma Corp", "GAMC", 65_000, 4_100, 18_200, 4_914, 4_004, 3_500, 22_000, 750),
    CompanyData("Delta Inc", "DLTA", 19_500, 800, 5_400, 1_188, 918, 720, 6_800, 200),
    CompanyData("Epsilon Ltd", "EPSL", 35_000, 2_000, 10_500, 2_625, 2_100, 1_800, 13_000, 410),
]


# ── Core Valuation Logic ───────────────────────────────────────────────────────

def run_valuation(ticker: str, use_demo_data: bool = True) -> dict:
    """
    Full valuation pipeline:
    1. Extract & normalize financial statements
    2. Compute WACC
    3. Project FCFs and discount
    4. Compute terminal value
    5. Run market multiples analysis
    6. Return consolidated results dict
    """

    print(f"\n{'='*60}")
    print(f"  EQUITY VALUATION MODEL — {ticker.upper()}")
    print(f"{'='*60}\n")

    # ── 1. Financial Statements ────────────────────────────────────────────────
    extractor = FinancialExtractor(ticker)

    if use_demo_data:
        print("📊 Loading sample financial data (demo mode)...")
        statements = FinancialExtractor.sample_statements(ticker)
    else:
        print(f"📡 Fetching live financial data for {ticker}...")
        statements = extractor.fetch()

    normalizer = FinancialNormalizer(statements)
    ltm = normalizer.ltm_metrics()

    print(f"✅ LTM Revenue:   ${ltm.get('revenue', 0):,.0f}M")
    print(f"✅ LTM EBITDA:    ${ltm.get('ebitda', 0):,.0f}M")
    print(f"✅ LTM Net Income:${ltm.get('net_income', 0):,.0f}M")
    print(f"✅ Net Debt:      ${ltm.get('total_debt', 0) - ltm.get('cash', 0):,.0f}M\n")

    # ── 2. WACC ───────────────────────────────────────────────────────────────
    wacc_calc = WACCCalculator(
        risk_free_rate=0.040,
        equity_risk_premium=0.055,
        beta=1.15,
        cost_of_debt=0.045,
        tax_rate=0.21,
        debt_weight=0.25,
        size_premium=0.01,
    )
    wacc = wacc_calc.compute()
    print(f"📐 WACC Components:")
    for k, v in wacc_calc.summary().items():
        print(f"   {k:<30} {v}")
    print()

    # ── 3. FCF Projection ─────────────────────────────────────────────────────
    base_revenue = ltm.get("revenue") or 12_600
    fcf_proj = FCFProjector(
        base_revenue=base_revenue,
        revenue_growth=[0.12, 0.10, 0.09, 0.08, 0.07],
        ebitda_margin=0.27,
        da_pct_revenue=0.05,
        capex_pct=0.06,
        nwc_change_pct=0.015,
        tax_rate=0.21,
    )
    fcf_df = fcf_proj.project()
    fcf_list = fcf_proj.get_fcf_list()
    pv_fcfs = fcf_proj.pv_fcfs(wacc)

    print("📈 FCF Projections:")
    print(fcf_df[["Revenue ($M)", "EBITDA ($M)", "FCF ($M)"]].to_string())
    print()

    # ── 4. Terminal Value ─────────────────────────────────────────────────────
    tv_calc = TerminalValueCalculator(method="gordon_growth", terminal_growth=0.025)
    tv = tv_calc.compute(fcf_list[-1], wacc)
    pv_tv = tv_calc.pv(tv, wacc, years=5)
    tv_pct = tv_calc.tv_as_pct_of_ev(pv_tv, pv_fcfs)

    enterprise_value = pv_fcfs + pv_tv
    net_debt = (ltm.get("total_debt") or 4_200) - (ltm.get("cash") or 2_100)
    equity_value = enterprise_to_equity(enterprise_value, net_debt)

    print(f"💰 DCF Valuation:")
    print(f"   PV of FCFs:        ${pv_fcfs:,.0f}M")
    print(f"   Terminal Value:    ${tv:,.0f}M")
    print(f"   PV(Terminal Value):${pv_tv:,.0f}M  ({tv_pct:.1%} of EV)")
    print(f"   Enterprise Value:  ${enterprise_value:,.0f}M")
    print(f"   Net Debt:          ${net_debt:,.0f}M")
    print(f"   Equity Value:      ${equity_value:,.0f}M\n")

    # ── 4b. Sensitivity ───────────────────────────────────────────────────────
    tv_sensitivity = tv_calc.sensitivity_table(fcf_list[-1], wacc)

    wacc_range = np.linspace(wacc - 0.015, wacc + 0.015, 5)
    growth_range = np.linspace(0.015, 0.035, 5)

    def ev_from_params(w, g):
        tv_ = TerminalValueCalculator(method="gordon_growth", terminal_growth=g).compute(fcf_list[-1], w)
        pv_tv_ = TerminalValueCalculator(method="gordon_growth", terminal_growth=g).pv(tv_, w, 5)
        pv_f_ = fcf_proj.discount_fcfs(fcf_list, w)
        return round(pv_f_ + pv_tv_ - net_debt, 0)

    sensitivity_df = sensitivity_2d(
        wacc_range.tolist(),
        growth_range.tolist(),
        ev_from_params,
        row_label="WACC",
        col_label="g"
    )
    equity_values = [ev_from_params(w, 0.025) for w in wacc_range]
    ev_low, ev_high = min(equity_values), max(equity_values)

    print(f"📊 WACC × Terminal Growth Sensitivity (Equity Value $M):")
    print(sensitivity_df.to_string())
    print()

    # ── 5. Market Multiples ───────────────────────────────────────────────────
    analyzer = MultiplesAnalyzer(DEMO_COMPS)
    multiples_tbl = analyzer.multiples_table()
    implied = analyzer.implied_value(
        target_revenue=ltm.get("revenue") or 12_600,
        target_ebitda=ltm.get("ebitda") or 3_400,
        target_ebit=(ltm.get("ebitda") or 3_400) * 0.85,
        target_net_income=ltm.get("net_income") or 2_800,
        target_book_equity=14_200,
        net_debt_target=net_debt,
    )

    print("📊 Market Multiples — Implied Equity Values:")
    for method, data in implied.items():
        print(f"   {method:<15} {data['multiple']:.1f}x → ${data['implied_equity']:,.0f}M")
    print()

    # ── 6. Football Field ─────────────────────────────────────────────────────
    ff_data = analyzer.football_field_data(
        target_financials={"ebitda": ltm.get("ebitda") or 3_400, "revenue": ltm.get("revenue") or 12_600},
        net_debt=net_debt,
        dcf_low=ev_low,
        dcf_high=ev_high,
    )

    return {
        "ticker": ticker,
        "company_name": statements.info.get("shortName", ticker),
        "wacc": wacc,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "equity_value_low": ev_low,
        "equity_value_high": ev_high,
        "tv_pct": tv_pct,
        "intrinsic_equity_value": equity_value,
        "fcf_table": fcf_df,
        "income_df": normalizer.income_kpis(),
        "balance_df": normalizer.balance_sheet_kpis(),
        "cashflow_df": normalizer.cashflow_kpis(),
        "multiples_df": multiples_tbl,
        "football_field_df": ff_data,
        "tv_sensitivity": tv_sensitivity,
    }


def generate_report(results: dict, output_dir: str = "reports"):
    """Generate PDF valuation report from results dict."""
    path = os.path.join(output_dir, f"{results['ticker']}_valuation_report.pdf")
    gen = ValuationReportGenerator(path)
    gen.generate(
        ticker=results["ticker"],
        company_name=results["company_name"],
        income_df=results["income_df"],
        balance_df=results["balance_df"],
        cashflow_df=results["cashflow_df"],
        dcf_summary=results,
        multiples_df=results["multiples_df"],
        football_field_df=results["football_field_df"],
        sensitivity_data=results.get("tv_sensitivity"),
        analyst_notes="Projections based on LTM financials and management guidance.",
    )
    return path


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Equity Valuation Model")
    parser.add_argument("--ticker", type=str, default="DEMO", help="Stock ticker symbol")
    parser.add_argument("--report", action="store_true", help="Generate PDF report")
    parser.add_argument("--output", type=str, default="reports", help="Output directory")
    parser.add_argument("--live", action="store_true", help="Fetch live data (requires internet)")
    parser.add_argument("--demo", action="store_true", help="Run demo with sample data")
    args = parser.parse_args()

    use_demo = not args.live
    results = run_valuation(args.ticker, use_demo_data=use_demo)

    if args.report:
        pdf_path = generate_report(results, args.output)
        print(f"\n✅ Report saved to: {pdf_path}")

    print(f"\n{'='*60}")
    print(f"  VALUATION SUMMARY — {results['ticker'].upper()}")
    print(f"{'='*60}")
    print(f"  WACC:              {results['wacc']:.2%}")
    print(f"  Enterprise Value:  ${results['enterprise_value']:,.0f}M")
    print(f"  Equity Value:      ${results['equity_value']:,.0f}M")
    print(f"  Equity Range:      ${results['equity_value_low']:,.0f}M – ${results['equity_value_high']:,.0f}M")
    print(f"  TV as % of EV:     {results['tv_pct']:.1%}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
