"""
Financial Calculation Utilities
Shared helpers used across DCF, multiples, and due diligence modules.
"""

import numpy as np
import pandas as pd
from typing import List, Optional, Tuple


def xnpv(rate: float, cash_flows: List[float], dates: List) -> float:
    """
    Net Present Value with irregular cash flow dates (XNPV).
    dates: list of datetime.date objects.
    """
    import datetime
    t0 = dates[0]
    return sum(
        cf / (1 + rate) ** ((d - t0).days / 365.0)
        for cf, d in zip(cash_flows, dates)
    )


def xirr(cash_flows: List[float], dates: List, guess: float = 0.1) -> float:
    """
    Internal Rate of Return for irregular cash flow dates (XIRR).
    Uses Newton-Raphson method.
    """
    from scipy.optimize import brentq
    try:
        return brentq(lambda r: xnpv(r, cash_flows, dates), -0.999, 100.0, xtol=1e-6)
    except ValueError:
        return np.nan


def gordon_growth_value(fcf: float, wacc: float, g: float) -> float:
    """
    Gordon Growth Model terminal value.
    TV = FCF × (1+g) / (WACC - g)
    """
    if wacc <= g:
        raise ValueError(f"WACC ({wacc:.2%}) must exceed g ({g:.2%})")
    return fcf * (1 + g) / (wacc - g)


def capm(risk_free: float, beta: float, erp: float) -> float:
    """CAPM cost of equity: Ke = Rf + β × ERP"""
    return risk_free + beta * erp


def unlevered_beta(levered_beta: float, tax_rate: float, debt_equity_ratio: float) -> float:
    """
    Hamada equation: βu = βl / [1 + (1-T) × (D/E)]
    """
    return levered_beta / (1 + (1 - tax_rate) * debt_equity_ratio)


def levered_beta(unlevered: float, tax_rate: float, debt_equity_ratio: float) -> float:
    """Re-lever beta for target capital structure."""
    return unlevered * (1 + (1 - tax_rate) * debt_equity_ratio)


def enterprise_to_equity(ev: float, net_debt: float, minorities: float = 0.0,
                          preferred_equity: float = 0.0) -> float:
    """
    Bridge from Enterprise Value to Equity Value.
    Equity Value = EV - Net Debt - Minorities - Preferred
    """
    return ev - net_debt - minorities - preferred_equity


def equity_value_per_share(equity_value: float, shares_outstanding: float) -> float:
    """Equity value per diluted share."""
    return equity_value / shares_outstanding if shares_outstanding else np.nan


def implied_growth_rate(ev: float, last_fcf: float, wacc: float) -> float:
    """Solve for implied terminal growth rate given market EV."""
    # EV × (WACC - g) = FCF × (1+g)  → solve for g
    # g = (EV × WACC - FCF) / (EV + FCF)
    return (ev * wacc - last_fcf) / (ev + last_fcf)


def debt_capacity(ebitda: float, leverage_multiple: float, cash: float = 0.0) -> dict:
    """
    Estimates debt capacity and implied capital structure.
    """
    gross_debt = ebitda * leverage_multiple
    net_debt = gross_debt - cash
    return {
        "gross_debt": round(gross_debt, 1),
        "net_debt": round(net_debt, 1),
        "implied_leverage_ratio": round(leverage_multiple, 1),
    }


def sensitivity_2d(
    row_values: List[float],
    col_values: List[float],
    func,
    row_label: str = "Row",
    col_label: str = "Col",
) -> pd.DataFrame:
    """
    Generic 2D sensitivity table.
    func(row_val, col_val) → result.
    """
    data = {}
    for col_val in col_values:
        col_key = f"{col_label}={col_val:.3g}"
        data[col_key] = [func(row_val, col_val) for row_val in row_values]

    df = pd.DataFrame(data, index=[f"{row_label}={v:.3g}" for v in row_values])
    return df


def format_currency(value: float, unit: str = "$M", decimals: int = 0) -> str:
    """Format a number as currency string."""
    if pd.isna(value):
        return "N/A"
    fmt = f"{{:,.{decimals}f}}"
    return f"{unit} {fmt.format(value)}"


def cagr(start: float, end: float, periods: int) -> float:
    """Compound Annual Growth Rate."""
    if start <= 0 or periods <= 0:
        return np.nan
    return (end / start) ** (1 / periods) - 1
