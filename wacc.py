"""
WACC Calculator — Weighted Average Cost of Capital
Uses CAPM for cost of equity, pre-tax yield for cost of debt.
"""

from dataclasses import dataclass, field
from typing import Optional
import numpy as np


@dataclass
class WACCCalculator:
    """
    Computes WACC using:
        WACC = Ke × (E/V) + Kd × (1 - T) × (D/V)

    Parameters
    ----------
    risk_free_rate       : 10-year government bond yield
    equity_risk_premium  : Damodaran ERP or implied ERP
    beta                 : Levered equity beta
    cost_of_debt         : Pre-tax cost of debt (yield to maturity)
    tax_rate             : Effective / marginal tax rate
    debt_weight          : D / (D + E) — market value weights
    equity_weight        : Optional override; if None, uses 1 - debt_weight
    size_premium         : Optional Duff & Phelps size premium
    specific_risk        : Optional company-specific risk premium
    """
    risk_free_rate: float
    equity_risk_premium: float
    beta: float
    cost_of_debt: float
    tax_rate: float
    debt_weight: float
    equity_weight: Optional[float] = None
    size_premium: float = 0.0
    specific_risk: float = 0.0

    def __post_init__(self):
        if self.equity_weight is None:
            self.equity_weight = 1.0 - self.debt_weight
        self._validate()

    def _validate(self):
        if not (0 < self.risk_free_rate < 1):
            raise ValueError(f"risk_free_rate must be between 0 and 1, got {self.risk_free_rate}")
        if not (0 <= self.tax_rate < 1):
            raise ValueError(f"tax_rate must be between 0 and 1, got {self.tax_rate}")
        if abs(self.debt_weight + self.equity_weight - 1.0) > 1e-6:
            raise ValueError("debt_weight + equity_weight must equal 1.0")

    @property
    def cost_of_equity(self) -> float:
        """CAPM: Ke = Rf + β × ERP + size_premium + specific_risk"""
        return (
            self.risk_free_rate
            + self.beta * self.equity_risk_premium
            + self.size_premium
            + self.specific_risk
        )

    @property
    def after_tax_cost_of_debt(self) -> float:
        """Kd × (1 - T)"""
        return self.cost_of_debt * (1 - self.tax_rate)

    def compute(self) -> float:
        """Return WACC as a decimal (e.g. 0.0847 = 8.47%)"""
        wacc = (
            self.cost_of_equity * self.equity_weight
            + self.after_tax_cost_of_debt * self.debt_weight
        )
        return round(wacc, 6)

    def sensitivity_table(
        self,
        beta_range: tuple = (0.7, 1.5),
        erp_range: tuple = (0.04, 0.07),
        steps: int = 5,
    ) -> dict:
        """
        2-D sensitivity: WACC as function of beta × ERP.
        Returns dict of {(beta, erp): wacc}.
        """
        betas = np.linspace(*beta_range, steps)
        erps = np.linspace(*erp_range, steps)
        table = {}
        for b in betas:
            for e in erps:
                calc = WACCCalculator(
                    risk_free_rate=self.risk_free_rate,
                    equity_risk_premium=e,
                    beta=b,
                    cost_of_debt=self.cost_of_debt,
                    tax_rate=self.tax_rate,
                    debt_weight=self.debt_weight,
                )
                table[(round(b, 2), round(e, 3))] = calc.compute()
        return table

    def summary(self) -> dict:
        return {
            "cost_of_equity": f"{self.cost_of_equity:.2%}",
            "after_tax_cost_of_debt": f"{self.after_tax_cost_of_debt:.2%}",
            "equity_weight": f"{self.equity_weight:.1%}",
            "debt_weight": f"{self.debt_weight:.1%}",
            "wacc": f"{self.compute():.2%}",
        }

    def __repr__(self):
        return (
            f"WACCCalculator(wacc={self.compute():.2%}, "
            f"Ke={self.cost_of_equity:.2%}, "
            f"Kd(at)={self.after_tax_cost_of_debt:.2%})"
        )
