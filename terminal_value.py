"""
Terminal Value Calculator
Supports Gordon Growth Model and Exit Multiple Method.
"""

from dataclasses import dataclass
from typing import Literal
import numpy as np


@dataclass
class TerminalValueCalculator:
    """
    Computes terminal value using two methods:

    1. Gordon Growth Model (Perpetuity):
       TV = FCFn × (1 + g) / (WACC - g)

    2. Exit Multiple Method:
       TV = EBITDAn × Exit Multiple

    Parameters
    ----------
    method           : "gordon_growth" or "exit_multiple"
    terminal_growth  : Long-term FCF growth rate (Gordon Growth)
    exit_multiple    : EV/EBITDA exit multiple
    """
    method: Literal["gordon_growth", "exit_multiple"] = "gordon_growth"
    terminal_growth: float = 0.025
    exit_multiple: float = 12.0

    def compute_gordon(self, last_fcf: float, wacc: float) -> float:
        """TV = FCFn × (1+g) / (WACC - g)"""
        if wacc <= self.terminal_growth:
            raise ValueError(
                f"WACC ({wacc:.2%}) must exceed terminal growth ({self.terminal_growth:.2%})"
            )
        return last_fcf * (1 + self.terminal_growth) / (wacc - self.terminal_growth)

    def compute_exit_multiple(self, last_ebitda: float) -> float:
        """TV = EBITDAn × Exit Multiple"""
        return last_ebitda * self.exit_multiple

    def compute(self, last_fcf: float, wacc: float, last_ebitda: float = None) -> float:
        """Dispatch to selected method."""
        if self.method == "gordon_growth":
            return self.compute_gordon(last_fcf, wacc)
        elif self.method == "exit_multiple":
            if last_ebitda is None:
                raise ValueError("last_ebitda required for exit_multiple method")
            return self.compute_exit_multiple(last_ebitda)
        else:
            raise ValueError(f"Unknown method: {self.method}")

    def pv(self, terminal_value: float, wacc: float, years: int) -> float:
        """Discount terminal value to present: TV / (1+WACC)^n"""
        return round(terminal_value / (1 + wacc) ** years, 2)

    def tv_as_pct_of_ev(
        self, pv_tv: float, pv_fcfs: float
    ) -> float:
        """TV contribution as % of total enterprise value."""
        total = pv_tv + pv_fcfs
        return pv_tv / total if total != 0 else 0.0

    def sensitivity_table(
        self,
        last_fcf: float,
        wacc: float,
        growth_range: tuple = (0.01, 0.04),
        multiple_range: tuple = (8.0, 16.0),
        steps: int = 5,
    ) -> dict:
        """Sensitivity: TV across growth rates (GGM) or exit multiples."""
        table = {}
        if self.method == "gordon_growth":
            import numpy as np
            for g in np.linspace(*growth_range, steps):
                calc = TerminalValueCalculator(method="gordon_growth", terminal_growth=g)
                table[f"g={g:.2%}"] = round(calc.compute_gordon(last_fcf, wacc), 1)
        else:
            import numpy as np
            for m in np.linspace(*multiple_range, steps):
                table[f"{m:.1f}x"] = round(self.compute_exit_multiple(last_fcf) * m / self.exit_multiple, 1)
        return table
