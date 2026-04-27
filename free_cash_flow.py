"""
Free Cash Flow Projector
Projects unlevered FCF and discounts back to present value.

FCF = EBIT × (1 - Tax Rate) + D&A - ΔCapEx - ΔNWC
"""

from dataclasses import dataclass, field
from typing import List, Optional
import numpy as np
import pandas as pd


@dataclass
class FCFProjector:
    """
    Projects free cash flows over a forecast horizon.

    Parameters
    ----------
    base_revenue      : Last twelve months (LTM) revenue ($M)
    revenue_growth    : List of annual revenue growth rates (e.g. [0.12, 0.10, ...])
    ebitda_margin     : EBITDA margin assumption (stable or per-year list)
    da_pct_revenue    : D&A as % of revenue
    capex_pct         : CapEx as % of revenue
    nwc_change_pct    : Change in NWC as % of revenue delta
    tax_rate          : Effective tax rate
    """
    base_revenue: float
    revenue_growth: List[float]
    ebitda_margin: float | List[float]
    da_pct_revenue: float = 0.05
    capex_pct: float = 0.05
    nwc_change_pct: float = 0.02
    tax_rate: float = 0.21

    def __post_init__(self):
        self.n_years = len(self.revenue_growth)
        if isinstance(self.ebitda_margin, float):
            self.ebitda_margins = [self.ebitda_margin] * self.n_years
        else:
            self.ebitda_margins = self.ebitda_margin

    def _build_revenue(self) -> List[float]:
        revenues = []
        rev = self.base_revenue
        for g in self.revenue_growth:
            rev = rev * (1 + g)
            revenues.append(rev)
        return revenues

    def project(self) -> pd.DataFrame:
        """
        Returns a DataFrame with all projected line items.
        Columns: Year, Revenue, EBITDA, EBIT, NOPAT, DA, CapEx, ΔNWC, FCF
        """
        revenues = self._build_revenue()
        records = []

        prev_revenue = self.base_revenue
        for i, (rev, margin) in enumerate(zip(revenues, self.ebitda_margins)):
            ebitda = rev * margin
            da = rev * self.da_pct_revenue
            ebit = ebitda - da
            nopat = ebit * (1 - self.tax_rate)
            capex = rev * self.capex_pct
            delta_nwc = (rev - prev_revenue) * self.nwc_change_pct
            fcf = nopat + da - capex - delta_nwc

            records.append({
                "Year": i + 1,
                "Revenue ($M)": round(rev, 1),
                "Revenue Growth": f"{self.revenue_growth[i]:.1%}",
                "EBITDA ($M)": round(ebitda, 1),
                "EBITDA Margin": f"{margin:.1%}",
                "D&A ($M)": round(da, 1),
                "EBIT ($M)": round(ebit, 1),
                "NOPAT ($M)": round(nopat, 1),
                "CapEx ($M)": round(capex, 1),
                "ΔNWC ($M)": round(delta_nwc, 1),
                "FCF ($M)": round(fcf, 1),
            })
            prev_revenue = rev

        return pd.DataFrame(records).set_index("Year")

    def get_fcf_list(self) -> List[float]:
        """Return plain list of projected FCFs."""
        df = self.project()
        return df["FCF ($M)"].tolist()

    def discount_fcfs(self, fcfs: List[float], wacc: float) -> float:
        """PV of projected FCFs."""
        pv = sum(fcf / (1 + wacc) ** t for t, fcf in enumerate(fcfs, start=1))
        return round(pv, 2)

    def pv_fcfs(self, wacc: float) -> float:
        """Convenience: project + discount in one call."""
        return self.discount_fcfs(self.get_fcf_list(), wacc)

    def to_excel(self, path: str):
        """Export projection table to Excel."""
        df = self.project()
        df.to_excel(path, sheet_name="FCF Projection")
        print(f"FCF projection saved → {path}")
