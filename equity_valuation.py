"""
Equity Valuation Model – DCF + Comparable Companies (Comps)
============================================================
Author  : Quang Long LE
Date    : 2026
Methods : Discounted Cash Flow (DCF) + Trading Multiples (EV/EBITDA, P/E)
"""

import numpy as np
import pandas as pd

# ──────────────────────────────────────────────────────────────────────────────
# 1. INPUT PARAMETERS
# ──────────────────────────────────────────────────────────────────────────────

company = {
    "name"          : "Example Corp",
    "revenue"       : 1_000,   # in M€
    "ebitda"        : 250,     # in M€
    "ebit"          : 180,     # in M€
    "net_income"    : 120,     # in M€
    "shares"        : 50,      # millions of shares outstanding
    "net_debt"      : 200,     # in M€
    "tax_rate"      : 0.25,    # corporate tax rate
    "capex"         : 60,      # in M€
    "d_and_a"       : 70,      # depreciation & amortisation in M€
    "change_in_wc"  : 20,      # change in working capital in M€
}

dcf_params = {
    "revenue_growth"   : [0.08, 0.07, 0.06, 0.05, 0.05],  # 5-year growth rates
    "ebitda_margin"    : 0.25,     # stable EBITDA margin assumption
    "wacc"             : 0.09,     # weighted average cost of capital
    "terminal_growth"  : 0.02,     # Gordon growth model terminal rate
    "projection_years" : 5,
}

comps = pd.DataFrame({
    "Company"       : ["Peer A", "Peer B", "Peer C", "Peer D"],
    "EV_EBITDA"     : [10.0,     12.5,     9.5,      11.0    ],
    "P_E"           : [18.0,     22.0,     16.5,      20.0   ],
})

# ──────────────────────────────────────────────────────────────────────────────
# 2. DCF VALUATION
# ──────────────────────────────────────────────────────────────────────────────

def project_fcf(company: dict, dcf_params: dict) -> pd.DataFrame:
    """Project Free Cash Flows over the forecast period."""
    rows = []
    revenue = company["revenue"]

    for i, g in enumerate(dcf_params["revenue_growth"], start=1):
        revenue     *= (1 + g)
        ebitda       = revenue * dcf_params["ebitda_margin"]
        ebit         = ebitda  - company["d_and_a"]
        nopat        = ebit    * (1 - company["tax_rate"])
        fcf          = nopat + company["d_and_a"] - company["capex"] - company["change_in_wc"]

        rows.append({
            "Year"      : f"Y{i}",
            "Revenue"   : round(revenue, 1),
            "EBITDA"    : round(ebitda,  1),
            "EBIT"      : round(ebit,    1),
            "NOPAT"     : round(nopat,   1),
            "FCF"       : round(fcf,     1),
        })

    return pd.DataFrame(rows)


def compute_wacc_sensitivity(base_wacc: float,
                              steps: list[float] = [-0.02, -0.01, 0, 0.01, 0.02]
                              ) -> list[float]:
    """Return a range of WACC values for sensitivity analysis."""
    return [round(base_wacc + s, 4) for s in steps]


def dcf_valuation(company: dict, dcf_params: dict) -> dict:
    """
    Compute DCF equity value per share.

    Returns
    -------
    dict with:
        - projected FCFs
        - PV of FCFs
        - terminal value
        - enterprise value
        - equity value per share
        - sensitivity table (WACC × terminal growth)
    """
    proj = project_fcf(company, dcf_params)
    wacc = dcf_params["wacc"]
    tg   = dcf_params["terminal_growth"]

    # Present value of projected FCFs
    pv_fcfs = [
        fcf / (1 + wacc) ** (i + 1)
        for i, fcf in enumerate(proj["FCF"])
    ]
    pv_fcf_total = sum(pv_fcfs)

    # Terminal value (Gordon Growth Model) and its PV
    terminal_fcf = proj["FCF"].iloc[-1] * (1 + tg)
    terminal_value = terminal_fcf / (wacc - tg)
    pv_terminal = terminal_value / (1 + wacc) ** dcf_params["projection_years"]

    # Enterprise value → Equity value
    enterprise_value = pv_fcf_total + pv_terminal
    equity_value     = enterprise_value - company["net_debt"]
    price_per_share  = equity_value / company["shares"]

    # ── Sensitivity table: WACC × Terminal Growth Rate ──────────────────────
    wacc_range = [wacc + s for s in [-0.02, -0.01, 0, 0.01, 0.02]]
    tg_range   = [tg   + s for s in [-0.01,  0,    0.01, 0.02   ]]

    sensitivity = pd.DataFrame(index=[f"{w:.1%}" for w in wacc_range],
                               columns=[f"{t:.1%}" for t in tg_range])
    for w in wacc_range:
        for t in tg_range:
            if w > t:
                tv   = proj["FCF"].iloc[-1] * (1 + t) / (w - t)
                pv_t = tv / (1 + w) ** dcf_params["projection_years"]
                pv_f = sum(proj["FCF"].iloc[i] / (1 + w) ** (i + 1)
                           for i in range(len(proj)))
                ev   = pv_f + pv_t
                eq   = ev - company["net_debt"]
                sensitivity.loc[f"{w:.1%}", f"{t:.1%}"] = round(eq / company["shares"], 1)
            else:
                sensitivity.loc[f"{w:.1%}", f"{t:.1%}"] = "N/A"

    return {
        "projections"       : proj,
        "pv_fcf_total"      : round(pv_fcf_total, 1),
        "terminal_value"    : round(terminal_value, 1),
        "pv_terminal"       : round(pv_terminal, 1),
        "enterprise_value"  : round(enterprise_value, 1),
        "equity_value"      : round(equity_value, 1),
        "price_per_share"   : round(price_per_share, 2),
        "sensitivity"       : sensitivity,
    }

# ──────────────────────────────────────────────────────────────────────────────
# 3. COMPARABLE COMPANIES (COMPS)
# ──────────────────────────────────────────────────────────────────────────────

def comps_valuation(company: dict, comps: pd.DataFrame) -> dict:
    """
    Derive implied equity value from trading multiples.

    Methods
    -------
    - EV/EBITDA median → Enterprise Value → Equity Value
    - P/E median       → Equity Value directly
    """
    median_ev_ebitda = comps["EV_EBITDA"].median()
    median_pe        = comps["P_E"].median()

    # EV/EBITDA implied value
    implied_ev_ebitda  = median_ev_ebitda * company["ebitda"]
    implied_eq_ebitda  = implied_ev_ebitda - company["net_debt"]
    implied_px_ebitda  = implied_eq_ebitda / company["shares"]

    # P/E implied value
    eps               = company["net_income"] / company["shares"]
    implied_px_pe     = median_pe * eps

    return {
        "median_EV_EBITDA"          : median_ev_ebitda,
        "implied_EV_from_EBITDA"    : round(implied_ev_ebitda, 1),
        "implied_price_EV_EBITDA"   : round(implied_px_ebitda, 2),
        "median_PE"                 : median_pe,
        "EPS"                       : round(eps, 2),
        "implied_price_PE"          : round(implied_px_pe, 2),
    }

# ──────────────────────────────────────────────────────────────────────────────
# 4. FOOTBALL FIELD (VALUATION SUMMARY)
# ──────────────────────────────────────────────────────────────────────────────

def football_field(dcf_result: dict, comps_result: dict) -> pd.DataFrame:
    """Aggregate all valuation methods into a summary table (football field)."""
    dcf_price      = dcf_result["price_per_share"]
    ebitda_price   = comps_result["implied_price_EV_EBITDA"]
    pe_price       = comps_result["implied_price_PE"]

    # Simple ±10% range around each point estimate
    def rng(x, pct=0.10):
        return (round(x * (1 - pct), 2), round(x * (1 + pct), 2))

    rows = [
        {"Method": "DCF",        "Low": rng(dcf_price)[0],    "Mid": dcf_price,    "High": rng(dcf_price)[1]},
        {"Method": "EV/EBITDA",  "Low": rng(ebitda_price)[0], "Mid": ebitda_price, "High": rng(ebitda_price)[1]},
        {"Method": "P/E",        "Low": rng(pe_price)[0],      "Mid": pe_price,    "High": rng(pe_price)[1]},
    ]
    df = pd.DataFrame(rows).set_index("Method")

    all_mids = [dcf_price, ebitda_price, pe_price]
    df.loc["Implied range (avg)"] = {
        "Low" : round(min(all_mids) * 0.90, 2),
        "Mid" : round(np.mean(all_mids), 2),
        "High": round(max(all_mids) * 1.10, 2),
    }
    return df

# ──────────────────────────────────────────────────────────────────────────────
# 5. MAIN – DASHBOARD OUTPUT
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":

    SEP = "─" * 60

    # ── DCF ──────────────────────────────────────────────────
    dcf = dcf_valuation(company, dcf_params)

    print(f"\n{'═'*60}")
    print(f"  EQUITY VALUATION MODEL  |  {company['name']}")
    print(f"{'═'*60}")

    print(f"\n{SEP}")
    print("  1. PROJECTED FREE CASH FLOWS  (M€)")
    print(SEP)
    print(dcf["projections"].to_string(index=False))

    print(f"\n{SEP}")
    print("  2. DCF VALUATION  (M€)")
    print(SEP)
    print(f"  PV of FCFs          : {dcf['pv_fcf_total']:>10,.1f} M€")
    print(f"  Terminal Value (TV) : {dcf['terminal_value']:>10,.1f} M€")
    print(f"  PV of TV            : {dcf['pv_terminal']:>10,.1f} M€")
    print(f"  Enterprise Value    : {dcf['enterprise_value']:>10,.1f} M€")
    print(f"  Equity Value        : {dcf['equity_value']:>10,.1f} M€")
    print(f"  ➜  Price per share  : {dcf['price_per_share']:>10,.2f} €")

    print(f"\n{SEP}")
    print("  3. SENSITIVITY – Price/share (€)  |  WACC × Terminal Growth")
    print(SEP)
    print(dcf["sensitivity"].to_string())

    # ── COMPS ─────────────────────────────────────────────────
    comps_res = comps_valuation(company, comps)

    print(f"\n{SEP}")
    print("  4. COMPARABLE COMPANIES  (Comps)")
    print(SEP)
    print(comps.to_string(index=False))
    print(f"\n  Median EV/EBITDA    : {comps_res['median_EV_EBITDA']:.1f}x")
    print(f"  ➜  Implied price    : {comps_res['implied_price_EV_EBITDA']:>6.2f} €")
    print(f"\n  Median P/E          : {comps_res['median_PE']:.1f}x  |  EPS: {comps_res['EPS']:.2f} €")
    print(f"  ➜  Implied price    : {comps_res['implied_price_PE']:>6.2f} €")

    # ── FOOTBALL FIELD ────────────────────────────────────────
    ff = football_field(dcf, comps_res)

    print(f"\n{SEP}")
    print("  5. FOOTBALL FIELD – Valuation Summary  (€/share)")
    print(SEP)
    print(ff.to_string())
    print(f"\n{'═'*60}\n")
