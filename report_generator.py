"""
Valuation Report Generator
Produces a professional PDF due-diligence valuation report.
Uses matplotlib for charts and reportlab for PDF assembly.
"""

import os
from datetime import date
from typing import Optional
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table,
        TableStyle, Image as RLImage, PageBreak, HRFlowable
    )
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


class ValuationReportGenerator:
    """
    Generates a multi-page PDF valuation report with:
    - Executive Summary
    - Financial Statement Analysis
    - DCF Valuation
    - Comparable Company Analysis
    - Football Field Chart
    - Sensitivity Tables
    """

    BRAND_DARK = colors.HexColor("#0D1B2A")
    BRAND_BLUE = colors.HexColor("#1E6FBF")
    BRAND_ACCENT = colors.HexColor("#F0A500")
    BRAND_LIGHT = colors.HexColor("#F5F7FA")

    def __init__(self, output_path: str):
        self.output_path = output_path
        os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else ".", exist_ok=True)
        self._tmp_charts = []

    # ── chart helpers ─────────────────────────────────────────────────────────

    def _revenue_ebitda_chart(self, income_df: pd.DataFrame) -> str:
        fig, ax1 = plt.subplots(figsize=(7, 3.5))
        fig.patch.set_facecolor("#F5F7FA")
        ax1.set_facecolor("#F5F7FA")

        years = income_df.index.tolist()

        if "Revenue" in income_df.columns:
            rev = income_df["Revenue"].values / 1000
            ax1.bar(years, rev, color="#1E6FBF", alpha=0.85, label="Revenue ($B)")
            ax1.set_ylabel("Revenue ($B)", color="#1E6FBF", fontsize=9)

        ax2 = ax1.twinx()
        if "EBITDA" in income_df.columns:
            ebitda_margin = []
            for _, row in income_df.iterrows():
                m_str = row.get("EBITDA Margin", "n/a")
                try:
                    ebitda_margin.append(float(m_str.strip("%")) / 100)
                except Exception:
                    ebitda_margin.append(np.nan)
            ax2.plot(years, ebitda_margin, color="#F0A500", marker="o", linewidth=2, label="EBITDA Margin")
            ax2.set_ylabel("EBITDA Margin", color="#F0A500", fontsize=9)
            ax2.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))

        ax1.set_title("Revenue & EBITDA Margin Trend", fontsize=11, fontweight="bold", pad=10)
        ax1.set_xticks(years)
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=8)
        plt.tight_layout()

        path = "/tmp/_chart_rev_ebitda.png"
        plt.savefig(path, dpi=130, bbox_inches="tight")
        plt.close()
        self._tmp_charts.append(path)
        return path

    def _football_field_chart(self, ff_data: pd.DataFrame) -> str:
        fig, ax = plt.subplots(figsize=(7, 3))
        fig.patch.set_facecolor("#F5F7FA")
        ax.set_facecolor("#F5F7FA")

        colors_list = ["#1E6FBF", "#2ECC71", "#E74C3C", "#9B59B6"]
        for i, row in ff_data.iterrows():
            width = row["High"] - row["Low"]
            ax.barh(row["Method"], width, left=row["Low"],
                    color=colors_list[i % len(colors_list)], alpha=0.8, height=0.5)

        ax.set_xlabel("Implied Equity Value ($M)", fontsize=9)
        ax.set_title("Football Field — Valuation Range Summary", fontsize=11, fontweight="bold")
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
        ax.grid(axis="x", alpha=0.3)
        plt.tight_layout()

        path = "/tmp/_chart_football.png"
        plt.savefig(path, dpi=130, bbox_inches="tight")
        plt.close()
        self._tmp_charts.append(path)
        return path

    def _sensitivity_chart(self, sensitivity_data: dict, title: str = "DCF Sensitivity") -> str:
        if not sensitivity_data:
            return None
        labels = list(sensitivity_data.keys())
        values = list(sensitivity_data.values())

        fig, ax = plt.subplots(figsize=(6, 3))
        fig.patch.set_facecolor("#F5F7FA")
        ax.set_facecolor("#F5F7FA")
        bars = ax.bar(labels, values, color="#1E6FBF", alpha=0.8)
        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.set_ylabel("Terminal Value ($M)", fontsize=9)
        ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.01,
                    f"${val:,.0f}", ha="center", va="bottom", fontsize=7)
        plt.tight_layout()

        path = "/tmp/_chart_sensitivity.png"
        plt.savefig(path, dpi=130, bbox_inches="tight")
        plt.close()
        self._tmp_charts.append(path)
        return path

    # ── table helpers ─────────────────────────────────────────────────────────

    def _df_to_rl_table(self, df: pd.DataFrame, col_widths=None) -> Table:
        data = [[""] + list(df.columns)] if df.index.name else [list(df.columns)]
        if df.index.name or df.index.dtype != "int64":
            data = [[str(df.index.name or "Year")] + list(df.columns)]
            for idx, row in df.iterrows():
                data.append([str(idx)] + [str(v) for v in row.values])
        else:
            data = [list(df.columns)]
            for _, row in df.iterrows():
                data.append([str(v) for v in row.values])

        style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), self.BRAND_BLUE),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [self.BRAND_LIGHT, colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ])

        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(style)
        return t

    # ── main report ───────────────────────────────────────────────────────────

    def generate(
        self,
        ticker: str,
        company_name: str,
        income_df: pd.DataFrame,
        balance_df: pd.DataFrame,
        cashflow_df: pd.DataFrame,
        dcf_summary: dict,
        multiples_df: pd.DataFrame,
        football_field_df: pd.DataFrame,
        sensitivity_data: dict = None,
        analyst_notes: str = "",
    ):
        if not REPORTLAB_AVAILABLE:
            raise ImportError("reportlab not installed. Run: pip install reportlab")

        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=A4,
            rightMargin=2*cm, leftMargin=2*cm,
            topMargin=2*cm, bottomMargin=2*cm,
        )

        styles = getSampleStyleSheet()
        h1 = ParagraphStyle("H1", parent=styles["Title"], fontSize=20, textColor=self.BRAND_DARK,
                             spaceAfter=6, fontName="Helvetica-Bold")
        h2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=13, textColor=self.BRAND_BLUE,
                             spaceBefore=14, spaceAfter=4, fontName="Helvetica-Bold")
        body = ParagraphStyle("Body", parent=styles["Normal"], fontSize=9, leading=14)
        small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=7.5,
                               textColor=colors.HexColor("#666666"))

        story = []

        # ── Cover ──
        story.append(Spacer(1, 1.5*cm))
        story.append(Paragraph(f"Equity Valuation Report", h1))
        story.append(Paragraph(f"<b>{company_name}</b> ({ticker})", h2))
        story.append(HRFlowable(width="100%", thickness=2, color=self.BRAND_ACCENT))
        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph(f"Prepared: {date.today().strftime('%B %d, %Y')}", small))
        story.append(Paragraph("Methodology: DCF (WACC + Terminal Value) | Market Multiples (CCA)", small))
        story.append(Spacer(1, 0.5*cm))

        # ── Executive Summary ──
        story.append(Paragraph("Executive Summary", h2))
        ev_low = dcf_summary.get("equity_value_low", "N/A")
        ev_high = dcf_summary.get("equity_value_high", "N/A")
        wacc_val = dcf_summary.get("wacc", "N/A")
        intrinsic = dcf_summary.get("intrinsic_equity_value", "N/A")

        summary_data = [
            ["Metric", "Value"],
            ["WACC", f"{wacc_val:.2%}" if isinstance(wacc_val, float) else str(wacc_val)],
            ["Intrinsic Equity Value (DCF)", f"${intrinsic:,.0f}M" if isinstance(intrinsic, (int,float)) else str(intrinsic)],
            ["Equity Value Range (Sensitivity)", f"${ev_low:,.0f}M – ${ev_high:,.0f}M" if isinstance(ev_low, (int,float)) else "N/A"],
            ["TV as % of EV", f"{dcf_summary.get('tv_pct', 0):.1%}"],
        ]
        t = Table(summary_data, colWidths=[8*cm, 8*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), self.BRAND_DARK),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [self.BRAND_LIGHT, colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
            ("ALIGN", (1, 0), (1, -1), "CENTER"),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(t)

        if analyst_notes:
            story.append(Spacer(1, 0.3*cm))
            story.append(Paragraph(f"<i>Analyst Notes: {analyst_notes}</i>", small))

        # ── Financial Analysis ──
        story.append(Paragraph("Financial Statement Analysis", h2))

        if not income_df.empty:
            rev_chart = self._revenue_ebitda_chart(income_df)
            story.append(RLImage(rev_chart, width=14*cm, height=7*cm))
            story.append(Spacer(1, 0.3*cm))
            story.append(Paragraph("Income Statement Summary", ParagraphStyle("sub", parent=body, fontName="Helvetica-Bold")))
            story.append(self._df_to_rl_table(income_df.reset_index()))
            story.append(Spacer(1, 0.3*cm))

        if not balance_df.empty:
            story.append(Paragraph("Balance Sheet Key Metrics", ParagraphStyle("sub", parent=body, fontName="Helvetica-Bold")))
            story.append(self._df_to_rl_table(balance_df.reset_index()))
            story.append(Spacer(1, 0.3*cm))

        if not cashflow_df.empty:
            story.append(Paragraph("Cash Flow Summary", ParagraphStyle("sub", parent=body, fontName="Helvetica-Bold")))
            story.append(self._df_to_rl_table(cashflow_df.reset_index()))

        story.append(PageBreak())

        # ── DCF ──
        story.append(Paragraph("DCF Valuation", h2))
        story.append(Paragraph(
            "Enterprise value derived by discounting projected free cash flows at WACC "
            "and adding the present value of terminal value.", body
        ))
        if dcf_summary.get("fcf_table") is not None:
            story.append(Spacer(1, 0.2*cm))
            story.append(self._df_to_rl_table(dcf_summary["fcf_table"].reset_index()))

        if sensitivity_data:
            story.append(Spacer(1, 0.3*cm))
            sens_chart = self._sensitivity_chart(sensitivity_data)
            if sens_chart:
                story.append(RLImage(sens_chart, width=12*cm, height=6*cm))

        # ── Multiples ──
        story.append(Paragraph("Comparable Company Analysis", h2))
        if not multiples_df.empty:
            story.append(self._df_to_rl_table(multiples_df))

        # ── Football Field ──
        if not football_field_df.empty:
            story.append(Paragraph("Football Field — Valuation Summary", h2))
            ff_chart = self._football_field_chart(football_field_df)
            story.append(RLImage(ff_chart, width=14*cm, height=6*cm))

        # ── Disclaimer ──
        story.append(Spacer(1, 1*cm))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
        story.append(Paragraph(
            "Disclaimer: This report is for informational purposes only and does not constitute investment advice. "
            "Projections and valuations involve significant uncertainty. Past performance is not indicative of future results.",
            small
        ))

        doc.build(story)
        print(f"✅ Report generated → {self.output_path}")

        # cleanup tmp charts
        for f in self._tmp_charts:
            try:
                os.remove(f)
            except Exception:
                pass

        return self.output_path
