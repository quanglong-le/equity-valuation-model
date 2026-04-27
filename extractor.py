"""
Financial Statement Extractor
Fetches and parses income statement, balance sheet, and cash flow
from yfinance (Yahoo Finance) as a free data source.
"""

from dataclasses import dataclass, field
from typing import Optional, Dict
import pandas as pd
import numpy as np

try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False


@dataclass
class FinancialStatements:
    """Container for extracted financial statements."""
    ticker: str
    income_statement: pd.DataFrame = field(default_factory=pd.DataFrame)
    balance_sheet: pd.DataFrame = field(default_factory=pd.DataFrame)
    cash_flow: pd.DataFrame = field(default_factory=pd.DataFrame)
    info: dict = field(default_factory=dict)

    @property
    def is_empty(self) -> bool:
        return self.income_statement.empty and self.balance_sheet.empty


class FinancialExtractor:
    """
    Extracts financial statements for a given ticker.
    Supports Yahoo Finance (yfinance) as default source.
    """

    def __init__(self, ticker: str, source: str = "yfinance"):
        self.ticker = ticker.upper()
        self.source = source
        self._raw = None

    def fetch(self) -> FinancialStatements:
        """Download and return structured financial statements."""
        if self.source == "yfinance":
            return self._fetch_yfinance()
        raise ValueError(f"Unsupported source: {self.source}")

    def _fetch_yfinance(self) -> FinancialStatements:
        if not YFINANCE_AVAILABLE:
            raise ImportError("yfinance is not installed. Run: pip install yfinance")

        stock = yf.Ticker(self.ticker)
        fs = FinancialStatements(ticker=self.ticker)

        try:
            fs.income_statement = stock.financials.T
            fs.income_statement.index = pd.to_datetime(fs.income_statement.index).year
        except Exception:
            fs.income_statement = pd.DataFrame()

        try:
            fs.balance_sheet = stock.balance_sheet.T
            fs.balance_sheet.index = pd.to_datetime(fs.balance_sheet.index).year
        except Exception:
            fs.balance_sheet = pd.DataFrame()

        try:
            fs.cash_flow = stock.cashflow.T
            fs.cash_flow.index = pd.to_datetime(fs.cash_flow.index).year
        except Exception:
            fs.cash_flow = pd.DataFrame()

        try:
            fs.info = stock.info
        except Exception:
            fs.info = {}

        return fs

    @staticmethod
    def sample_statements(ticker: str = "DEMO") -> FinancialStatements:
        """
        Returns realistic sample financial statements for testing/demo purposes.
        Useful when no internet access or no API key.
        """
        years = [2021, 2022, 2023, 2024]

        income = pd.DataFrame({
            "Total Revenue": [8_500, 9_800, 11_200, 12_600],
            "Cost Of Revenue": [4_080, 4_606, 5_152, 5_670],
            "Gross Profit": [4_420, 5_194, 6_048, 6_930],
            "Operating Expense": [1_700, 1_960, 2_240, 2_520],
            "EBITDA": [2_720, 3_234, 3_808, 4_410],
            "Depreciation": [425, 490, 560, 630],
            "EBIT": [2_295, 2_744, 3_248, 3_780],
            "Interest Expense": [180, 195, 210, 220],
            "Pretax Income": [2_115, 2_549, 3_038, 3_560],
            "Tax Provision": [444, 535, 638, 748],
            "Net Income": [1_671, 2_014, 2_400, 2_812],
        }, index=years)

        balance = pd.DataFrame({
            "Cash And Cash Equivalents": [1_200, 1_450, 1_800, 2_100],
            "Total Current Assets": [3_800, 4_400, 5_200, 6_000],
            "Total Assets": [18_500, 20_800, 23_500, 26_200],
            "Total Current Liabilities": [2_400, 2_700, 3_100, 3_500],
            "Long Term Debt": [4_200, 4_000, 3_800, 3_600],
            "Total Debt": [4_800, 4_600, 4_400, 4_200],
            "Total Liabilities": [9_200, 10_000, 11_000, 12_000],
            "Total Stockholder Equity": [9_300, 10_800, 12_500, 14_200],
        }, index=years)

        cash_flow = pd.DataFrame({
            "Net Income": [1_671, 2_014, 2_400, 2_812],
            "Depreciation": [425, 490, 560, 630],
            "Change In Working Capital": [-180, -210, -250, -280],
            "Total Cash From Operating Activities": [1_916, 2_294, 2_710, 3_162],
            "Capital Expenditures": [-510, -588, -672, -756],
            "Free Cash Flow": [1_406, 1_706, 2_038, 2_406],
            "Total Cash From Investing Activities": [-650, -740, -840, -930],
            "Dividends Paid": [-200, -240, -280, -320],
            "Total Cash From Financing Activities": [-600, [-650], [-700], [-750]][0],
        }, index=years)

        info = {
            "shortName": f"{ticker} Corp.",
            "sector": "Technology",
            "industry": "Software",
            "country": "United States",
            "marketCap": 45_000_000_000,
            "sharesOutstanding": 500_000_000,
            "beta": 1.15,
        }

        return FinancialStatements(
            ticker=ticker,
            income_statement=income,
            balance_sheet=balance,
            cash_flow=cash_flow,
            info=info,
        )
