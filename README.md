# Portfolio Dashboard

Interactive portfolio analytics dashboard that reads an Excel workbook, fetches live market data from Yahoo Finance, and generates Bloomberg-style HTML reports.

## Features

- **Positions** — Stock/ETF/cash holdings with market values (CAD/USD), weights, beta, industry
- **Options** — Live option prices from yfinance chains, contract values, delta exposure analysis
- **Correlation Matrix** — Interactive heatmap with click-to-sort and hover tooltips
- **Risk Metrics** — VaR, Sharpe, Sortino, Calmar, drawdown, beta (CAD + USD). Beta computed from SPY returns for ETFs/crypto where yfinance lacks it
- **Stress Testing** — 14 scenarios (-50% to +50%) with option hedging impact. USD/CAD toggle on dollar values
- **Exposure** — Sector, currency, and account breakdowns including option contract values
- **Privacy mode** — Hide all dollar amounts with one click
- **Language toggle** — English / Traditional Chinese (繁體中文)

## Quick Start

```bash
uv sync
uv run python build_portfolio_report.py
```

Open `index.html` in a browser.

## Input

`sample_portfolio.xlsx` with three sheets:

| Sheet | Required Columns |
|-------|-----------------|
| Portfolio | Account, Symbol, Shares, Price, Currency, Mkt Value, Mkt Value (CAD) |
| Options | Symbol, Account, Expirty, Type, Strike, Shares, Price, Currency |
| Currency | Price (USD/CAD rate in row 1) |

Fundamentals (sector, industry, beta) are fetched from yfinance — no Fundamental sheet needed.

## Output

7 HTML reports + 1 JSON file in the workspace root:

`index.html` · `positions.html` · `options.html` · `correlation_matrix.html` · `risk_metrics.html` · `stress_testing.html` · `sector_exposure.html` · `risk_metrics.json`

## Dependencies

- Python ≥ 3.11
- numpy, pandas, yfinance, scipy, openpyxl (managed via `uv sync`)
