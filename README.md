# Portfolio Dashboard

**Live demo:** [chanmainvest.github.io/portfolio_dashboard](https://chanmainvest.github.io/portfolio_dashboard/)

Interactive portfolio analytics dashboard that reads an Excel workbook, fetches live market data from Yahoo Finance, and generates a Bloomberg-style single-page HTML report.

## Features

- **Single-page app (SPA)** — All analytics tabs in one `index.html` with hash-based navigation (`#dashboard`, `#positions`, `#options`, `#correlation`, `#risk`, `#stress`, `#exposure`)
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

1 HTML report + 1 JSON file in the workspace root:

`index.html` · `risk_metrics.json`

## Dependencies

- Python ≥ 3.11
- numpy, pandas, yfinance, scipy, openpyxl (managed via `uv sync`)

## Disclaimer

This project is provided for educational and informational purposes only and does **not** constitute investment, financial, legal, or tax advice.

Nothing in this repository (including code, reports, examples, outputs, comments, or documentation) is a recommendation, solicitation, endorsement, or offer to buy or sell any security, option, or other financial instrument.

You are solely responsible for your own investment decisions. Always conduct your own research and consult a licensed financial advisor or other qualified professional before making any investment decision.

Past performance is not indicative of future results. Markets are volatile, and all investing involves risk, including the possible loss of principal.
