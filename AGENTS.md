---
name: Stock Portfolio Tracker
description: Build an institutional-grade single-page portfolio analytics HTML report from an Excel workbook — risk metrics, correlation matrix, stress testing, sector exposure, and options delta hedging.
tools: ['vscode', 'execute', 'read', 'edit', 'search', 'web', 'duckduckgo/*', 'agent', 'todo']
---

> **Source:** This guide is adapted from [Building an Institutional-Grade Portfolio Tracker with AI](https://mispricedassets.substack.com/p/building-an-institutional-grade-portfolio) by Mispriced Assets.
Why You Should Spend a Couple Hours Learning This
Learning to use an AI coding agent is one of the highest-ROI skills you can develop right now. It feels like a video game.

Think about it: you describe what you want in plain English, and an AI builds it for you in real-time. You watch it write code, fetch data, create files, fix its own errors. When something isn’t quite right, you just say “make the headers darker” or “add a correlation matrix” and it does it. Almost instantaneously.

The investment: A couple hours to learn the basics and a subscription to an AI coding tool.

The payoff: You now have a personal AI developer who works 24/7, never gets tired, never judges your questions, and can build literally anything you can describe. Portfolio trackers today. Automated reports tomorrow. Data analysis tools next week. The skill compounds forever.

This guide will walk you through building one specific project - a portfolio tracker - but the real value is learning how to work with an AI coding agent. Once you get it, you’ll start seeing opportunities everywhere. “Wait, the AI could automate this.” “The AI could build me a tool for that.” It changes how you think about problems.

Trust me: the couple hours you spend on this will pay for itself many times over in time saved and opportunities captured.

> **Model Requirements (as of 2026/02/15):** Only use frontier-tier models — **Claude Opus 4.6** or **Codex 5.3**. You will be wasting your time if you use weaker LLM models; they lack the reasoning depth to produce correct financial calculations and multi-sheet report architectures. A **$20/month GitHub Copilot** plan works just as well as a direct Claude subscription for this task.

What We Are Building
By the end of this guide, you will have a fully automated portfolio analytics system that:

Fetches live stock prices automatically from Yahoo Finance

Calculates P&L, returns, and portfolio weights

Includes institutional-grade risk analytics (VaR, Sharpe Ratio, correlations)

Features Bloomberg Terminal-style professional formatting with interactive HTML reports

Updates with a single command — no manual data entry ever again

The best part? You do not need to know how to code. The AI writes all the code for you. You just need to describe what you want in plain English.

Why AI Coding Agents Are Different
Regular AI chatbots can only give you text responses. You ask for code, they show you code, and then you have to figure out how to run it yourself. An AI coding agent actually executes the code for you.

An AI coding agent can:

Read your existing files (like your trade log spreadsheet)

Write and execute Python code directly on your computer

Create reports with complex calculations and formatting

Fetch live data from the internet (stock prices, company info)

Make changes iteratively based on your feedback

Fix its own errors when something goes wrong

It is like having a developer sitting next to you, except this developer works instantly, never gets frustrated, and is available at 3am when you have a random idea.

Why Build a Custom Portfolio Tracker?
The Problem with Existing Solutions
Brokerage Platforms: Limited customization. You cannot add your own metrics, investment theses, or risk analytics. You are stuck with whatever they give you.

Google Sheets/Excel Templates: Static and dumb. They do not fetch live data. Every update requires manual copy-paste. They break constantly.

Portfolio Apps (Sharesight, Stock Events, etc.): Monthly subscription fees. Your data lives on their servers. You cannot see how calculations work. You cannot export your logic. If they shut down, you lose everything.

Bloomberg Terminal: $24,000 per year. Enough said.

The Solution: Your Own Glass Box System
With an AI coding agent, you build exactly what you want:

Full transparency — every calculation is in the source code you own

No subscriptions — runs on your computer, you own everything forever

Fully customizable — want a new metric? Just ask the AI to add it

Live updates — re-run the script and everything recalculates instantly

Professional output — Bloomberg-quality formatting without Bloomberg prices

Your data stays private — nothing leaves your computer

What You Need Before Starting
1. Python Installation
Python is a programming language that the AI agent will use to build your reports. Do not worry — you will not need to write any Python yourself. The AI does all of that.

Download from: https://www.python.org/downloads/

Windows Installation
Go to python.org and click “Download Python” (get the latest version, like 3.12)

Run the installer

⚠️ CRITICAL: Check the box that says “Add Python to PATH” during installation. If you miss this, nothing will work.

Click “Install Now”

Verify it worked: Open Command Prompt (search “cmd” in Windows) and type:

python --version
You should see something like “Python 3.12.0”

Mac Installation
Option A: Download from python.org

Go to https://www.python.org/downloads/

Download the macOS installer

Run the .pkg file and follow the prompts

Option B: Using Homebrew (recommended for Mac users)

brew install python
Verify it worked: Open Terminal (search “Terminal” in Spotlight) and type:

python3 --version
⚠️ Important Mac Note: On Mac, use python3 and pip3 instead of python and pip.

2. Your Trade Log Data
You have two options for getting your trade data into the system:

Option A: Manual Excel File (Tedious)

Create an Excel file (e.g., “My Portfolio.xlsx”) with a sheet named exactly Trade_Log containing these columns:

Trade_ID - Unique number for each trade (1, 2, 3...)

Date - Trade date (YYYY-MM-DD format, e.g., 2025-12-31)

Ticker - Stock symbol (e.g., AAPL, MSFT, GOOGL)

Action - BUY or SELL (must be capitalized)

Shares - Number of shares traded

Price - Price per share at time of trade

Gross_Value - Shares × Price

Commission - Trading fees (put 0 if no commission)

Net_Cost - Gross_Value + Commission for buys, Gross_Value - Commission for sells

Cash_Balance - Running cash balance after each trade

Notes - Optional notes (e.g., “Initial allocation”)

Option B: Connect to Your Broker API (Automated)

Most major brokers offer APIs that let you pull your trade history automatically. Here’s how to set this up with popular brokers:

Interactive Brokers (IBKR)

pip install ib_insync
Tell the AI:

Connect to my Interactive Brokers account using ib_insync and pull my trade history. My TWS is running on localhost port 7497.
TD Ameritrade / Charles Schwab

pip install schwab-py
Tell the AI:

Use the Schwab API to fetch my trade history. Here's my API key: [your key]. Save it in the Trade_Log format.
Robinhood (Unofficial)

pip install robin_stocks
Tell the AI:

Use robin_stocks to log into my Robinhood account and export my trade history. My credentials are [email] and I'll enter my MFA code when prompted.
Fidelity / Vanguard / Others

These brokers don’t have public APIs. Export your trades as CSV from their website, then tell the AI:

Read my exported trades from [path]/fidelity_trades.csv and convert them to the Trade_Log format.
Example Prompt for Broker Connection:

Connect to my Alpaca brokerage account and sync my trades:
- API Key: PKXXXXXXXXXXXXXXXX
- Secret Key: XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
- Use paper trading endpoint (or: use live endpoint)

Pull all trades from the last 12 months and format them into my Trade_Log sheet. Then fetch current positions and reconcile against my trade history.
⚠️ Security Note: Never share API keys in public channels. Consider using environment variables:

Tell the AI:

Read my Alpaca API credentials from environment variables ALPACA_API_KEY and ALPACA_SECRET_KEY
Setting Up Environment Variables:

Windows (Command Prompt):

set ALPACA_API_KEY=your_key_here
set ALPACA_SECRET_KEY=your_secret_here
Mac/Linux (Terminal):

export ALPACA_API_KEY=your_key_here
export ALPACA_SECRET_KEY=your_secret_here
Which Option Should You Choose?

Just getting started, few trades → Option A (Manual Excel)

Active trader, 10+ trades/month → Option B (Broker API)

Multiple brokers → Option A + manual consolidation

Want fully automated updates → Option B with daily sync script

Once your Trade_Log is populated (either way), the AI agent can build the rest of the system on top of it.

Information You Need to Gather Before Starting
Before launching the AI agent, gather this information. Having it ready will make the process much smoother.

Required Information
File Path to Your Trade Log: The exact location of your Excel file.

Windows Mac C:\Users\YourName\Desktop\My Portfolio.xlsx /Users/YourName/Desktop/My Portfolio.xlsx

Output File Path: Where should the AI save the new workbook? Use a DIFFERENT name to avoid overwriting.

Windows Mac C:\Users\YourName\Desktop\My Portfolio_BUILT.xlsx /Users/YourName/Desktop/My Portfolio_BUILT.xlsx

Starting Capital: How much cash did you start with? Example: $1,000,000

Start Date: When did you begin? This is usually when you made your first trades. Example: December 31, 2024

💡 Mac Tip: Getting File Paths
To get the exact path of any file on Mac:

Right-click the file in Finder

Hold Option key

Click “Copy as Pathname”

Step-by-Step Instructions
Step 1: Prepare Your Trade Log
Create an Excel file with your trades. Here is an example of what the Trade_Log sheet should look like:




⚠️ Important: The Cash_Balance column should show your running cash balance AFTER each trade. Start with your initial capital and subtract each BUY (add back each SELL).

Step 2: Give the AI Your Initial Prompt
This is the most important step. Copy the prompt below and modify the bracketed sections with your actual information.

⚠️ CRITICAL: Notice how the prompt explicitly tells the AI to use FORMULAS, not hardcoded values. This is essential.

I want you to build me a portfolio tracking spreadsheet. Here are my requirements:

**FILE LOCATIONS:**
- Input: [YOUR PATH]/My Portfolio.xlsx (contains Trade_Log sheet with my trades)
- Output: [YOUR PATH]/My Portfolio_BUILT.xlsx (create NEW file, do NOT overwrite my original)

**PORTFOLIO DETAILS:**
- Starting capital: $[YOUR AMOUNT] (e.g., $1,000,000)
- Start date: [YOUR DATE] (e.g., December 31, 2025)
- All initial positions bought at closing prices on the start date

**CRITICAL - FORMULA-BASED ARCHITECTURE:**
The Python script should ONLY:
1. Read my Trade_Log
2. Fetch market data via yfinance (prices, volumes, company info)
3. Build the Excel structure with FORMULAS

IMPORTANT: Use Excel FORMULAS for ALL calculations. Do NOT hardcode any calculated values.
Every P&L, every percentage, every total should be a formula like =D2*F2, not a static number.
The spreadsheet must update automatically when I change any price in Market_Data.

**SHEETS TO CREATE:**

Data Lakes (Python fills these with raw data):
- Trade_Log: Copy of my trades
- Market_Data: Daily closing prices (rows = dates, columns = tickers)
- Reference_Data: Company name, sector, industry, country, beta for each ticker
- Daily_Units: Number of shares held per ticker per day (calculated from trades)
- Daily_Cash: Cash balance per day

Analytics Sheets (ALL FORMULAS - no hardcoded values):
- Dashboard: KPIs (NAV, Total Return, Sharpe Ratio, Max Drawdown), NAV chart, allocation pie chart
- Positions: Ticker, Company, Sector, Shares, Avg Cost, Current Price, Day Change %, Cost Basis, Market Value, Unrealized P&L, P&L %, Portfolio Weight
- Equity_Curve: Date, Cash, Invested Value, Total NAV, Daily Return, Cumulative Return, Drawdown
- Sector_Exposure: Sector breakdown with dollar amounts and weights
- Geographic_Exposure: Country breakdown with dollar amounts and weights

**STYLING:**
- Professional Bloomberg Terminal look
- Segoe UI font, 85% zoom, hide gridlines
- Dark navy headers (#1C2541) with white text
- Alternating row colors (zebra striping) for readability
- Green (#007A33) for positive values, Red (#B81D13) for negative
- Freeze the header row and first column on each sheet

Please start by reading my Trade_Log file, then build the complete system.
Step 3: Watch the AI Work
This is the fun part. The AI agent will start working and you will see real-time output:

-> Reading Trade_Log...
-> Found 15 unique tickers: AAPL, MSFT, GOOGL...
-> Fetching market data from Yahoo Finance...
  + AAPL: 18 trading days
  + MSFT: 18 trading days
  ...
-> Fetching reference data...
  + AAPL: Apple Inc., Technology, United States
  ...
-> Building workbook...
-> Creating Dashboard sheet...
-> Creating Positions sheet with formulas...
-> Creating Equity_Curve sheet...
  ...
-> BUILD COMPLETE
-> Saved to: [Your output path]
The AI may ask permission to install Python packages (like yfinance, xlsxwriter). Say yes.

Step 4: Open and Verify the Formulas
Tell the AI:

open the spreadsheet
Excel will open. Now verify the formulas are working correctly:

Go to the Positions sheet

Click on any P&L cell

Look at the formula bar at the top - you should see a formula like =I2-H2, NOT just a number

Go to Market_Data sheet

Manually change one price to a different number

Go back to Positions - the P&L should have updated automatically

⚠️ If the P&L did NOT update: Tell the AI: “The P&L cells have hardcoded values. Rebuild with formulas instead.”

Enhancing Your Portfolio Tracker
Once you have the basic system working, you can keep asking the AI to add more features. This is where it gets really fun - you just describe what you want and watch it appear.

Adding Risk Analytics
Copy and paste this prompt:

Add the following risk analytics sheets to my portfolio:

- Correlation: A matrix showing the correlation between all my positions. Use 3-color conditional formatting (green for low correlation, yellow for medium, red for high).

- Risk_Metrics: A summary sheet with these metrics:
  - Annualized Return
  - Annualized Volatility
  - Sharpe Ratio
  - Sortino Ratio
  - Maximum Drawdown
  - Beta to SPY
  - VaR 95% (Value at Risk)
  - VaR 99%
  - CVaR 95% (Conditional VaR)

- Stress_Testing: Show what happens to my portfolio in these scenarios:
  - Market crash: -20%
  - Correction: -10%
  - Flash crash: -5%
  - Rally: +10%
  Calculate the impact using my portfolio's beta.

Use formulas where possible. Apply the same professional styling as the other sheets.
Adding Investment Theses
Provide your investment thesis for each position:

Add an Investment_Theses sheet with columns: Ticker, Weight (formula linking to Positions), Thesis, Catalyst, Edge

Here are my theses (adjust column widths so the text is readable, use text wrapping):

AAPL:
- Thesis: iPhone installed base approaching 2 billion creates massive services revenue opportunity. Hardware is mature but services growing 15%+ annually.
- Catalyst: India manufacturing expansion reduces China risk, unlocks next billion users.
- Edge: Market treats it as ex-growth hardware company, missing services compounding.

MSFT:
- Thesis: Cloud + AI leadership. Azure growing faster than AWS. Copilot monetization just starting.
- Catalyst: Enterprise Copilot adoption hitting inflection point in 2026.
- Edge: Multiple expansion as market realizes AI revenue is real, not hype.

[Continue for each of your tickers...]

Set row heights to 50-60 pixels so the text fits. Use zebra striping.
Creating the Update Script
You do not want to rebuild the entire workbook every day. Ask the AI to create an update script:

Create an update_portfolio.py script that I can run daily to update my portfolio. It should:

1. Read the existing workbook (not rebuild it)
2. Fetch new market data from Yahoo Finance for all dates after the last date in Market_Data
3. Append the new prices to Market_Data
4. Check my original Trade_Log file for any new trades and sync them
5. Update Daily_Units and Daily_Cash with the new dates
6. If there is a new ticker I have never traded before, add it to Reference_Data and create new columns in Market_Data and Daily_Units

Include three ways to run it:
- python update_portfolio.py (full sync: trades + prices)
- python update_portfolio.py quick (just fetch latest prices, faster)
- python update_portfolio.py trade 34 2026-01-27 AAPL BUY 100 150.50 (add a single trade manually)

Important: This script should APPEND data, not rebuild. The Excel formulas will automatically recalculate when the new data appears.
Your Daily Workflow (After Setup)
Once everything is set up, your daily routine is simple:

To Update Prices (30 seconds)
Windows Mac Open Command Prompt Open Terminal cd Desktop cd ~/Desktop python update_portfolio.py quick python3 update_portfolio.py quick

Open your spreadsheet - all prices and calculations are updated.

To Log a New Trade
Option A - Add to your original Trade_Log Excel file, then run:

Windows Mac python update_portfolio.py python3 update_portfolio.py

Option B - Command line:

Windows Mac python update_portfolio.py trade 34 2026-01-27 AAPL BUY 100 150.50 python3 update_portfolio.py trade 34 2026-01-27 AAPL BUY 100 150.50

To Add a New Feature
Launch an AI coding agent session

Describe what you want: “Add a sheet that shows my best and worst performers this month”

Watch the AI build it

If it’s not quite right: “Make the percentages show 2 decimal places”

Repeat until perfect

Troubleshooting Common Issues
“python” is not recognized as an internal or external command
Windows
Python was not added to PATH during installation. Uninstall Python, reinstall it, and CHECK THE BOX that says “Add Python to PATH”.

Mac
On Mac, use python3 instead of python:

python3 --version
python3 update_portfolio.py
“pip: command not found” (Mac)
Use pip3 instead:

pip3 install yfinance
“PermissionError: [Errno 13] Permission denied”
The Excel file is open. Close Excel completely before running the script.

“ModuleNotFoundError: No module named ‘yfinance’”
Run this command to install the missing package:

Windows Mac pip install yfinance pip3 install yfinance

Permission errors when installing packages (Mac)
Add --user flag:

pip3 install --user yfinance openpyxl xlsxwriter
Excel file won’t open from Terminal (Mac)
On Mac, use:

open "My Portfolio_BUILT.xlsx"
Or tell the AI to use: open "/Users/YourName/Desktop/My Portfolio_BUILT.xlsx"

“yfinance returns empty data”
Either the ticker symbol is wrong (check Yahoo Finance for the correct symbol) or Yahoo Finance is temporarily unavailable. Try again in a few minutes.

“UnicodeEncodeError or encoding errors”
Run Python with:

python -X utf8 your_script.py
“The AI ran out of context”
Start a new conversation. Tell the AI: “Read my file at [path] and continue where we left off.” The AI can read the files from previous sessions.

“Cells show numbers instead of formulas”
The AI hardcoded values instead of using formulas. Tell the AI: “The cells in column J have hardcoded values. Rebuild this column using formulas like =I2-H2 so they update automatically.”

Homebrew not installed? (Mac)
Install it first (one-time setup):

/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
What Success Looks Like
When complete, your workbook will have sheets like these:

Sheet What It Shows Dashboard Executive summary with big KPIs, NAV chart, allocation donut Positions Live P&L for each stock with data bars showing weights Equity_Curve Daily portfolio value with running drawdown Investment_Theses Your thesis, catalyst, and edge for each position Valuation P/E, EV/EBITDA, P/B with heat map formatting Profitability Margins and ROE with 3-color scales Risk_Trading Beta, short interest, distance from 52-week high Analyst_Estimates Price targets and rating badges Smart_Labels Thematic exposure tags with color coding Correlation Matrix showing how your stocks move together Risk_Metrics VaR, Sharpe, Sortino, Beta, Max Drawdown Stress_Testing What happens in a crash Sector_Exposure Sector breakdown with weight bars Geographic_Exposure Country breakdown Trade_Log All your transactions Market_Data Daily prices (the data lake) Reference_Data Company info (the data lake) Daily_Units Shares held each day Daily_Cash Cash balance each day

Quick Reference: Windows vs Mac Commands
Task Windows Mac Open terminal Search "cmd" Search "Terminal" Navigate to Desktop cd Desktop cd ~/Desktop Run Python python script.py python3 script.py Install packages pip install package pip3 install package Check Python version python --version python3 --version Open Excel file start "file.xlsx" open "file.xlsx" File paths C:\Users\Name\Desktop\file.xlsx /Users/Name/Desktop/file.xlsx

## Generated Output Reference

When you run `build_portfolio_report.py` against `sample_portfolio.xlsx`, it generates **1 single-page HTML report** and **1 JSON file** in the workspace root.

> **Package Manager:** This project uses **uv** for dependency management. All commands should be run via `uv run python ...`. Install uv from <https://docs.astral.sh/uv/>.

### Single-Page HTML Report

The script produces a single `index.html` file (~160 KB) containing all analytics as a **tab-based single-page application (SPA)**. Navigation between sections uses JavaScript tab switching with hash-based URLs (e.g., `#dashboard`, `#positions`, `#options`, `#correlation`, `#risk`, `#stress`, `#exposure`).

| Tab (hash) | Description |
|------------|-------------|
| `#dashboard` | **Dashboard** — Landing page with summary KPIs (portfolio value CAD/USD, annualized return, Sharpe ratio, max drawdown, beta, option delta CAD/USD), card links to all other tabs, and a footer disclaimer stating the content is informational/educational only and not investment advice. |
| `#positions` | **Positions** — All holdings (stocks, ETFs, cash) with **live prices fetched from yfinance** (latest closing price replaces stale spreadsheet values), recalculated market values (CAD and USD), portfolio weights, beta, industry, and option badge counts. **Option-only tickers** (symbols with options but 0 shares) appear with a purple "opts only" badge. **Sortable columns** — click any header to sort ascending/descending. Privacy mode hides share counts and dollar amounts. |
| `#options` | **Options** — All option contracts with live option prices fetched from yfinance option chains, underlying price, and calculated contract value (opt price x shares in CAD). No P/L or notional columns — replaced by Opt Price and Contract Value. Includes a **delta exposure table** showing estimated delta, net delta, and notional delta (CAD) per contract. KPI summary includes total options holding value. Sortable columns. |
| `#correlation` | **Correlation Matrix** — Pairwise return correlation heatmap. **Click any ticker** in the header or row labels to sort the matrix by correlation to that ticker. **Hover any cell** to see a tooltip showing the two tickers and their correlation value. |
| `#risk` | **Risk Metrics** — KPI cards grouped into sections (Portfolio Overview, Risk-Adjusted Returns, Drawdown & Market Risk, Value at Risk, Distribution Shape, Option Hedging). Portfolio value shown in both CAD and USD. VaR ($) shown in both CAD and USD. Net Delta shown in both USD and CAD. **Hover any KPI card** for a plain-English explanation of the financial term. Individual position risk table with beta computed from SPY returns for tickers where yfinance does not provide it (ETFs, crypto). Sortable. |
| `#stress` | **Stress Testing** — 14 scenarios from -50% depression to +50% mania showing unhedged impact, option hedge P&L, hedged impact, and estimated NAV. Summary box shows portfolio value and option delta in both CAD and USD. **USD/CAD toggle button** on the scenario table switches all dollar values between CAD and USD. Displays 1Y portfolio return for context. |
| `#exposure` | **Exposure Analysis** — Three-column layout showing sector, currency, and brokerage account breakdowns with weight bars. All three views include option contract values (opt price × shares) alongside stock holdings — not the inflated notional value. Sector/industry data fetched from yfinance. Sector table shows both CAD and USD columns. |

### JSON Output

| File | Description |
|------|-------------|
| `risk_metrics.json` | Machine-readable risk metrics (Sharpe, Sortino, VaR, Beta, etc.) for programmatic consumption. |

### Input Requirements

The script reads from `sample_portfolio.xlsx` in the workspace root, which must contain:

- **Portfolio** sheet: Columns `Account, Symbol, Shares, Price, Currency, Mkt Value, Mkt Value (CAD)` (Price and Mkt Value columns from the spreadsheet are overridden at runtime with live prices from yfinance; extra columns like historical returns are ignored)
- **Options** sheet: Columns `Symbol, Account, Expirty, Type (CALL/PUT), Strike, Shares, Price, Currency, Cost` (P/L and N.Value columns are ignored; live option prices are fetched from yfinance)
- **Currency** sheet: Row 1 must have the USD/CAD exchange rate in the `Price` column

Fundamentals (sector, industry, beta, type) are fetched live from yfinance for all tradeable tickers. No Fundamental sheet is required.

### Interactive Features

- **Tab Navigation**: All analytics sections are displayed as tabs within a single `index.html` page. Click navigation links to switch between Dashboard, Positions, Options, Correlation, Risk Metrics, Stress Testing, and Exposure tabs. URL hash updates automatically (e.g., `#positions`) for direct linking and browser back/forward support.
- **Language Toggle**: Click the "中" button in the top navigation bar to switch the entire UI to Hong Kong Traditional Chinese (繁體中文). All page titles, section headers, KPI labels, table column headers, metric names, tooltips, stress scenario names, dashboard card titles/descriptions, and navigation links are translated. Click "En" to switch back to English. Language preference persists via localStorage.
- **Investment Disclaimer**: Dashboard page includes a footer disclaimer clarifying the report is for informational and educational purposes only and is not investment advice (language-aware).
- **Privacy Toggle**: Click the "$ Hide" button in the top navigation bar to hide all dollar amounts and share counts across every tab (only percentages and ratios remain visible). State persists via localStorage. Click "$ Show" to reveal amounts again. Button text follows the current language.
- **Sortable Tables**: All position/risk/option tables support click-to-sort on column headers with ascending/descending toggle and sort arrow indicators.
- **Correlation Sort**: Click a ticker in the correlation matrix to sort all rows (or columns) by that ticker's correlation values.
- **Cell Tooltips**: Hover any correlation cell to see "Ticker A vs Ticker B: 0.XX" in a floating tooltip (language-aware).
- **Term Tooltips**: Hover any risk metric KPI card to see a plain-English (or Chinese) explanation of what the metric means (Sharpe Ratio, Sortino Ratio, Kurtosis, Calmar Ratio, VaR, CVaR, etc.).
- **Option Hedging**: Stress testing and risk metrics incorporate option delta exposure to show hedged vs unhedged portfolio impacts.
- **USD/CAD Toggle**: Stress testing scenario table has a toggle button to switch all dollar amounts between CAD and USD display.
- **Live Stock Prices**: Position prices and market values are updated at build time with the latest closing price from yfinance. The spreadsheet Price/Mkt Value columns are used only as fallback when live data is unavailable.
- **Live Option Prices**: Options page fetches current option premiums from yfinance option chains (mid of bid/ask, falls back to intrinsic value when chain unavailable).
- **Computed Beta**: Individual position risk table computes beta from 1Y daily returns vs SPY for tickers where yfinance `.info` does not provide beta (ETFs, crypto ETFs, commodity ETFs).

### Running the Report

This project uses **uv** as its Python package manager. Dependencies are declared in `pyproject.toml` and managed automatically.

```bash
uv run python build_portfolio_report.py
```

Open `index.html` in a browser to view the dashboard. All tabs are contained in a single file — no other HTML files are generated.

