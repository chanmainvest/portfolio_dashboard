"""
Stock Portfolio Analytics Report Generator (v2)
Reads sample_portfolio.xlsx (Portfolio + Options + Currency sheets),
fetches market data, and generates HTML reports:
- Positions (sortable columns)
- Correlation Matrix (sortable, cell tooltips)
- Risk Metrics (term tooltips, includes option hedging)
- Stress Testing (includes option hedging impact)
- Sector & Currency Exposure
- Options (dedicated page with delta exposure)
"""

import json
import sys
from pathlib import Path
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import yfinance as yf
from scipy import stats

# ─── Configuration ──────────────────────────────────────────────────────────────
PORTFOLIO_FILE = Path(__file__).parent / "sample_portfolio.xlsx"
OUTPUT_DIR = Path(__file__).parent
RISK_FREE_RATE = 0.043  # ~4.3% US T-bill rate
TRADING_DAYS = 252
LOOKBACK_DAYS = 365  # 1 year of history for analytics

# ─── Ticker mapping for Canadian tickers on Yahoo Finance ────────────────────
YAHOO_TICKER_MAP = {
    "XIU": "XIU.TO",
    "FNV": "FNV.TO",
    "WPM": "WPM.TO",
    "CCO": "CCO.TO",
    "TRI": "TRI.TO",
}

# ─── Risk metric descriptions for tooltips ──────────────────────────────────
METRIC_TOOLTIPS = {
    "Total Portfolio Value": "The total market value of all portfolio holdings converted to CAD, including stocks, ETFs, options, mutual funds, and cash.",
    "Annualized Return": "The average daily return extrapolated to a full year (252 trading days). Represents the expected yearly return if current performance continues.",
    "Annualized Volatility": "Standard deviation of daily returns scaled to annual. Measures how much the portfolio value fluctuates. Higher = more risk.",
    "Sharpe Ratio": "Risk-adjusted return: (Portfolio Return - Risk-Free Rate) / Volatility. Above 1.0 is good, above 2.0 is very good. Measures excess return per unit of total risk.",
    "Sortino Ratio": "Like Sharpe but only penalizes downside volatility. (Return - Risk-Free Rate) / Downside Deviation. Higher is better. Ignores upside 'risk'.",
    "Maximum Drawdown": "The largest peak-to-trough decline in portfolio value. Measures the worst-case loss from a high point. E.g., -15% means you lost 15% from a peak.",
    "Beta to SPY": "Portfolio sensitivity to S&P 500 (SPY) movements. Beta=1 means the portfolio moves with the market. Beta<1 = less volatile, Beta>1 = more volatile than market.",
    "VaR 95%": "Value at Risk at 95%% confidence: the maximum daily loss expected 95%% of the time. There's a 5%% chance the daily loss exceeds this amount.",
    "VaR 99%": "Value at Risk at 99%% confidence: the maximum daily loss expected 99%% of the time. More conservative than VaR 95%%.",
    "CVaR 95%": "Conditional VaR (Expected Shortfall): the average loss on days when losses exceed VaR 95%%. Measures 'how bad it gets' in the worst 5%% of days.",
    "VaR 95% ($)": "Dollar amount at risk on a given day at 95%% confidence level. The maximum dollar loss expected 19 out of 20 trading days.",
    "VaR 99% ($)": "Dollar amount at risk on a given day at 99%% confidence level. The maximum dollar loss expected 99 out of 100 trading days.",
    "Skewness": "Measures asymmetry of returns. Negative skew = more extreme losses than gains (fat left tail). Positive = more extreme gains. Zero = symmetric.",
    "Kurtosis": "Measures 'fat tails' - how likely extreme events are vs. normal distribution. Higher kurtosis = more frequent extreme moves. Normal distribution = 3.",
    "Calmar Ratio": "Annualized Return / Maximum Drawdown. Measures return per unit of drawdown risk. Higher = better risk-adjusted returns. Above 3.0 is excellent.",
    "Option Delta Exposure": "Net delta exposure from all option positions in CAD. Positive delta = bullish, negative = bearish. Measures the portfolio's effective stock-equivalent exposure from options.",
    "Option Hedging Impact": "The ratio of option delta exposure to portfolio value. Shows how much options modify the portfolio's effective market exposure.",
    "Hedged VaR 95%": "Value at Risk adjusted for option hedging. Option positions (especially protective puts) can reduce downside risk.",
    "Hedged VaR 99%": "Value at Risk at 99%% adjusted for option hedging effects.",
    "Net Delta (USD)": "Total notional delta exposure from options in USD. Represents the stock-equivalent directional bet from all option positions combined.",
}


def read_portfolio(filepath):
    """Read portfolio data from Excel file - stocks, options, cash."""
    print("Reading portfolio from:", filepath)
    wb = pd.ExcelFile(filepath)

    # ── Portfolio sheet (stocks + ETFs + cash) ──
    df = pd.read_excel(wb, sheet_name="Portfolio", header=0)
    core_cols = ["Account", "Symbol", "Shares", "Price", "Currency", "Mkt Value", "Mkt Value (CAD)"]
    df = df[core_cols].copy()
    df = df.dropna(subset=["Symbol"])
    df = df[df["Mkt Value (CAD)"].notna()]

    cash_symbols = {"Cash", "Short Cash"}
    df["PositionType"] = "Stock/ETF"
    df.loc[df["Symbol"].isin(cash_symbols), "PositionType"] = "Cash"
    df["Sector"] = ""

    # ── Read USD/CAD rate from Currency sheet ──
    usd_cad_rate = 1.37
    try:
        cur_df = pd.read_excel(wb, sheet_name="Currency", header=0)
        if "Price" in cur_df.columns and len(cur_df) > 0:
            usd_cad_rate = float(cur_df["Price"].iloc[0])
            print(f"  USD/CAD rate from Currency sheet: {usd_cad_rate}")
    except Exception as e:
        print(f"  Warning: Could not read Currency sheet: {e}")

    # ── Options sheet ──
    opts_df = pd.read_excel(wb, sheet_name="Options", header=0)
    opts_df = opts_df.rename(columns={"Expirty": "Expiry"})
    opt_cols = ["Symbol", "Account", "Expiry", "Type", "Strike", "Shares", "Price",
                "Currency", "P/L", "P/L (CAD)", "Cost", "N.Value", "N.Value (CAD)"]
    available_opt_cols = [c for c in opt_cols if c in opts_df.columns]
    opts_df = opts_df[available_opt_cols].copy()
    opts_df = opts_df.dropna(subset=["Symbol"])
    opts_df["Sector"] = ""

    return df, opts_df, usd_cad_rate


def fetch_fundamentals(tickers):
    """Fetch sector, industry, beta, P/E, and type from Yahoo Finance."""
    print(f"Fetching fundamentals for {len(tickers)} tickers...")
    rows = []
    for symbol in tickers:
        yahoo_sym = get_yahoo_ticker(symbol)
        try:
            info = yf.Ticker(yahoo_sym).info
            rows.append({
                "Symbol": symbol,
                "Type": info.get("quoteType", "Equity"),
                "Beta": info.get("beta"),
                "P/E": info.get("trailingPE"),
                "Industry": info.get("industry", info.get("category", "")),
                "Sector": info.get("sector", info.get("category", "")),
            })
            print(f"    {symbol}: {rows[-1]['Sector']} / {rows[-1]['Industry']} / beta={rows[-1]['Beta']}")
        except Exception as e:
            print(f"    {symbol}: failed ({e})")
            rows.append({"Symbol": symbol, "Type": "", "Beta": None, "P/E": None, "Industry": "", "Sector": ""})
    return pd.DataFrame(rows)


def get_yahoo_ticker(symbol):
    """Map local symbol to Yahoo Finance ticker."""
    return YAHOO_TICKER_MAP.get(symbol, symbol)


def fetch_option_prices(opts_df):
    """Fetch live option premiums from yfinance option chains.

    Returns a list of option mid-prices aligned to opts_df rows.
    Falls back to intrinsic value when the contract can't be found.
    """
    print("Fetching live option prices...")
    prices = []
    cache = {}
    for _, row in opts_df.iterrows():
        symbol = row["Symbol"]
        yahoo_sym = get_yahoo_ticker(symbol)
        strike = row.get("Strike", 0)
        opt_type = row.get("Type", "")
        expiry = row.get("Expiry", None)
        ul_price = row.get("Price", 0) or 0

        if pd.isna(strike) or pd.isna(expiry):
            prices.append(0.0)
            continue

        expiry_str = pd.Timestamp(expiry).strftime("%Y-%m-%d")
        cache_key = (yahoo_sym, expiry_str)

        try:
            if cache_key not in cache:
                tk = yf.Ticker(yahoo_sym)
                chain = tk.option_chain(expiry_str)
                cache[cache_key] = chain
            chain = cache[cache_key]

            if opt_type == "CALL":
                df_chain = chain.calls
            elif opt_type == "PUT":
                df_chain = chain.puts
            else:
                prices.append(0.0)
                continue

            match = df_chain[df_chain["strike"] == strike]
            if not match.empty:
                bid = match["bid"].values[0]
                ask = match["ask"].values[0]
                last = match["lastPrice"].values[0]
                mid = (bid + ask) / 2 if bid > 0 and ask > 0 else last
                prices.append(float(mid) if not pd.isna(mid) else 0.0)
                print(f"    {symbol} {opt_type} {strike} {expiry_str}: ${mid:.2f}")
            else:
                intrinsic = max(0, ul_price - strike) if opt_type == "CALL" else max(0, strike - ul_price)
                prices.append(float(intrinsic))
                print(f"    {symbol} {opt_type} {strike} {expiry_str}: intrinsic ${intrinsic:.2f} (no chain match)")
        except Exception as e:
            intrinsic = max(0, ul_price - strike) if opt_type == "CALL" else max(0, strike - ul_price)
            prices.append(float(intrinsic))
            print(f"    {symbol} {opt_type} {strike} {expiry_str}: intrinsic ${intrinsic:.2f} ({e})")

    return prices


def fetch_price_history(tickers, period_days=LOOKBACK_DAYS):
    """Fetch daily closing prices for all tickers."""
    yahoo_tickers = [get_yahoo_ticker(t) for t in tickers]
    end_date = datetime.now()
    start_date = end_date - timedelta(days=period_days)

    print(f"Fetching price history for {len(yahoo_tickers)} tickers...")
    print(f"  Date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")

    # Download in bulk
    data = yf.download(
        yahoo_tickers,
        start=start_date.strftime("%Y-%m-%d"),
        end=end_date.strftime("%Y-%m-%d"),
        auto_adjust=True,
        progress=False,
    )

    if data.empty:
        print("  WARNING: No data returned from Yahoo Finance!")
        return pd.DataFrame()

    # Extract Close prices
    if isinstance(data.columns, pd.MultiIndex):
        prices = data["Close"]
    else:
        prices = data[["Close"]].copy()
        prices.columns = yahoo_tickers

    # Rename columns back to original symbols
    reverse_map = {v: k for k, v in YAHOO_TICKER_MAP.items()}
    rename_map = {}
    for col in prices.columns:
        if col in reverse_map:
            rename_map[col] = reverse_map[col]
    prices = prices.rename(columns=rename_map)

    # Drop columns that are all NaN
    prices = prices.dropna(axis=1, how="all")

    # Forward fill then back fill
    prices = prices.ffill().bfill()

    print(f"  Retrieved data for {len(prices.columns)} tickers, {len(prices)} trading days")
    return prices


def compute_returns(prices):
    """Compute daily log returns."""
    return np.log(prices / prices.shift(1)).dropna()


def compute_correlation_matrix(returns):
    """Compute correlation matrix of daily returns."""
    return returns.corr()


def compute_option_delta_exposure(opts_df, usd_cad_rate=1.36):
    """Compute net delta exposure from options positions.

    For simplicity, use a delta model:
    - Deep ITM options: delta ~ +/-1.0
    - ATM options: delta ~ +/-0.5
    - Deep OTM options: delta ~ +/-0.0
    Multiply by shares (negative shares = short).
    CALLs have positive delta (long), PUTs have negative delta (long).
    """
    results = []
    total_delta_usd = 0.0

    for _, row in opts_df.iterrows():
        symbol = row["Symbol"]
        opt_type = row.get("Type", "")
        strike = row.get("Strike", 0)
        shares = row.get("Shares", 0)
        price = row.get("Price", 0)
        currency = row.get("Currency", "USD")

        if pd.isna(strike) or pd.isna(shares) or pd.isna(price) or price == 0:
            continue

        # Moneyness ratio
        moneyness = price / strike if strike != 0 else 1.0

        # Simple delta estimation
        if opt_type == "CALL":
            if moneyness > 1.2:
                delta = 0.95  # Deep ITM
            elif moneyness > 1.0:
                delta = 0.5 + 0.45 * (moneyness - 1.0) / 0.2
            elif moneyness > 0.8:
                delta = 0.05 + 0.45 * (moneyness - 0.8) / 0.2
            else:
                delta = 0.05  # Deep OTM
        elif opt_type == "PUT":
            if moneyness < 0.8:
                delta = -0.95  # Deep ITM put
            elif moneyness < 1.0:
                delta = -(0.5 + 0.45 * (1.0 - moneyness) / 0.2)
            elif moneyness < 1.2:
                delta = -(0.05 + 0.45 * (1.2 - moneyness) / 0.2)
            else:
                delta = -0.05  # Deep OTM put
        else:
            delta = 0

        # Net delta = delta * shares (shares already signed: negative = short)
        net_delta = delta * shares
        notional_delta = net_delta * price

        fx_rate = usd_cad_rate if currency == "USD" else 1.0
        notional_delta_cad = notional_delta * fx_rate

        total_delta_usd += notional_delta

        results.append({
            "Symbol": symbol,
            "Type": opt_type,
            "Strike": strike,
            "Shares": shares,
            "Underlying Price": price,
            "Currency": currency,
            "Moneyness": moneyness,
            "Delta": delta,
            "Net Delta": net_delta,
            "Notional Delta (USD)": notional_delta,
            "Notional Delta (CAD)": notional_delta_cad,
        })

    return pd.DataFrame(results), total_delta_usd


def compute_risk_metrics(returns, weights, portfolio_value, option_delta_usd=0, usd_cad_rate=1.36):
    """Compute comprehensive risk metrics including option hedging."""
    metrics = {}

    # Portfolio returns
    portfolio_returns = (returns * weights).sum(axis=1)

    # Annualized return
    mean_daily = portfolio_returns.mean()
    metrics["Annualized Return"] = mean_daily * TRADING_DAYS

    # Annualized Volatility
    daily_vol = portfolio_returns.std()
    metrics["Annualized Volatility"] = daily_vol * np.sqrt(TRADING_DAYS)

    # Sharpe Ratio
    excess_return = metrics["Annualized Return"] - RISK_FREE_RATE
    metrics["Sharpe Ratio"] = excess_return / metrics["Annualized Volatility"] if metrics["Annualized Volatility"] != 0 else 0

    # Sortino Ratio
    downside_returns = portfolio_returns[portfolio_returns < 0]
    downside_deviation = downside_returns.std() * np.sqrt(TRADING_DAYS)
    metrics["Sortino Ratio"] = excess_return / downside_deviation if downside_deviation != 0 else 0

    # Maximum Drawdown
    cumulative = (1 + portfolio_returns).cumprod()
    running_max = cumulative.cummax()
    drawdown = (cumulative - running_max) / running_max
    metrics["Maximum Drawdown"] = drawdown.min()

    # Beta to SPY
    try:
        spy_data = yf.download("SPY", period="1y", auto_adjust=True, progress=False)
        if not spy_data.empty:
            spy_close = spy_data["Close"]
            if isinstance(spy_close, pd.DataFrame):
                spy_close = spy_close.iloc[:, 0]
            spy_returns_series = np.log(spy_close / spy_close.shift(1)).dropna()
            common_dates = portfolio_returns.index.intersection(spy_returns_series.index)
            if len(common_dates) > 10:
                pr = portfolio_returns.loc[common_dates].values.flatten()
                sr = spy_returns_series.loc[common_dates].values.flatten()
                cov_mat = np.cov(pr, sr)
                metrics["Beta to SPY"] = cov_mat[0, 1] / cov_mat[1, 1] if cov_mat[1, 1] != 0 else 0
            else:
                metrics["Beta to SPY"] = "N/A"
        else:
            metrics["Beta to SPY"] = "N/A"
    except Exception as e:
        print(f"  Warning: Could not compute Beta to SPY: {e}")
        metrics["Beta to SPY"] = "N/A"

    # VaR
    metrics["VaR 95%"] = np.percentile(portfolio_returns, 5)
    metrics["VaR 99%"] = np.percentile(portfolio_returns, 1)

    # CVaR
    var_95 = metrics["VaR 95%"]
    tail = portfolio_returns[portfolio_returns <= var_95]
    metrics["CVaR 95%"] = tail.mean() if len(tail) > 0 else var_95

    # Dollar VaR
    metrics["VaR 95% ($)"] = abs(metrics["VaR 95%"]) * portfolio_value
    metrics["VaR 99% ($)"] = abs(metrics["VaR 99%"]) * portfolio_value

    # Skewness and Kurtosis
    metrics["Skewness"] = portfolio_returns.skew()
    metrics["Kurtosis"] = portfolio_returns.kurtosis()

    # Calmar Ratio
    if metrics["Maximum Drawdown"] != 0:
        metrics["Calmar Ratio"] = metrics["Annualized Return"] / abs(metrics["Maximum Drawdown"])
    else:
        metrics["Calmar Ratio"] = 0

    # Option hedging impact
    option_delta_cad = option_delta_usd * usd_cad_rate
    metrics["Net Delta (USD)"] = option_delta_usd
    metrics["Option Delta Exposure"] = option_delta_cad

    # Hedged VaR: option delta offsets a portion of the drawdown
    # Positive delta = long exposure (adds risk), negative delta = hedge (reduces risk)
    hedge_ratio = option_delta_cad / portfolio_value if portfolio_value != 0 else 0
    metrics["Hedged VaR 95%"] = metrics["VaR 95%"] * (1 + hedge_ratio)
    metrics["Hedged VaR 99%"] = metrics["VaR 99%"] * (1 + hedge_ratio)
    metrics["Option Hedging Impact"] = hedge_ratio

    return metrics, portfolio_returns


def compute_stress_testing(portfolio_returns, weights, returns, portfolio_value, beta, option_delta_usd=0, usd_cad_rate=1.36):
    """Compute stress testing scenarios including option hedging."""
    scenarios = {
        "Depression (-50%)": -0.50,
        "Severe Bear (-40%)": -0.40,
        "Bear Market (-30%)": -0.30,
        "Market Crash (-20%)": -0.20,
        "Severe Correction (-15%)": -0.15,
        "Correction (-10%)": -0.10,
        "Flash Crash (-5%)": -0.05,
        "Mild Pullback (-3%)": -0.03,
        "Rally (+5%)": 0.05,
        "Strong Rally (+10%)": 0.10,
        "Bull Run (+20%)": 0.20,
        "Euphoria (+30%)": 0.30,
        "Bubble (+40%)": 0.40,
        "Mania (+50%)": 0.50,
    }

    beta_val = beta if isinstance(beta, (int, float)) else 1.0
    option_delta_cad = option_delta_usd * usd_cad_rate

    results = []
    for scenario_name, market_move in scenarios.items():
        portfolio_impact = market_move * beta_val
        dollar_impact = portfolio_impact * portfolio_value

        # Option hedging effect: delta exposure acts as a modifier
        option_pnl = option_delta_cad * market_move
        hedged_dollar_impact = dollar_impact + option_pnl

        results.append({
            "Scenario": scenario_name,
            "Market Move": market_move,
            "Portfolio Beta": beta_val,
            "Unhedged Impact (%)": portfolio_impact,
            "Unhedged Impact ($)": dollar_impact,
            "Option Hedge P&L ($)": option_pnl,
            "Hedged Impact ($)": hedged_dollar_impact,
            "Hedged Impact (%)": hedged_dollar_impact / portfolio_value if portfolio_value else 0,
            "Estimated NAV": portfolio_value + hedged_dollar_impact,
        })

    return pd.DataFrame(results)


def compute_individual_risk(returns, fund_df, spy_returns=None):
    """Compute per-ticker risk metrics. Computes beta from returns when fund_df has NaN."""
    results = []
    for col in returns.columns:
        r = returns[col].dropna()
        if len(r) < 20:
            continue

        ann_return = r.mean() * TRADING_DAYS
        ann_vol = r.std() * np.sqrt(TRADING_DAYS)
        sharpe = (ann_return - RISK_FREE_RATE) / ann_vol if ann_vol != 0 else 0

        cumulative = (1 + r).cumprod()
        running_max = cumulative.cummax()
        dd = ((cumulative - running_max) / running_max).min()

        var_95 = np.percentile(r, 5)

        fund_row = fund_df[fund_df["Symbol"] == col]
        beta = fund_row["Beta"].values[0] if len(fund_row) > 0 and "Beta" in fund_row.columns else None

        if (beta is None or (isinstance(beta, float) and np.isnan(beta))) and spy_returns is not None:
            common_dates = r.index.intersection(spy_returns.index)
            if len(common_dates) > 20:
                tr = r.loc[common_dates].values.flatten()
                sr = spy_returns.loc[common_dates].values.flatten()
                cov_mat = np.cov(tr, sr)
                beta = float(cov_mat[0, 1] / cov_mat[1, 1]) if cov_mat[1, 1] != 0 else 0.0

        results.append({
            "Ticker": col,
            "Ann. Return": ann_return,
            "Ann. Volatility": ann_vol,
            "Sharpe Ratio": sharpe,
            "Max Drawdown": dd,
            "VaR 95%": var_95,
            "Beta": beta,
        })

    return pd.DataFrame(results)


# ═══════════════════════════════════════════════════════════════════════════════
# HTML GENERATORS
# ═══════════════════════════════════════════════════════════════════════════════

COMMON_CSS = """
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #0B1220; color: #E0E0E0; margin: 20px; }
    h1, h2 { color: #E0E0E0; border-bottom: 2px solid #2A3F5F; padding-bottom: 10px; }
    .info { color: #8899AA; margin-bottom: 20px; font-size: 14px; }
    .nav { background: #1C2541; padding: 10px 16px; border-radius: 8px; margin-bottom: 20px; display: flex; gap: 16px; flex-wrap: wrap; align-items: center; }
    .nav a { color: #7AAFFF; text-decoration: none; font-size: 13px; padding: 4px 10px; border-radius: 4px; }
    .nav a:hover { background: #2A3F5F; }
    .nav a.active { background: #3A7BD5; color: white; }
    .nav .spacer { flex: 1; }
    .nav .privacy-toggle { background: #2A3F5F; color: #D4A843; border: 1px solid #3A7BD5; padding: 4px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: bold; }
    .nav .privacy-toggle:hover { background: #3A7BD5; color: white; }
    .nav .lang-toggle { background: #2A3F5F; color: #7AAFFF; border: 1px solid #3A7BD5; padding: 4px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: bold; margin-left: 8px; }
    .nav .lang-toggle:hover { background: #3A7BD5; color: white; }
    body.privacy-mode .dollar-amount { visibility: hidden; }
    body.privacy-mode .dollar-amount::after { content: '***'; visibility: visible; }
    .positive { color: #007A33; }
    .negative { color: #B81D13; }
    .timestamp { color: #556677; font-size: 12px; margin-top: 20px; }
"""

NAV_LINKS = [
    ("index.html", "Dashboard", "nav_dashboard"),
    ("positions.html", "Positions", "nav_positions"),
    ("options.html", "Options", "nav_options"),
    ("correlation_matrix.html", "Correlation", "nav_correlation"),
    ("risk_metrics.html", "Risk Metrics", "nav_risk_metrics"),
    ("stress_testing.html", "Stress Testing", "nav_stress_testing"),
    ("sector_exposure.html", "Exposure", "nav_exposure"),
]


PRIVACY_JS = """
<script>
(function() {
    var key = 'portfolio_privacy_mode';
    if (localStorage.getItem(key) === 'true') {
        document.body.classList.add('privacy-mode');
    }
    window.togglePrivacy = function() {
        document.body.classList.toggle('privacy-mode');
        localStorage.setItem(key, document.body.classList.contains('privacy-mode'));
        var btn = document.getElementById('privacy-btn');
        if (btn) {
            btn.setAttribute('data-i18n', document.body.classList.contains('privacy-mode') ? 'privacy_show' : 'privacy_hide');
            if (window.applyLanguage) applyLanguage();
        }
    };
    document.addEventListener('DOMContentLoaded', function() {
        var btn = document.getElementById('privacy-btn');
        if (btn && document.body.classList.contains('privacy-mode')) {
            btn.setAttribute('data-i18n', 'privacy_show');
        }
    });
})();
</script>
"""

TRANSLATIONS = {
    "nav_dashboard": {"en": "Dashboard", "zh": "儀表板"},
    "nav_positions": {"en": "Positions", "zh": "持倉"},
    "nav_options": {"en": "Options", "zh": "期權"},
    "nav_correlation": {"en": "Correlation", "zh": "相關性"},
    "nav_risk_metrics": {"en": "Risk Metrics", "zh": "風險指標"},
    "nav_stress_testing": {"en": "Stress Testing", "zh": "壓力測試"},
    "nav_exposure": {"en": "Exposure", "zh": "風險敞口"},
    "privacy_hide": {"en": "$ Hide", "zh": "$ 隱藏"},
    "privacy_show": {"en": "$ Show", "zh": "$ 顯示"},
    "generated": {"en": "Generated:", "zh": "生成時間:"},
    "disclaimer": {"en": "Disclaimer: This dashboard is for informational and educational purposes only and is not investment advice.", "zh": "免責聲明：本儀表板僅供資訊及教育用途，不構成任何投資建議。"},
    "title_dashboard": {"en": "Stock Portfolio Analytics Dashboard", "zh": "投資組合分析儀表板"},
    "desc_dashboard": {"en": "Comprehensive portfolio analysis with risk metrics, correlations, option hedging, and stress testing", "zh": "涵蓋風險指標、相關性、期權對沖及壓力測試的全面投資組合分析"},
    "kpi_pv_cad": {"en": "Portfolio Value (CAD)", "zh": "投資組合價值 (加元)"},
    "kpi_pv_usd": {"en": "Portfolio Value (USD)", "zh": "投資組合價值 (美元)"},
    "kpi_ann_ret": {"en": "Annualized Return", "zh": "年化回報"},
    "kpi_sharpe": {"en": "Sharpe Ratio", "zh": "夏普比率"},
    "kpi_max_dd": {"en": "Max Drawdown", "zh": "最大回撤"},
    "kpi_beta": {"en": "Beta to SPY", "zh": "SPY Beta"},
    "kpi_pos_opt": {"en": "Positions / Options", "zh": "持倉 / 期權"},
    "kpi_delta_cad": {"en": "Option Delta (CAD)", "zh": "期權 Delta (加元)"},
    "kpi_delta_usd": {"en": "Option Delta (USD)", "zh": "期權 Delta (美元)"},
    "card_positions": {"en": "Positions", "zh": "持倉"},
    "card_positions_desc": {"en": "All portfolio positions: stocks, ETFs, mutual funds, cash. Market values, weights, beta, and industry. Sortable columns.", "zh": "所有投資組合持倉：股票、ETF、互惠基金、現金。市值、權重、Beta及行業。可排序欄位。"},
    "card_options": {"en": "Options", "zh": "期權"},
    "card_options_desc": {"en": "All option contracts with delta exposure analysis. Calls, puts, spreads, and their hedging impact on the portfolio.", "zh": "所有期權合約及Delta敞口分析。認購、認沽期權及其對投資組合的對沖影響。"},
    "card_correlation": {"en": "Correlation Matrix", "zh": "相關性矩陣"},
    "card_correlation_desc": {"en": "Pairwise return correlations with heatmap. Click tickers to sort. Hover cells for ticker pair details.", "zh": "配對回報相關性熱圖。點擊股票代號排序。懸停查看配對詳情。"},
    "card_risk": {"en": "Risk Metrics", "zh": "風險指標"},
    "card_risk_desc": {"en": "VaR, Sharpe, Sortino, Calmar, Maximum Drawdown, Beta, option hedging impact. Hover cards for term explanations.", "zh": "VaR、夏普、索提諾、卡瑪比率、最大回撤、Beta、期權對沖影響。懸停查看術語解釋。"},
    "card_stress": {"en": "Stress Testing", "zh": "壓力測試"},
    "card_stress_desc": {"en": "Scenario analysis from -50% crash to +50% rally, showing both unhedged and option-hedged impacts with 1Y return context.", "zh": "從-50%崩盤到+50%上升的情景分析，顯示未對沖及期權對沖後的影響，附一年回報背景。"},
    "card_exposure": {"en": "Sector, Currency & Account Exposure", "zh": "行業、貨幣及帳戶風險敞口"},
    "card_exposure_desc": {"en": "Portfolio breakdown by sector allocation (incl. option notional), currency denomination, and brokerage account.", "zh": "按行業（含期權名義值）、貨幣及券商帳戶的投資組合分佈。"},
    "title_positions": {"en": "Portfolio Positions", "zh": "投資組合持倉"},
    "info_positions": {"en": "All positions including stocks, ETFs, mutual funds, and cash. Click column headers to sort.", "zh": "所有持倉包括股票、ETF、互惠基金及現金。點擊欄位標題排序。"},
    "pos_total": {"en": "Total Positions", "zh": "持倉總數"},
    "pos_pv_cad": {"en": "Portfolio Value (CAD)", "zh": "投資組合價值 (加元)"},
    "pos_stocks": {"en": "Stocks", "zh": "股票"},
    "pos_etfs": {"en": "ETFs", "zh": "ETF"},
    "pos_mf": {"en": "Mutual Funds", "zh": "互惠基金"},
    "pos_opt_contracts": {"en": "Option Contracts", "zh": "期權合約"},
    "pos_cash": {"en": "Cash", "zh": "現金"},
    "th_hash": {"en": "#", "zh": "#"},
    "th_symbol": {"en": "Symbol", "zh": "代號"},
    "th_account": {"en": "Account", "zh": "帳戶"},
    "th_sector": {"en": "Sector", "zh": "行業"},
    "th_type": {"en": "Type", "zh": "類型"},
    "th_shares": {"en": "Shares", "zh": "股數"},
    "th_price": {"en": "Price", "zh": "價格"},
    "th_currency": {"en": "Currency", "zh": "貨幣"},
    "th_mkt_cad": {"en": "Mkt Value (CAD)", "zh": "市值 (加元)"},
    "th_mkt_usd": {"en": "Mkt Value (USD)", "zh": "市值 (美元)"},
    "th_weight": {"en": "Weight", "zh": "權重"},
    "th_weight_bar": {"en": "Weight Bar", "zh": "權重條"},
    "th_beta": {"en": "Beta", "zh": "Beta"},
    "th_industry": {"en": "Industry", "zh": "行業分類"},
    "th_options": {"en": "Options", "zh": "期權"},
    "title_options": {"en": "Options Positions & Delta Exposure", "zh": "期權持倉及Delta敞口"},
    "info_options": {"en": "All option contracts with estimated delta exposure. Negative shares = short position. Click headers to sort.", "zh": "所有期權合約及估算Delta敞口。負數股數 = 沽空倉位。點擊標題排序。"},
    "opt_total": {"en": "Total Contracts", "zh": "合約總數"},
    "opt_calls": {"en": "Calls", "zh": "認購"},
    "opt_puts": {"en": "Puts", "zh": "認沽"},
    "opt_delta_usd": {"en": "Net Delta (USD)", "zh": "淨Delta (美元)"},
    "opt_delta_cad": {"en": "Net Delta (CAD)", "zh": "淨Delta (加元)"},
    "h2_opt_contracts": {"en": "Option Contracts", "zh": "期權合約"},
    "h2_delta_exposure": {"en": "Delta Exposure by Position", "zh": "持倉Delta敞口"},
    "th_expiry": {"en": "Expiry", "zh": "到期日"},
    "th_strike": {"en": "Strike", "zh": "行使價"},
    "th_pl_cad": {"en": "P/L (CAD)", "zh": "損益 (加元)"},
    "th_notional_cad": {"en": "Notional (CAD)", "zh": "名義值 (加元)"},
    "th_ul_price": {"en": "UL Price", "zh": "標的價格"},
    "th_moneyness": {"en": "Moneyness", "zh": "價值狀態"},
    "th_delta": {"en": "Delta", "zh": "Delta"},
    "th_net_delta": {"en": "Net Delta", "zh": "淨Delta"},
    "th_not_delta_cad": {"en": "Notional Delta (CAD)", "zh": "名義Delta (加元)"},
    "title_correlation": {"en": "Portfolio Correlation Matrix", "zh": "投資組合相關性矩陣"},
    "info_correlation": {"en": "Correlation of daily log returns over the past 12 months. Click any ticker header or row label to sort. Hover cells to see ticker pair.", "zh": "過去12個月每日對數回報的相關性。點擊任意股票代號標題或行標籤排序。懸停查看配對。"},
    "th_ticker": {"en": "Ticker", "zh": "代號"},
    "legend_strong_neg": {"en": "\u2264 -0.4 (Strong neg.)", "zh": "\u2264 -0.4 (強負相關)"},
    "legend_low": {"en": "~0 (Low)", "zh": "~0 (低相關)"},
    "legend_moderate": {"en": "~0.4-0.7 (Moderate)", "zh": "~0.4-0.7 (中度相關)"},
    "legend_strong_pos": {"en": "\u2265 0.7 (Strong pos.)", "zh": "\u2265 0.7 (強正相關)"},
    "title_risk": {"en": "Portfolio Risk Metrics", "zh": "投資組合風險指標"},
    "info_risk": {"en": "Risk analytics based on 1-year daily return history. Risk-free rate: 4.3%. Hover KPI cards for explanations.", "zh": "基於一年每日回報歷史的風險分析。無風險利率：4.3%。懸停卡片查看術語解釋。"},
    "sec_overview": {"en": "Portfolio Overview", "zh": "投資組合概覽"},
    "sec_risk_adj": {"en": "Risk-Adjusted Returns", "zh": "風險調整回報"},
    "sec_drawdown": {"en": "Drawdown & Market Risk", "zh": "回撤及市場風險"},
    "sec_var": {"en": "Value at Risk", "zh": "風險價值"},
    "sec_dist": {"en": "Distribution Shape", "zh": "分佈形態"},
    "sec_hedge": {"en": "Option Hedging", "zh": "期權對沖"},
    "m_total_pv": {"en": "Total Portfolio Value", "zh": "投資組合總值"},
    "m_ann_ret": {"en": "Annualized Return", "zh": "年化回報"},
    "m_ann_vol": {"en": "Annualized Volatility", "zh": "年化波動率"},
    "m_sharpe": {"en": "Sharpe Ratio", "zh": "夏普比率"},
    "m_sortino": {"en": "Sortino Ratio", "zh": "索提諾比率"},
    "m_calmar": {"en": "Calmar Ratio", "zh": "卡瑪比率"},
    "m_max_dd": {"en": "Maximum Drawdown", "zh": "最大回撤"},
    "m_beta": {"en": "Beta to SPY", "zh": "SPY Beta"},
    "m_var95": {"en": "VaR 95%", "zh": "風險價值 95%"},
    "m_var99": {"en": "VaR 99%", "zh": "風險價值 99%"},
    "m_cvar95": {"en": "CVaR 95%", "zh": "條件風險價值 95%"},
    "m_var95d": {"en": "VaR 95% ($)", "zh": "風險價值 95% ($)"},
    "m_var99d": {"en": "VaR 99% ($)", "zh": "風險價值 99% ($)"},
    "m_skew": {"en": "Skewness", "zh": "偏度"},
    "m_kurt": {"en": "Kurtosis", "zh": "峰度"},
    "m_net_delta": {"en": "Net Delta (USD)", "zh": "淨Delta (美元)"},
    "m_delta_exp": {"en": "Option Delta Exposure", "zh": "期權Delta敞口"},
    "m_hedge_impact": {"en": "Option Hedging Impact", "zh": "期權對沖影響"},
    "m_hvar95": {"en": "Hedged VaR 95%", "zh": "對沖後VaR 95%"},
    "m_hvar99": {"en": "Hedged VaR 99%", "zh": "對沖後VaR 99%"},
    "h2_ind_risk": {"en": "Individual Position Risk", "zh": "個別持倉風險"},
    "th_ann_ret": {"en": "Ann. Return", "zh": "年化回報"},
    "th_ann_vol": {"en": "Ann. Volatility", "zh": "年化波動率"},
    "th_sharpe": {"en": "Sharpe", "zh": "夏普"},
    "th_max_dd": {"en": "Max Drawdown", "zh": "最大回撤"},
    "th_var95": {"en": "VaR 95%", "zh": "VaR 95%"},
    "th_ticker_col": {"en": "Ticker", "zh": "代號"},
    "tt_total_pv": {"en": "Total market value of all portfolio positions including stocks, ETFs, mutual funds, and cash in CAD.", "zh": "所有投資組合持倉的總市值，包括股票、ETF、互惠基金及現金（加元計）。"},
    "tt_ann_ret": {"en": "The compound annual growth rate of the portfolio over the measurement period. Calculated from daily returns annualized to 252 trading days.", "zh": "投資組合在測量期間的複合年增長率。由每日回報年化至252個交易日計算。"},
    "tt_ann_vol": {"en": "Standard deviation of portfolio returns annualized. Higher values indicate greater price fluctuation and uncertainty.", "zh": "投資組合回報的年化標準差。數值越高表示價格波動及不確定性越大。"},
    "tt_sharpe": {"en": "Excess return per unit of total risk. Values above 1.0 are good, above 2.0 excellent. Calculated as (Return - Risk-Free Rate) / Volatility.", "zh": "每單位總風險的超額回報。數值高於1.0為良好，高於2.0為優秀。計算方式：（回報 - 無風險利率）/ 波動率。"},
    "tt_sortino": {"en": "Similar to Sharpe but only penalizes downside volatility. Higher is better. More appropriate when returns are not symmetrically distributed.", "zh": "與夏普類似，但只懲罰下行波動。數值越高越好。當回報分佈不對稱時更為適用。"},
    "tt_calmar": {"en": "Annualized return divided by maximum drawdown. Measures return earned per unit of peak-to-trough decline risk.", "zh": "年化回報除以最大回撤。衡量每單位峰值到谷底下降風險所獲得的回報。"},
    "tt_max_dd": {"en": "Largest peak-to-trough decline in portfolio value. Represents the worst-case historical loss from any high point.", "zh": "投資組合價值最大峰值到谷底的下降。代表從任何高點的最壞歷史損失。"},
    "tt_beta": {"en": "Portfolio sensitivity to S&P 500 (SPY) movements. Beta of 1.0 means the portfolio moves in line with the market. Below 1.0 = less volatile than market.", "zh": "投資組合對標普500（SPY）走勢的敏感度。Beta為1.0表示組合與市場同步。低於1.0 = 波動低於市場。"},
    "tt_var95": {"en": "Value at Risk at 95% confidence. The maximum expected daily loss that should not be exceeded 95% of the time.", "zh": "95%信心水平的風險價值。95%時間內不應超過的最大預期每日損失。"},
    "tt_var99": {"en": "Value at Risk at 99% confidence. The maximum expected daily loss exceeded only 1% of the time.", "zh": "99%信心水平的風險價值。僅1%時間超過的最大預期每日損失。"},
    "tt_cvar95": {"en": "Conditional VaR (Expected Shortfall) at 95%. The average loss in the worst 5% of scenarios. More conservative than VaR.", "zh": "95%的條件VaR（預期虧損）。最差5%情景的平均損失。比VaR更保守。"},
    "tt_var95d": {"en": "VaR 95% expressed as a dollar amount based on current portfolio value.", "zh": "以當前投資組合價值的金額表示的95% VaR。"},
    "tt_var99d": {"en": "VaR 99% expressed as a dollar amount.", "zh": "以金額表示的99% VaR。"},
    "tt_skew": {"en": "Measures asymmetry of return distribution. Negative skew means more frequent small gains but occasional large losses (fat left tail).", "zh": "衡量回報分佈的不對稱性。負偏度表示較頻繁的小收益但偶爾出現大損失（左尾肥厚）。"},
    "tt_kurt": {"en": "Measures tail fatness of return distribution. Higher kurtosis means more extreme events than a normal distribution (fat tails).", "zh": "衡量回報分佈尾部的肥厚程度。峰度越高表示極端事件多於正態分佈（肥尾）。"},
    "tt_net_delta": {"en": "Total directional dollar exposure from all option positions. Positive = net long exposure, negative = net short/hedged.", "zh": "所有期權持倉的總方向性美元敞口。正數 = 淨多頭敞口，負數 = 淨空頭/對沖。"},
    "tt_delta_exp": {"en": "The estimated dollar-equivalent market exposure from option positions, converted to CAD.", "zh": "期權持倉的估算等值美元市場敞口，已轉換為加元。"},
    "tt_hedge_impact": {"en": "The percentage impact of option delta on portfolio risk. Negative values indicate options are reducing portfolio risk.", "zh": "期權Delta對投資組合風險的百分比影響。負值表示期權正在降低組合風險。"},
    "tt_hvar95": {"en": "VaR 95% after accounting for option hedging. Lower than unhedged VaR indicates effective hedging.", "zh": "計入期權對沖後的95% VaR。低於未對沖VaR表示對沖有效。"},
    "tt_hvar99": {"en": "VaR 99% after accounting for option hedging.", "zh": "計入期權對沖後的99% VaR。"},
    "title_stress": {"en": "Portfolio Stress Testing", "zh": "投資組合壓力測試"},
    "info_stress": {"en": "Simulated impact of market-wide moves on portfolio value using beta, including option hedging effects.", "zh": "使用Beta模擬市場整體波動對投資組合價值的影響，包括期權對沖效果。"},
    "st_pv": {"en": "Portfolio Value", "zh": "投資組合價值"},
    "st_beta": {"en": "Portfolio Beta", "zh": "投資組合Beta"},
    "st_delta_cad": {"en": "Option Delta (CAD)", "zh": "期權Delta (加元)"},
    "st_1y_ret": {"en": "1Y Portfolio Return", "zh": "一年組合回報"},
    "th_scenario": {"en": "Scenario", "zh": "情景"},
    "th_mkt_move": {"en": "Market Move", "zh": "市場變動"},
    "th_unhedged_pct": {"en": "Unhedged Impact (%)", "zh": "未對沖影響 (%)"},
    "th_unhedged_d": {"en": "Unhedged Impact ($)", "zh": "未對沖影響 ($)"},
    "th_opt_pnl": {"en": "Option Hedge P&L ($)", "zh": "期權對沖損益 ($)"},
    "th_hedged_pct": {"en": "Hedged Impact (%)", "zh": "對沖後影響 (%)"},
    "th_hedged_d": {"en": "Hedged Impact ($)", "zh": "對沖後影響 ($)"},
    "th_est_nav": {"en": "Estimated NAV", "zh": "估計資產淨值"},
    "sc_bear": {"en": "Bear Capitulation (-50%)", "zh": "熊市投降 (-50%)"},
    "sc_deep": {"en": "Deep Crash (-40%)", "zh": "深度崩盤 (-40%)"},
    "sc_major_crash": {"en": "Major Crash (-30%)", "zh": "重大暴跌 (-30%)"},
    "sc_crash": {"en": "Market Crash (-20%)", "zh": "市場崩盤 (-20%)"},
    "sc_severe": {"en": "Severe Correction (-15%)", "zh": "嚴重調整 (-15%)"},
    "sc_major_corr": {"en": "Major Correction (-10%)", "zh": "大幅修正 (-10%)"},
    "sc_minor": {"en": "Minor Correction (-5%)", "zh": "小幅修正 (-5%)"},
    "sc_flat": {"en": "Flat (0%)", "zh": "持平 (0%)"},
    "sc_modest": {"en": "Modest Rally (+5%)", "zh": "小幅上升 (+5%)"},
    "sc_moderate": {"en": "Moderate Rally (+10%)", "zh": "溫和上升 (+10%)"},
    "sc_strong": {"en": "Strong Rally (+15%)", "zh": "強勁上升 (+15%)"},
    "sc_bull": {"en": "Bull Run (+20%)", "zh": "牛市行情 (+20%)"},
    "sc_surge": {"en": "Major Surge (+30%)", "zh": "大幅飆升 (+30%)"},
    "sc_euphoric": {"en": "Euphoric Rally (+50%)", "zh": "狂熱上升 (+50%)"},
    "title_exposure": {"en": "Portfolio Exposure Analysis", "zh": "投資組合風險敞口分析"},
    "info_exposure": {"en": "Breakdown by sector (including option notional), currency, and brokerage account.", "zh": "按行業（含期權名義值）、貨幣及券商帳戶的分佈。"},
    "h2_sector": {"en": "Sector Exposure", "zh": "行業風險敞口"},
    "h2_currency": {"en": "Currency Exposure", "zh": "貨幣風險敞口"},
    "h2_account": {"en": "Account Exposure", "zh": "帳戶風險敞口"},
    "th_value_cad": {"en": "Value (CAD)", "zh": "價值 (加元)"},
    "th_value_usd": {"en": "Value (USD)", "zh": "價值 (美元)"},
}

_LANG_JS_CODE = """
    var langKey = 'portfolio_language';
    function getLang() { return localStorage.getItem(langKey) || 'en'; }
    window.applyLanguage = function() {
        var lang = getLang();
        document.querySelectorAll('[data-i18n]').forEach(function(el) {
            var k = el.getAttribute('data-i18n');
            if (T[k] && T[k][lang]) el.textContent = T[k][lang];
        });
        document.querySelectorAll('[data-i18n-html]').forEach(function(el) {
            var k = el.getAttribute('data-i18n-html');
            if (T[k] && T[k][lang]) el.innerHTML = T[k][lang];
        });
        var langBtn = document.getElementById('lang-btn');
        if (langBtn) langBtn.textContent = lang === 'en' ? '中' : 'En';
    };
    window.toggleLanguage = function() {
        var lang = getLang() === 'en' ? 'zh' : 'en';
        localStorage.setItem(langKey, lang);
        applyLanguage();
    };
    document.addEventListener('DOMContentLoaded', applyLanguage);
"""

LANG_JS = (
    "\n<script>\n(function() {\n    var T = "
    + json.dumps(TRANSLATIONS, ensure_ascii=False)
    + ";\n"
    + _LANG_JS_CODE
    + "})();\n</script>\n"
)


def _nav(active_page=""):
    """Return nav HTML with active page highlighted and language toggle."""
    links = []
    for href, label, i18n_key in NAV_LINKS:
        cls = ' class="active"' if href == active_page else ""
        links.append(f'<a href="{href}"{cls} data-i18n="{i18n_key}">{label}</a>')
    toggle_btn = '<span class="spacer"></span>'
    toggle_btn += '<button id="privacy-btn" class="privacy-toggle" data-i18n="privacy_hide" onclick="togglePrivacy()">$ Hide</button>'
    toggle_btn += '<button id="lang-btn" class="lang-toggle" onclick="toggleLanguage()">中</button>'
    return '<div class="nav">' + "".join(links) + toggle_btn + "</div>" + PRIVACY_JS + LANG_JS


SORTABLE_JS = """
<script>
function sortTable(tableId, colIdx, isNumeric) {
    const table = document.getElementById(tableId);
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const header = table.querySelectorAll('thead th')[colIdx];

    // Determine current direction
    const curDir = header.getAttribute('data-sort-dir') || 'none';
    const newDir = curDir === 'asc' ? 'desc' : 'asc';

    // Clear all headers
    table.querySelectorAll('thead th').forEach(th => {
        th.setAttribute('data-sort-dir', 'none');
        const arrow = th.querySelector('.sort-arrow');
        if (arrow) arrow.remove();
    });

    header.setAttribute('data-sort-dir', newDir);
    const arrow = document.createElement('span');
    arrow.className = 'sort-arrow';
    arrow.style.marginLeft = '4px';
    arrow.style.fontSize = '10px';
    arrow.textContent = newDir === 'asc' ? '\\u25B2' : '\\u25BC';
    header.appendChild(arrow);

    rows.sort((a, b) => {
        let aVal = a.cells[colIdx].getAttribute('data-val') || a.cells[colIdx].textContent.trim();
        let bVal = b.cells[colIdx].getAttribute('data-val') || b.cells[colIdx].textContent.trim();
        if (isNumeric) {
            aVal = parseFloat(aVal.replace(/[$,%+]/g, '')) || 0;
            bVal = parseFloat(bVal.replace(/[$,%+]/g, '')) || 0;
        } else {
            aVal = aVal.toLowerCase();
            bVal = bVal.toLowerCase();
        }
        if (aVal < bVal) return newDir === 'asc' ? -1 : 1;
        if (aVal > bVal) return newDir === 'asc' ? 1 : -1;
        return 0;
    });

    rows.forEach(r => tbody.appendChild(r));
}
</script>
"""


def _corr_color(val, is_diag):
    """Return (background, text) color for correlation value."""
    if is_diag:
        return "#2A3F5F", "#AABBCC"
    if val >= 0.7:
        return "rgb(200,50,50)", "white"
    elif val >= 0.4:
        r = int(200 + (val - 0.4) / 0.3 * 55)
        g = int(180 - (val - 0.4) / 0.3 * 130)
        return f"rgb({min(r, 255)},{max(g, 0)},50)", "white"
    elif val >= 0.0:
        r = int(50 + val / 0.4 * 150)
        g = int(150 + val / 0.4 * 30)
        return f"rgb({r},{g},50)", "white"
    elif val >= -0.4:
        g = int(150 + abs(val) / 0.4 * 50)
        b = int(100 + abs(val) / 0.4 * 100)
        return f"rgb(50,{g},{b})", "white"
    else:
        g = int(100 + abs(val + 0.4) / 0.6 * 100)
        return f"rgb(30,{g},200)", "white"


def generate_html_correlation(corr_matrix):
    """Generate HTML for correlation matrix with sort-by-click and cell tooltips."""
    tickers = list(corr_matrix.columns)

    # Serialize correlation data as JSON for JS sort
    corr_json = {}
    for t1 in tickers:
        corr_json[t1] = {}
        for t2 in tickers:
            v = corr_matrix.loc[t1, t2]
            corr_json[t1][t2] = round(float(v), 4) if not np.isnan(v) else None

    # Build header row
    header_cells = ""
    for t in tickers:
        header_cells += f'<th onclick="sortByCol(\'{t}\')">{t}</th>'

    # Build body rows
    body_rows = ""
    for t1 in tickers:
        cells = f'<td class="row-header" onclick="sortByRow(\'{t1}\')">{t1}</td>'
        for t2 in tickers:
            val = corr_matrix.loc[t1, t2]
            if np.isnan(val):
                cells += f'<td data-r="{t1}" data-c="{t2}" style="background:#1A2744;">-</td>'
            else:
                bg, fg = _corr_color(val, t1 == t2)
                cells += f'<td data-r="{t1}" data-c="{t2}" style="background:{bg};color:{fg};">{val:.2f}</td>'
        body_rows += f"<tr>{cells}</tr>\n"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Portfolio Correlation Matrix</title>
<style>
    {COMMON_CSS}
    .table-container {{ overflow-x: auto; position: relative; }}
    table {{ border-collapse: collapse; font-size: 11px; }}
    th {{ background: #1C2541; color: white; padding: 6px 8px; position: sticky; top: 0; z-index: 2; white-space: nowrap; cursor: pointer; user-select: none; }}
    th:hover {{ background: #2A3F5F; }}
    th.row-header {{ position: sticky; left: 0; z-index: 3; background: #1C2541; cursor: pointer; }}
    td {{ padding: 5px 7px; text-align: center; border: 1px solid #1A2744; font-size: 11px; white-space: nowrap; }}
    td.row-header {{ position: sticky; left: 0; background: #1C2541; color: white; font-weight: bold; z-index: 1; text-align: left; cursor: pointer; }}
    td.row-header:hover {{ background: #2A3F5F; }}
    .legend {{ margin-top: 20px; display: flex; gap: 20px; align-items: center; font-size: 13px; flex-wrap: wrap; }}
    .legend-item {{ display: flex; align-items: center; gap: 6px; }}
    .legend-box {{ width: 20px; height: 20px; border-radius: 3px; }}
    #cell-tooltip {{ position: fixed; background: #1C2541; color: #E0E0E0; border: 1px solid #3A7BD5; padding: 8px 12px; border-radius: 6px; font-size: 12px; pointer-events: none; z-index: 100; display: none; box-shadow: 0 4px 12px rgba(0,0,0,0.5); }}
</style>
</head>
<body>
{_nav("correlation_matrix.html")}
<h1 data-i18n="title_correlation">Portfolio Correlation Matrix</h1>
<p class="info" data-i18n="info_correlation">Correlation of daily log returns over the past 12 months. Click any ticker header or row label to sort. Hover cells to see ticker pair.</p>
<div id="cell-tooltip"></div>
<div class="table-container">
<table id="corr-table">
<thead><tr><th class="row-header" onclick="resetSort()" data-i18n="th_ticker">Ticker</th>
{header_cells}
</tr></thead>
<tbody>
{body_rows}
</tbody>
</table>
</div>
<div class="legend">
    <div class="legend-item"><div class="legend-box" style="background:rgb(30,200,200);"></div> <span data-i18n="legend_strong_neg">&le; -0.4 (Strong neg.)</span></div>
    <div class="legend-item"><div class="legend-box" style="background:rgb(50,150,100);"></div> <span data-i18n="legend_low">~0 (Low)</span></div>
    <div class="legend-item"><div class="legend-box" style="background:rgb(200,180,50);"></div> <span data-i18n="legend_moderate">~0.4-0.7 (Moderate)</span></div>
    <div class="legend-item"><div class="legend-box" style="background:rgb(200,50,50);"></div> <span data-i18n="legend_strong_pos">&ge; 0.7 (Strong pos.)</span></div>
</div>
<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>

<script>
const CORR = {json.dumps(corr_json)};
const TICKERS = {json.dumps(tickers)};
let currentSort = null;
let sortAsc = true;

// Tooltip
const tooltip = document.getElementById('cell-tooltip');
function attachTooltips() {{
    document.querySelectorAll('#corr-table td[data-r]').forEach(td => {{
        td.addEventListener('mouseenter', function(e) {{
            const r = this.getAttribute('data-r');
            const c = this.getAttribute('data-c');
            const v = this.textContent;
            const isZh = (localStorage.getItem('portfolio_language') || 'en') === 'zh';
            tooltip.innerHTML = '<strong>' + r + '</strong> ' + (isZh ? '對' : 'vs') + ' <strong>' + c + '</strong><br>' + (isZh ? '相關性: ' : 'Correlation: ') + v;
            tooltip.style.display = 'block';
        }});
        td.addEventListener('mousemove', function(e) {{
            tooltip.style.left = (e.clientX + 14) + 'px';
            tooltip.style.top = (e.clientY + 14) + 'px';
        }});
        td.addEventListener('mouseleave', function() {{ tooltip.style.display = 'none'; }});
    }});
}}
attachTooltips();

function getCorrColor(val, isDiag) {{
    if (isDiag) return ['#2A3F5F', '#AABBCC'];
    if (val >= 0.7) return ['rgb(200,50,50)', 'white'];
    if (val >= 0.4) {{
        var r = Math.min(255, Math.round(200 + (val - 0.4) / 0.3 * 55));
        var g = Math.max(0, Math.round(180 - (val - 0.4) / 0.3 * 130));
        return ['rgb(' + r + ',' + g + ',50)', 'white'];
    }}
    if (val >= 0) {{
        var r = Math.round(50 + val / 0.4 * 150);
        var g = Math.round(150 + val / 0.4 * 30);
        return ['rgb(' + r + ',' + g + ',50)', 'white'];
    }}
    if (val >= -0.4) {{
        var g = Math.round(150 + Math.abs(val) / 0.4 * 50);
        var b = Math.round(100 + Math.abs(val) / 0.4 * 100);
        return ['rgb(50,' + g + ',' + b + ')', 'white'];
    }}
    var g = Math.round(100 + Math.abs(val + 0.4) / 0.6 * 100);
    return ['rgb(30,' + g + ',200)', 'white'];
}}

function rebuildRow(row, sorted) {{
    var rowTicker = row.querySelector('td.row-header').textContent;
    while (row.children.length > 1) row.removeChild(row.lastChild);
    sorted.forEach(function(colTicker) {{
        var td = document.createElement('td');
        td.setAttribute('data-r', rowTicker);
        td.setAttribute('data-c', colTicker);
        var v = CORR[rowTicker] && CORR[rowTicker][colTicker] != null ? CORR[rowTicker][colTicker] : null;
        if (v === null) {{
            td.style.background = '#1A2744';
            td.textContent = '-';
        }} else {{
            td.textContent = v.toFixed(2);
            var colors = getCorrColor(v, rowTicker === colTicker);
            td.style.background = colors[0];
            td.style.color = colors[1];
        }}
        row.appendChild(td);
    }});
}}

function sortByCol(ticker) {{
    if (currentSort === 'col_' + ticker) {{ sortAsc = !sortAsc; }} else {{ currentSort = 'col_' + ticker; sortAsc = false; }}
    var tbody = document.querySelector('#corr-table tbody');
    var rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort(function(a, b) {{
        var aLabel = a.querySelector('td.row-header').textContent;
        var bLabel = b.querySelector('td.row-header').textContent;
        var aVal = CORR[aLabel] && CORR[aLabel][ticker] != null ? CORR[aLabel][ticker] : -999;
        var bVal = CORR[bLabel] && CORR[bLabel][ticker] != null ? CORR[bLabel][ticker] : -999;
        return sortAsc ? aVal - bVal : bVal - aVal;
    }});
    rows.forEach(function(r) {{ tbody.appendChild(r); }});
    attachTooltips();
}}

function sortByRow(ticker) {{
    // Sort columns by correlation to this ticker
    if (currentSort === 'row_' + ticker) {{ sortAsc = !sortAsc; }} else {{ currentSort = 'row_' + ticker; sortAsc = false; }}
    var sorted = TICKERS.slice().sort(function(a, b) {{
        var aVal = CORR[ticker] && CORR[ticker][a] != null ? CORR[ticker][a] : -999;
        var bVal = CORR[ticker] && CORR[ticker][b] != null ? CORR[ticker][b] : -999;
        return sortAsc ? aVal - bVal : bVal - aVal;
    }});
    // Rebuild header
    var thead = document.querySelector('#corr-table thead tr');
    while (thead.children.length > 1) thead.removeChild(thead.lastChild);
    sorted.forEach(function(t) {{
        var th = document.createElement('th');
        th.textContent = t;
        th.onclick = function() {{ sortByCol(t); }};
        thead.appendChild(th);
    }});
    // Rebuild each row
    var rows = document.querySelectorAll('#corr-table tbody tr');
    rows.forEach(function(row) {{ rebuildRow(row, sorted); }});
    attachTooltips();
}}

function resetSort() {{
    location.reload();
}}
</script>
</body></html>"""
    return html


def generate_html_risk_metrics(metrics, individual_risk_df, portfolio_value, usd_cad_rate=1.37):
    """Generate HTML for risk metrics with term description tooltips."""
    portfolio_value_usd = portfolio_value / usd_cad_rate
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Portfolio Risk Metrics</title>
<style>
    {COMMON_CSS}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; margin-bottom: 30px; }}
    .kpi-card {{ background: #1C2541; border-radius: 8px; padding: 18px; border-left: 4px solid #3A7BD5; position: relative; cursor: help; }}
    .kpi-card.pos {{ border-left-color: #007A33; }}
    .kpi-card.neg {{ border-left-color: #B81D13; }}
    .kpi-card.neut {{ border-left-color: #D4A843; }}
    .kpi-label {{ color: #8899AA; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }}
    .kpi-value {{ font-size: 24px; font-weight: bold; }}
    .kpi-value.pos {{ color: #007A33; }}
    .kpi-value.neg {{ color: #B81D13; }}
    .kpi-value.neut {{ color: #D4A843; }}
    .kpi-tooltip {{ display: none; position: absolute; bottom: calc(100% + 8px); left: 0; right: 0; background: #0D1526; border: 1px solid #3A7BD5; border-radius: 6px; padding: 10px 14px; font-size: 12px; color: #CCDDEE; line-height: 1.5; z-index: 20; box-shadow: 0 4px 16px rgba(0,0,0,0.5); pointer-events: none; }}
    .kpi-card:hover .kpi-tooltip {{ display: block; }}
    .help-icon {{ font-size: 11px; color: #556677; margin-left: 4px; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; }}
    th {{ background: #1C2541; color: white; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; cursor: pointer; user-select: none; }}
    th:hover {{ background: #2A3F5F; }}
    td {{ padding: 8px 12px; border-bottom: 1px solid #1A2744; font-size: 13px; }}
    tr:nth-child(even) {{ background: #111B2E; }}
    tr:nth-child(odd) {{ background: #0D1526; }}
    tr:hover {{ background: #1A2744; }}
    h2 {{ margin-top: 30px; }}
    .section-label {{ background: #2A3F5F; color: #AAC0DD; padding: 8px 14px; border-radius: 6px; font-size: 13px; margin: 24px 0 12px 0; display: inline-block; }}
</style>
</head>
<body>
{_nav("risk_metrics.html")}
<h1 data-i18n="title_risk">Portfolio Risk Metrics</h1>
<p class="info" data-i18n="info_risk">Risk analytics based on 1-year daily return history. Risk-free rate: {RISK_FREE_RATE:.1%}. Hover KPI cards for explanations.</p>
"""

    def _fmt_metric(key):
        """Format metric value for display."""
        v = metrics.get(key, 0)
        if isinstance(v, str):
            return v
        if key in ("Annualized Return", "Annualized Volatility", "Maximum Drawdown",
                    "VaR 95%", "VaR 99%", "CVaR 95%", "Hedged VaR 95%", "Hedged VaR 99%",
                    "Option Hedging Impact"):
            return f"{v:.2%}" if not pd.isna(v) else "N/A"
        if key in ("Sharpe Ratio", "Sortino Ratio", "Calmar Ratio", "Beta to SPY", "Skewness", "Kurtosis"):
            return f"{v:.3f}" if isinstance(v, (int, float)) and not pd.isna(v) else str(v)
        if key in dollar_keys:
            return f"${v:,.0f}" if isinstance(v, (int, float)) else str(v)
        return str(v)

    def _kpi_cls(key):
        """Determine KPI card styling class."""
        v = metrics.get(key, 0)
        if isinstance(v, str):
            return "neut"
        if key in ("Annualized Return", "Sharpe Ratio", "Sortino Ratio", "Calmar Ratio"):
            return "pos" if v > 0 else "neg"
        if key in ("Maximum Drawdown", "VaR 95%", "VaR 99%", "CVaR 95%",
                    "VaR 95% (CAD)", "VaR 95% (USD)", "VaR 99% (CAD)", "VaR 99% (USD)",
                    "Hedged VaR 95%", "Hedged VaR 99%"):
            return "neg"
        return "neut"

    # Group KPIs
    groups = [
        ("Portfolio Overview", ["Total Portfolio Value (CAD)", "Total Portfolio Value (USD)", "Annualized Return", "Annualized Volatility"]),
        ("Risk-Adjusted Returns", ["Sharpe Ratio", "Sortino Ratio", "Calmar Ratio"]),
        ("Drawdown & Market Risk", ["Maximum Drawdown", "Beta to SPY"]),
        ("Value at Risk", ["VaR 95%", "VaR 99%", "CVaR 95%", "VaR 95% (CAD)", "VaR 95% (USD)", "VaR 99% (CAD)", "VaR 99% (USD)"]),
        ("Distribution Shape", ["Skewness", "Kurtosis"]),
        ("Option Hedging", ["Net Delta (USD)", "Net Delta (CAD)", "Option Hedging Impact", "Hedged VaR 95%", "Hedged VaR 99%"]),
    ]

    # Set derived metrics
    metrics["Total Portfolio Value (CAD)"] = portfolio_value
    metrics["Total Portfolio Value (USD)"] = portfolio_value_usd
    metrics["VaR 95% (CAD)"] = abs(metrics.get("VaR 95%", 0)) * portfolio_value
    metrics["VaR 95% (USD)"] = abs(metrics.get("VaR 95%", 0)) * portfolio_value_usd
    metrics["VaR 99% (CAD)"] = abs(metrics.get("VaR 99%", 0)) * portfolio_value
    metrics["VaR 99% (USD)"] = abs(metrics.get("VaR 99%", 0)) * portfolio_value_usd
    net_delta_usd = metrics.get("Net Delta (USD)", 0)
    metrics["Net Delta (CAD)"] = net_delta_usd * usd_cad_rate if isinstance(net_delta_usd, (int, float)) else 0

    # Keys that display dollar amounts
    dollar_keys = {
        "Total Portfolio Value (CAD)", "Total Portfolio Value (USD)",
        "VaR 95% (CAD)", "VaR 95% (USD)", "VaR 99% (CAD)", "VaR 99% (USD)",
        "Net Delta (USD)", "Net Delta (CAD)",
    }

    # i18n key mappings for risk metrics
    _sec_i18n = {"Portfolio Overview": "sec_overview", "Risk-Adjusted Returns": "sec_risk_adj",
                 "Drawdown & Market Risk": "sec_drawdown", "Value at Risk": "sec_var",
                 "Distribution Shape": "sec_dist", "Option Hedging": "sec_hedge"}
    _m_i18n = {"Total Portfolio Value (CAD)": "m_total_pv_cad", "Total Portfolio Value (USD)": "m_total_pv_usd",
               "Annualized Return": "m_ann_ret",
               "Annualized Volatility": "m_ann_vol", "Sharpe Ratio": "m_sharpe",
               "Sortino Ratio": "m_sortino", "Calmar Ratio": "m_calmar",
               "Maximum Drawdown": "m_max_dd", "Beta to SPY": "m_beta",
               "VaR 95%": "m_var95", "VaR 99%": "m_var99", "CVaR 95%": "m_cvar95",
               "VaR 95% (CAD)": "m_var95_cad", "VaR 95% (USD)": "m_var95_usd",
               "VaR 99% (CAD)": "m_var99_cad", "VaR 99% (USD)": "m_var99_usd",
               "Skewness": "m_skew", "Kurtosis": "m_kurt",
               "Net Delta (USD)": "m_net_delta_usd", "Net Delta (CAD)": "m_net_delta_cad",
               "Option Hedging Impact": "m_hedge_impact", "Hedged VaR 95%": "m_hvar95",
               "Hedged VaR 99%": "m_hvar99"}
    _tt_i18n = {"Total Portfolio Value (CAD)": "tt_total_pv", "Total Portfolio Value (USD)": "tt_total_pv",
                "Annualized Return": "tt_ann_ret",
                "Annualized Volatility": "tt_ann_vol", "Sharpe Ratio": "tt_sharpe",
                "Sortino Ratio": "tt_sortino", "Calmar Ratio": "tt_calmar",
                "Maximum Drawdown": "tt_max_dd", "Beta to SPY": "tt_beta",
                "VaR 95%": "tt_var95", "VaR 99%": "tt_var99", "CVaR 95%": "tt_cvar95",
                "VaR 95% (CAD)": "tt_var95d", "VaR 95% (USD)": "tt_var95d",
                "VaR 99% (CAD)": "tt_var99d", "VaR 99% (USD)": "tt_var99d",
                "Skewness": "tt_skew", "Kurtosis": "tt_kurt",
                "Net Delta (USD)": "tt_net_delta", "Net Delta (CAD)": "tt_net_delta",
                "Option Hedging Impact": "tt_hedge_impact", "Hedged VaR 95%": "tt_hvar95",
                "Hedged VaR 99%": "tt_hvar99"}

    for group_name, keys in groups:
        sec_key = _sec_i18n.get(group_name, "")
        html += f'<div class="section-label" data-i18n="{sec_key}">{group_name}</div>\n<div class="kpi-grid">\n'
        for key in keys:
            cls = _kpi_cls(key)
            value = _fmt_metric(key)
            tooltip = METRIC_TOOLTIPS.get(key, "")
            tt_key = _tt_i18n.get(key, "")
            tooltip_div = f'<div class="kpi-tooltip" data-i18n-html="{tt_key}">{tooltip}</div>' if tooltip else ""
            dollar_cls = " dollar-amount" if key in dollar_keys else ""
            m_key = _m_i18n.get(key, "")
            html += f"""    <div class="kpi-card {cls}">
        {tooltip_div}
        <div class="kpi-label"><span data-i18n="{m_key}">{key}</span> <span class="help-icon">&#9432;</span></div>
        <div class="kpi-value {cls}{dollar_cls}">{value}</div>
    </div>\n"""
        html += '</div>\n'

    # Individual ticker risk table
    html += '<h2 data-i18n="h2_ind_risk">Individual Position Risk</h2>\n'
    html += '<table id="ind-risk-table">\n<thead><tr>'
    ind_cols = [("Ticker", False, "th_ticker_col"), ("Ann. Return", True, "th_ann_ret"), ("Ann. Volatility", True, "th_ann_vol"),
                ("Sharpe", True, "th_sharpe"), ("Max Drawdown", True, "th_max_dd"), ("VaR 95%", True, "th_var95"), ("Beta", True, "th_beta")]
    for idx, (col, is_num, i18n_key) in enumerate(ind_cols):
        html += f'<th onclick="sortTable(\'ind-risk-table\',{idx},{str(is_num).lower()})" data-i18n="{i18n_key}">{col}</th>'
    html += '</tr></thead>\n<tbody>\n'

    if not individual_risk_df.empty:
        individual_risk_df_sorted = individual_risk_df.sort_values("Ann. Return", ascending=False)
        for _, row in individual_risk_df_sorted.iterrows():
            ret_cls = "positive" if row["Ann. Return"] > 0 else "negative"
            beta_val = f'{row["Beta"]:.2f}' if isinstance(row["Beta"], (int, float)) and not pd.isna(row["Beta"]) else str(row["Beta"])
            html += f"""<tr>
    <td><strong>{row['Ticker']}</strong></td>
    <td class="{ret_cls}" data-val="{row['Ann. Return']:.6f}">{row['Ann. Return']:.2%}</td>
    <td data-val="{row['Ann. Volatility']:.6f}">{row['Ann. Volatility']:.2%}</td>
    <td data-val="{row['Sharpe Ratio']:.6f}">{row['Sharpe Ratio']:.3f}</td>
    <td class="negative" data-val="{row['Max Drawdown']:.6f}">{row['Max Drawdown']:.2%}</td>
    <td class="negative" data-val="{row['VaR 95%']:.6f}">{row['VaR 95%']:.2%}</td>
    <td data-val="{row['Beta'] if isinstance(row['Beta'], (int, float)) else 0}">{beta_val}</td>
</tr>\n"""

    html += '</tbody></table>\n'
    html += SORTABLE_JS
    html += f'<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>'
    html += "</body></html>"
    return html


def generate_html_stress_testing(stress_df, portfolio_value, beta, option_delta_usd=0, usd_cad_rate=1.37, ann_return=0):
    """Generate HTML for stress testing including option hedging."""
    option_delta_cad = option_delta_usd * usd_cad_rate
    portfolio_value_usd = portfolio_value / usd_cad_rate
    option_delta_usd_val = option_delta_usd

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Portfolio Stress Testing</title>
<style>
    {COMMON_CSS}
    .summary-box {{ background: #1C2541; padding: 16px 24px; border-radius: 8px; margin-bottom: 24px; display: flex; gap: 30px; flex-wrap: wrap; font-size: 14px; }}
    .summary-box .item {{ }}
    .summary-box .label {{ color: #8899AA; font-size: 11px; text-transform: uppercase; }}
    .summary-box .value {{ color: #D4A843; font-weight: bold; font-size: 18px; }}
    .currency-toggle {{ background: #2A3F5F; color: #7AAFFF; border: 1px solid #3A7BD5; padding: 6px 14px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: bold; margin-bottom: 12px; }}
    .currency-toggle:hover {{ background: #3A7BD5; color: white; }}
    .currency-toggle.active {{ background: #3A7BD5; color: white; }}
    table {{ border-collapse: collapse; width: 100%; max-width: 1400px; }}
    th {{ background: #1C2541; color: white; padding: 12px 16px; text-align: left; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }}
    td {{ padding: 12px 16px; border-bottom: 1px solid #1A2744; font-size: 14px; }}
    tr:nth-child(even) {{ background: #111B2E; }}
    tr:nth-child(odd) {{ background: #0D1526; }}
    tr:hover {{ background: #1A2744; }}
</style>
</head>
<body>
{_nav("stress_testing.html")}
<h1 data-i18n="title_stress">Portfolio Stress Testing</h1>
<p class="info" data-i18n="info_stress">Simulated impact of market-wide moves on portfolio value using beta, including option hedging effects.</p>
<div class="summary-box">
    <div class="item"><div class="label">Portfolio Value (CAD)</div><div class="value dollar-amount">${portfolio_value:,.0f}</div></div>
    <div class="item"><div class="label">Portfolio Value (USD)</div><div class="value dollar-amount">${portfolio_value_usd:,.0f}</div></div>
    <div class="item"><div class="label" data-i18n="st_beta">Portfolio Beta</div><div class="value">{beta:.3f}</div></div>
    <div class="item"><div class="label">Option Delta (CAD)</div><div class="value dollar-amount">${option_delta_cad:+,.0f}</div></div>
    <div class="item"><div class="label">Option Delta (USD)</div><div class="value dollar-amount">${option_delta_usd_val:+,.0f}</div></div>
    <div class="item"><div class="label" data-i18n="st_1y_ret">1Y Portfolio Return</div><div class="value {'positive' if ann_return > 0 else 'negative'}">{ann_return:.2%}</div></div>
</div>

<button class="currency-toggle active" id="ccy-toggle" onclick="toggleCurrency()">Showing: CAD &mdash; click for USD</button>

<table id="stress-table">
<thead>
<tr>
    <th data-i18n="th_scenario">Scenario</th>
    <th data-i18n="th_mkt_move">Market Move</th>
    <th data-i18n="th_unhedged_pct">Unhedged Impact (%)</th>
    <th>Unhedged Impact ($)</th>
    <th>Option Hedge P&amp;L ($)</th>
    <th data-i18n="th_hedged_pct">Hedged Impact (%)</th>
    <th>Hedged Impact ($)</th>
    <th>Estimated NAV</th>
</tr>
</thead>
<tbody>
"""

    _sc_i18n = {
        "Bear Capitulation (-50%)": "sc_bear", "Deep Crash (-40%)": "sc_deep",
        "Major Crash (-30%)": "sc_major_crash", "Market Crash (-20%)": "sc_crash",
        "Severe Correction (-15%)": "sc_severe", "Major Correction (-10%)": "sc_major_corr",
        "Minor Correction (-5%)": "sc_minor", "Flat (0%)": "sc_flat",
        "Modest Rally (+5%)": "sc_modest", "Moderate Rally (+10%)": "sc_moderate",
        "Strong Rally (+15%)": "sc_strong", "Bull Run (+20%)": "sc_bull",
        "Major Surge (+30%)": "sc_surge", "Euphoric Rally (+50%)": "sc_euphoric",
    }

    for _, row in stress_df.iterrows():
        cls = "positive" if row["Hedged Impact (%)"] > 0 else "negative"
        cls_un = "positive" if row["Unhedged Impact (%)"] > 0 else "negative"
        opt_cls = "positive" if row["Option Hedge P&L ($)"] > 0 else "negative" if row["Option Hedge P&L ($)"] < 0 else ""
        sc_key = _sc_i18n.get(row['Scenario'], '')
        sc_attr = f' data-i18n="{sc_key}"' if sc_key else ''

        unhedged_cad = row['Unhedged Impact ($)']
        opt_pnl_cad = row['Option Hedge P&L ($)']
        hedged_cad = row['Hedged Impact ($)']
        nav_cad = row['Estimated NAV']

        html += f"""<tr>
    <td><strong{sc_attr}>{row['Scenario']}</strong></td>
    <td class="{cls_un}">{row['Market Move']:.0%}</td>
    <td class="{cls_un}">{row['Unhedged Impact (%)']:.2%}</td>
    <td class="{cls_un} dollar-amount ccy-cell" data-cad="{unhedged_cad:.0f}" data-usd="{unhedged_cad / usd_cad_rate:.0f}">${unhedged_cad:+,.0f}</td>
    <td class="{opt_cls} dollar-amount ccy-cell" data-cad="{opt_pnl_cad:.0f}" data-usd="{opt_pnl_cad / usd_cad_rate:.0f}">${opt_pnl_cad:+,.0f}</td>
    <td class="{cls}">{row['Hedged Impact (%)']:.2%}</td>
    <td class="{cls} dollar-amount ccy-cell" data-cad="{hedged_cad:.0f}" data-usd="{hedged_cad / usd_cad_rate:.0f}">${hedged_cad:+,.0f}</td>
    <td class="dollar-amount ccy-cell" data-cad="{nav_cad:.0f}" data-usd="{nav_cad / usd_cad_rate:.0f}">${nav_cad:,.0f}</td>
</tr>\n"""

    html += f"""</tbody>
</table>
<script>
(function() {{
    var showCAD = true;
    window.toggleCurrency = function() {{
        showCAD = !showCAD;
        var btn = document.getElementById('ccy-toggle');
        btn.textContent = showCAD ? 'Showing: CAD \\u2014 click for USD' : 'Showing: USD \\u2014 click for CAD';
        document.querySelectorAll('.ccy-cell').forEach(function(td) {{
            var raw = showCAD ? parseFloat(td.getAttribute('data-cad')) : parseFloat(td.getAttribute('data-usd'));
            if (isNaN(raw)) return;
            var sign = raw >= 0 ? '' : '-';
            var abs = Math.abs(raw);
            var formatted = abs.toLocaleString('en-US', {{maximumFractionDigits: 0}});
            var hasPlus = td.getAttribute('data-cad') && parseFloat(td.getAttribute('data-cad')) !== parseFloat(td.textContent.replace(/[$,+]/g, ''));
            if (td.textContent.indexOf('+') === 1 || td.textContent.indexOf('-') === 1) {{
                td.textContent = '$' + (raw >= 0 ? '+' : '') + (raw < 0 ? '-' : '') + formatted;
            }} else {{
                td.textContent = '$' + sign + formatted;
            }}
        }});
    }};
}})();
</script>
<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
</body></html>"""
    return html


def generate_html_positions(portfolio_df, opts_df, fund_df, portfolio_value, usd_cad_rate=1.37):
    """Generate positions analytics HTML with sortable table."""
    # Merge with fundamentals
    fund_cols = ["Symbol", "Type", "Beta", "Industry"]
    if "P/E" in fund_df.columns:
        fund_cols.append("P/E")
    available_cols = [c for c in fund_cols if c in fund_df.columns]
    merged = portfolio_df.merge(fund_df[available_cols], on="Symbol", how="left")

    merged.loc[merged["PositionType"] == "Cash", "Type"] = "Cash"
    merged.loc[merged["PositionType"] == "Cash", "Beta"] = 0.0

    merged["Weight"] = merged["Mkt Value (CAD)"] / portfolio_value if portfolio_value else 0

    # Count options by underlying
    option_counts = opts_df.groupby("Symbol").size().to_dict()

    # Add option-only tickers (tickers with options but 0 shares in portfolio)
    portfolio_symbols = set(merged["Symbol"].unique())
    option_only_symbols = [s for s in opts_df["Symbol"].unique() if s not in portfolio_symbols]
    if option_only_symbols:
        opt_only_rows = []
        for sym in option_only_symbols:
            opt_rows = opts_df[opts_df["Symbol"] == sym]
            sector = opt_rows["Sector"].iloc[0] if "Sector" in opt_rows.columns and not opt_rows["Sector"].isna().all() else "-"
            currency = opt_rows["Currency"].iloc[0] if "Currency" in opt_rows.columns and not opt_rows["Currency"].isna().all() else "USD"
            # Look up beta and industry from fund_df
            fund_row = fund_df[fund_df["Symbol"] == sym]
            beta = fund_row["Beta"].values[0] if len(fund_row) > 0 and "Beta" in fund_row.columns and pd.notna(fund_row["Beta"].values[0]) else None
            industry = fund_row["Industry"].values[0] if len(fund_row) > 0 and "Industry" in fund_row.columns and pd.notna(fund_row["Industry"].values[0]) else "-"
            ftype = fund_row["Type"].values[0] if len(fund_row) > 0 and "Type" in fund_row.columns and pd.notna(fund_row["Type"].values[0]) else "-"

            opt_only_rows.append({
                "Symbol": sym,
                "Shares": 0,
                "Price": 0,
                "Currency": currency,
                "Mkt Value": 0,
                "Mkt Value (CAD)": 0,
                "Sector": sector,
                "Account": "Options Only",
                "PositionType": "Options Only",
                "Type": ftype,
                "Beta": beta,
                "Industry": industry,
                "Weight": 0,
            })
        opt_only_df = pd.DataFrame(opt_only_rows)
        merged = pd.concat([merged, opt_only_df], ignore_index=True)

    # Summary KPIs
    num_positions = len(merged)
    total_value = merged["Mkt Value (CAD)"].sum()
    num_stocks = len(merged[merged.get("Type", pd.Series(dtype=str)) == "Stock"]) if "Type" in merged.columns else 0
    num_etfs = len(merged[merged.get("Type", pd.Series(dtype=str)) == "ETF"]) if "Type" in merged.columns else 0
    num_mfs = len(merged[merged["PositionType"] == "Mutual Fund"])
    num_options = len(opts_df)
    cash_total = merged[merged["PositionType"] == "Cash"]["Mkt Value (CAD)"].sum()

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Portfolio Positions</title>
<style>
    {COMMON_CSS}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px; margin-bottom: 24px; }}
    .kpi-card {{ background: #1C2541; border-radius: 8px; padding: 14px; border-left: 4px solid #3A7BD5; }}
    .kpi-label {{ color: #8899AA; font-size: 11px; text-transform: uppercase; }}
    .kpi-value {{ font-size: 20px; font-weight: bold; color: #D4A843; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th {{ background: #1C2541; color: white; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; position: sticky; top: 0; cursor: pointer; user-select: none; }}
    th:hover {{ background: #2A3F5F; }}
    td {{ padding: 8px 12px; border-bottom: 1px solid #1A2744; font-size: 13px; }}
    tr:nth-child(even) {{ background: #111B2E; }}
    tr:nth-child(odd) {{ background: #0D1526; }}
    tr:hover {{ background: #1A2744; }}
    .weight-bar {{ height: 8px; background: #3A7BD5; border-radius: 4px; min-width: 2px; }}
    .opt-badge {{ background: #D4A843; color: #0B1220; font-size: 10px; padding: 1px 5px; border-radius: 3px; margin-left: 4px; font-weight: bold; }}
    .opt-only-badge {{ background: #7A5BD5; color: white; font-size: 10px; padding: 1px 5px; border-radius: 3px; margin-left: 4px; font-weight: bold; }}
</style>
</head>
<body>
{_nav("positions.html")}
<h1 data-i18n="title_positions">Portfolio Positions</h1>
<p class="info" data-i18n="info_positions">All positions including stocks, ETFs, mutual funds, and cash. Click column headers to sort.</p>
<div class="kpi-grid">
"""

    _pos_i18n = {"Total Positions": "pos_total", "Portfolio Value (CAD)": "pos_pv_cad",
                  "Stocks": "pos_stocks", "ETFs": "pos_etfs", "Mutual Funds": "pos_mf",
                  "Option Contracts": "pos_opt_contracts", "Cash": "pos_cash"}

    for label, value in [
        ("Total Positions", str(num_positions)),
        ("Portfolio Value (CAD)", f"${total_value:,.0f}"),
        ("Stocks", str(num_stocks)),
        ("ETFs", str(num_etfs)),
        ("Mutual Funds", str(num_mfs)),
        ("Option Contracts", str(num_options)),
        ("Cash", f"${cash_total:,.0f}"),
    ]:
        dollar_cls = " dollar-amount" if value.startswith("$") else ""
        i18n_key = _pos_i18n.get(label, "")
        html += f'<div class="kpi-card"><div class="kpi-label" data-i18n="{i18n_key}">{label}</div><div class="kpi-value{dollar_cls}">{value}</div></div>\n'
    html += '</div>\n'

    merged["Mkt Value (USD)"] = merged.apply(
        lambda r: r["Mkt Value"] if r.get("Currency") == "USD"
        else r["Mkt Value (CAD)"] / usd_cad_rate, axis=1
    )
    merged["Mkt Value (USD)"] = merged["Mkt Value (USD)"].fillna(0)

    # Column definitions: (header, is_numeric, i18n_key)
    columns = [
        ("#", True, "th_hash"), ("Symbol", False, "th_symbol"), ("Account", False, "th_account"),
        ("Sector", False, "th_sector"), ("Type", False, "th_type"),
        ("Shares", True, "th_shares"), ("Price", True, "th_price"), ("Currency", False, "th_currency"),
        ("Mkt Value (CAD)", True, "th_mkt_cad"), ("Mkt Value (USD)", True, "th_mkt_usd"),
        ("Weight", True, "th_weight"), ("Weight Bar", False, "th_weight_bar"),
        ("Beta", True, "th_beta"), ("Industry", False, "th_industry"), ("Options", True, "th_options"),
    ]

    html += '<table id="positions-table"><thead><tr>'
    for idx, (col_name, is_num, i18n_key) in enumerate(columns):
        if col_name == "Weight Bar":
            html += f'<th data-i18n="{i18n_key}">{col_name}</th>'
        else:
            html += f'<th onclick="sortTable(\'positions-table\',{idx},{str(is_num).lower()})" data-i18n="{i18n_key}">{col_name}</th>'
    html += '</tr></thead><tbody>\n'

    merged_sorted = merged.sort_values("Mkt Value (CAD)", ascending=False)
    max_weight = merged_sorted["Weight"].max() if not merged_sorted.empty else 1

    for idx, (_, row) in enumerate(merged_sorted.iterrows(), 1):
        weight_pct = row["Weight"] * 100
        bar_width = (row["Weight"] / max_weight * 100) if max_weight > 0 else 0
        beta_val = f'{row["Beta"]:.2f}' if "Beta" in row and pd.notna(row.get("Beta")) else "-"
        industry = row.get("Industry", "-")
        if pd.isna(industry) or str(industry) == "#VALUE!":
            industry = "-"
        type_val = row.get("Type", "-")
        if pd.isna(type_val):
            type_val = "-"
        sym = row["Symbol"]
        opt_count = option_counts.get(sym, 0)
        opt_badge = f'<span class="opt-badge">{opt_count} opts</span>' if opt_count > 0 else ""
        is_opt_only = row.get("PositionType") == "Options Only"
        if is_opt_only:
            opt_badge = f'<span class="opt-only-badge">opts only</span>{opt_badge}'
        account = row.get("Account", "-")
        if pd.isna(account):
            account = "-"

        mkt_usd = row.get("Mkt Value (USD)", 0)
        if pd.isna(mkt_usd):
            mkt_usd = 0

        html += f"""<tr>
    <td data-val="{idx}">{idx}</td>
    <td><strong>{sym}</strong>{opt_badge}</td>
    <td>{account}</td>
    <td>{row.get('Sector', '-')}</td>
    <td>{type_val}</td>
    <td data-val="{row['Shares']}" class="dollar-amount">{row['Shares']:,.0f}</td>
    <td data-val="{row['Price']}">{row['Price']:,.2f}</td>
    <td>{row['Currency']}</td>
    <td data-val="{row['Mkt Value (CAD)']}" class="dollar-amount">${row['Mkt Value (CAD)']:,.0f}</td>
    <td data-val="{mkt_usd}" class="dollar-amount">${mkt_usd:,.0f}</td>
    <td data-val="{weight_pct:.4f}">{weight_pct:.2f}%</td>
    <td><div class="weight-bar" style="width:{bar_width:.0f}%"></div></td>
    <td data-val="{row.get('Beta', 0) if pd.notna(row.get('Beta')) else 0}">{beta_val}</td>
    <td>{industry}</td>
    <td data-val="{opt_count}">{opt_count if opt_count > 0 else '-'}</td>
</tr>\n"""

    html += '</tbody></table>\n'
    html += SORTABLE_JS
    html += f'<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>'
    html += "</body></html>"
    return html


def generate_html_options(opts_df, option_delta_df, total_delta_usd, usd_cad_rate=1.37):
    """Generate dedicated options page with live prices and contract values."""
    total_delta_cad = total_delta_usd * usd_cad_rate

    total_contracts = len(opts_df)
    calls = len(opts_df[opts_df["Type"] == "CALL"]) if "Type" in opts_df.columns else 0
    puts = len(opts_df[opts_df["Type"] == "PUT"]) if "Type" in opts_df.columns else 0

    # Compute total options holding value (CAD)
    total_opt_value_cad = 0.0
    if "Opt Price" in opts_df.columns:
        for _, row in opts_df.iterrows():
            opt_price = row.get("Opt Price", 0) or 0
            shares = row.get("Shares", 0) or 0
            currency = row.get("Currency", "USD")
            fx = usd_cad_rate if currency == "USD" else 1.0
            total_opt_value_cad += opt_price * shares * fx

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Options Positions</title>
<style>
    {COMMON_CSS}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 14px; margin-bottom: 24px; }}
    .kpi-card {{ background: #1C2541; border-radius: 8px; padding: 14px; border-left: 4px solid #D4A843; }}
    .kpi-label {{ color: #8899AA; font-size: 11px; text-transform: uppercase; }}
    .kpi-value {{ font-size: 20px; font-weight: bold; color: #D4A843; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th {{ background: #1C2541; color: white; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; position: sticky; top: 0; cursor: pointer; user-select: none; }}
    th:hover {{ background: #2A3F5F; }}
    td {{ padding: 8px 12px; border-bottom: 1px solid #1A2744; font-size: 13px; }}
    tr:nth-child(even) {{ background: #111B2E; }}
    tr:nth-child(odd) {{ background: #0D1526; }}
    tr:hover {{ background: #1A2744; }}
    .call {{ color: #00C49A; }}
    .put {{ color: #FF6B6B; }}
</style>
</head>
<body>
{_nav("options.html")}
<h1 data-i18n="title_options">Options Positions &amp; Delta Exposure</h1>
<p class="info" data-i18n="info_options">All option contracts with live prices and estimated delta exposure. Negative shares = short position. Click headers to sort.</p>
<div class="kpi-grid">
"""

    _opt_i18n = {
        "Total Contracts": "opt_total", "Calls": "opt_calls", "Puts": "opt_puts",
        "Options Value (CAD)": "opt_value_cad",
        "Net Delta (USD)": "opt_delta_usd", "Net Delta (CAD)": "opt_delta_cad",
    }

    for label, value in [
        ("Total Contracts", str(total_contracts)),
        ("Calls", str(calls)),
        ("Puts", str(puts)),
        ("Options Value (CAD)", f"${total_opt_value_cad:,.0f}"),
        ("Net Delta (USD)", f"${total_delta_usd:+,.0f}"),
        ("Net Delta (CAD)", f"${total_delta_cad:+,.0f}"),
    ]:
        dollar_cls = " dollar-amount" if "$" in value else ""
        i18n_key = _opt_i18n.get(label, "")
        html += f'<div class="kpi-card"><div class="kpi-label" data-i18n="{i18n_key}">{label}</div><div class="kpi-value{dollar_cls}">{value}</div></div>\n'
    html += '</div>\n'

    # Options contracts table
    html += '<h2 data-i18n="h2_opt_contracts">Option Contracts</h2>\n'
    cols = [("#", False, "th_hash"), ("Symbol", False, "th_symbol"), ("Type", False, "th_type"),
            ("Expiry", False, "th_expiry"), ("Strike", True, "th_strike"),
            ("Shares", True, "th_shares"), ("UL Price", True, "th_ul_price"),
            ("Opt Price", True, "th_opt_price"), ("Currency", False, "th_currency"),
            ("Contract Value", True, "th_contract_value")]

    html += '<table id="options-table"><thead><tr>'
    for idx, (c, is_num, i18n_key) in enumerate(cols):
        html += f'<th onclick="sortTable(\'options-table\',{idx},{str(is_num).lower()})" data-i18n="{i18n_key}">{c}</th>'
    html += '</tr></thead><tbody>\n'

    for idx, (_, row) in enumerate(opts_df.iterrows(), 1):
        type_cls = "call" if row.get("Type") == "CALL" else "put"
        expiry = row.get("Expiry", "")
        if isinstance(expiry, (datetime, pd.Timestamp)):
            expiry = expiry.strftime("%Y-%m-%d")
        shares = row.get("Shares", 0) or 0
        opt_price = row.get("Opt Price", 0) or 0
        ul_price = row.get("Price", 0) or 0
        currency = row.get("Currency", "USD")
        fx = usd_cad_rate if currency == "USD" else 1.0
        contract_value_cad = opt_price * shares * fx
        val_cls = "positive" if contract_value_cad > 0 else "negative" if contract_value_cad < 0 else ""

        html += f"""<tr>
    <td>{idx}</td>
    <td><strong>{row['Symbol']}</strong></td>
    <td class="{type_cls}">{row.get('Type', '')}</td>
    <td>{expiry}</td>
    <td data-val="{row.get('Strike', 0)}" class="dollar-amount">{row.get('Strike', 0):,.1f}</td>
    <td data-val="{shares}">{shares:,.0f}</td>
    <td data-val="{ul_price}" class="dollar-amount">{ul_price:,.2f}</td>
    <td data-val="{opt_price}" class="dollar-amount">{opt_price:,.2f}</td>
    <td>{currency}</td>
    <td data-val="{contract_value_cad}" class="{val_cls} dollar-amount">${contract_value_cad:+,.0f}</td>
</tr>\n"""

    html += '</tbody></table>\n'

    # Delta exposure table
    if not option_delta_df.empty:
        html += '<h2 data-i18n="h2_delta_exposure">Delta Exposure by Position</h2>\n'
        html += '<table id="delta-table"><thead><tr>'
        delta_cols = [("#", False, "th_hash"), ("Symbol", False, "th_symbol"), ("Type", False, "th_type"),
                      ("Strike", True, "th_strike"), ("Shares", True, "th_shares"),
                      ("UL Price", True, "th_ul_price"), ("Moneyness", True, "th_moneyness"),
                      ("Delta", True, "th_delta"), ("Net Delta", True, "th_net_delta"),
                      ("Notional Delta (CAD)", True, "th_not_delta_cad")]
        for idx, (c, is_num, i18n_key) in enumerate(delta_cols):
            html += f'<th onclick="sortTable(\'delta-table\',{idx},{str(is_num).lower()})" data-i18n="{i18n_key}">{c}</th>'
        html += '</tr></thead><tbody>\n'

        for idx, (_, row) in enumerate(option_delta_df.iterrows(), 1):
            type_cls = "call" if row["Type"] == "CALL" else "put"
            delta_cls = "positive" if row["Notional Delta (CAD)"] > 0 else "negative"
            html += f"""<tr>
    <td>{idx}</td>
    <td><strong>{row['Symbol']}</strong></td>
    <td class="{type_cls}">{row['Type']}</td>
    <td data-val="{row['Strike']}" class="dollar-amount">{row['Strike']:,.1f}</td>
    <td data-val="{row['Shares']}">{row['Shares']:,.0f}</td>
    <td data-val="{row['Underlying Price']}" class="dollar-amount">{row['Underlying Price']:,.2f}</td>
    <td data-val="{row['Moneyness']}">{row['Moneyness']:.2f}</td>
    <td data-val="{row['Delta']}">{row['Delta']:.3f}</td>
    <td data-val="{row['Net Delta']}" class="dollar-amount">{row['Net Delta']:,.0f}</td>
    <td data-val="{row['Notional Delta (CAD)']}" class="{delta_cls} dollar-amount">${row['Notional Delta (CAD)']:+,.0f}</td>
</tr>\n"""

        html += '</tbody></table>\n'

    html += SORTABLE_JS
    html += f'<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>'
    html += "</body></html>"
    return html


def generate_html_sector_exposure(portfolio_df, opts_df, portfolio_value, usd_cad_rate=1.37):
    """Generate sector and currency exposure HTML."""
    # Use actual option contract value (opt price x shares) instead of notional
    opts_sector = opts_df.copy()
    if "Opt Price" in opts_sector.columns:
        opts_sector["Mkt Value (CAD)"] = opts_sector.apply(
            lambda r: (r.get("Opt Price", 0) or 0) * (r.get("Shares", 0) or 0)
            * (usd_cad_rate if r.get("Currency") == "USD" else 1.0),
            axis=1,
        )
    else:
        opts_sector["Mkt Value (CAD)"] = 0

    portfolio_with_usd = portfolio_df.copy()
    portfolio_with_usd["Mkt Value (USD)"] = portfolio_with_usd.apply(
        lambda r: r["Mkt Value"] if r.get("Currency") == "USD"
        else r["Mkt Value (CAD)"] / usd_cad_rate, axis=1
    )
    portfolio_with_usd["Mkt Value (USD)"] = portfolio_with_usd["Mkt Value (USD)"].fillna(0)

    opts_sector["Mkt Value (USD)"] = opts_sector["Mkt Value (CAD)"] / usd_cad_rate
    opts_sector["Mkt Value (USD)"] = opts_sector["Mkt Value (USD)"].fillna(0)

    # Combine stock positions + option contract values
    port_cols = ["Symbol", "Sector", "Mkt Value (CAD)", "Mkt Value (USD)", "Currency"]
    opt_cols = ["Symbol", "Sector", "Mkt Value (CAD)", "Mkt Value (USD)"]
    if "Currency" in opts_sector.columns:
        opt_cols.append("Currency")

    all_positions = pd.concat([
        portfolio_with_usd[port_cols],
        opts_sector[opt_cols],
    ], ignore_index=True)
    all_positions["Mkt Value (CAD)"] = all_positions["Mkt Value (CAD)"].fillna(0)
    all_positions["Mkt Value (USD)"] = all_positions["Mkt Value (USD)"].fillna(0)

    # Sector aggregation
    sector_data = all_positions.groupby("Sector").agg(
        total_value=("Mkt Value (CAD)", "sum"),
        total_value_usd=("Mkt Value (USD)", "sum"),
        num_positions=("Symbol", "count"),
    ).reset_index()
    total_sector = sector_data["total_value"].sum()
    sector_data["Weight"] = sector_data["total_value"] / total_sector if total_sector else 0
    sector_data = sector_data.sort_values("total_value", ascending=False)

    # Currency aggregation (portfolio + options)
    all_with_currency = all_positions[all_positions["Currency"].notna()].copy()
    currency_data = all_with_currency.groupby("Currency").agg(
        total_value=("Mkt Value (CAD)", "sum"),
        num_positions=("Symbol", "count"),
    ).reset_index()
    total_cur = currency_data["total_value"].sum()
    currency_data["Weight"] = currency_data["total_value"] / total_cur if total_cur else 0
    currency_data = currency_data.sort_values("total_value", ascending=False)

    # Account aggregation (portfolio + options)
    opts_with_account = opts_sector[["Symbol", "Account", "Mkt Value (CAD)"]].copy() if "Account" in opts_sector.columns else pd.DataFrame()
    port_with_account = portfolio_df[["Symbol", "Account", "Mkt Value (CAD)"]].copy()
    all_with_account = pd.concat([port_with_account, opts_with_account], ignore_index=True)
    all_with_account["Mkt Value (CAD)"] = all_with_account["Mkt Value (CAD)"].fillna(0)
    account_data = all_with_account.groupby("Account").agg(
        total_value=("Mkt Value (CAD)", "sum"),
        num_positions=("Symbol", "count"),
    ).reset_index()
    total_acct = account_data["total_value"].sum()
    account_data["Weight"] = account_data["total_value"] / total_acct if total_acct else 0
    account_data = account_data.sort_values("total_value", ascending=False)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Portfolio Exposure Analysis</title>
<style>
    {COMMON_CSS}
    .grid-3col {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 30px; }}
    @media (max-width: 1200px) {{ .grid-3col {{ grid-template-columns: 1fr 1fr; }} }}
    @media (max-width: 800px) {{ .grid-3col {{ grid-template-columns: 1fr; }} }}
    table {{ border-collapse: collapse; width: 100%; }}
    th {{ background: #1C2541; color: white; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; }}
    td {{ padding: 8px 12px; border-bottom: 1px solid #1A2744; font-size: 13px; }}
    tr:nth-child(even) {{ background: #111B2E; }}
    tr:nth-child(odd) {{ background: #0D1526; }}
    tr:hover {{ background: #1A2744; }}
    .weight-bar {{ height: 10px; border-radius: 5px; min-width: 2px; }}
    .sector-colors {{ background: linear-gradient(90deg, #3A7BD5, #00C49A); }}
    .currency-colors {{ background: linear-gradient(90deg, #D4A843, #E07A3A); }}
    .account-colors {{ background: linear-gradient(90deg, #7A5BD5, #BD5BA8); }}
</style>
</head>
<body>
{_nav("sector_exposure.html")}
<h1 data-i18n="title_exposure">Portfolio Exposure Analysis</h1>
<p class="info" data-i18n="info_exposure">Breakdown by sector (including option notional), currency, and brokerage account.</p>
<div class="grid-3col">
<div>
<h2 data-i18n="h2_sector">Sector Exposure</h2>
<table>
<thead><tr><th data-i18n="th_sector">Sector</th><th data-i18n="th_value_cad">Value (CAD)</th><th data-i18n="th_value_usd">Value (USD)</th><th data-i18n="th_weight">Weight</th><th>#</th><th></th></tr></thead>
<tbody>
"""

    max_sw = sector_data["Weight"].max() if not sector_data.empty else 1
    for _, row in sector_data.iterrows():
        bar_w = (row["Weight"] / max_sw * 100) if max_sw > 0 else 0
        html += f"""<tr>
    <td><strong>{row['Sector']}</strong></td>
    <td class="dollar-amount">${row['total_value']:,.0f}</td>
    <td class="dollar-amount">${row['total_value_usd']:,.0f}</td>
    <td>{row['Weight']:.1%}</td>
    <td>{row['num_positions']}</td>
    <td><div class="weight-bar sector-colors" style="width:{bar_w:.0f}%"></div></td>
</tr>\n"""

    html += """</tbody></table>
</div>
<div>
<h2 data-i18n="h2_currency">Currency Exposure</h2>
<table>
<thead><tr><th data-i18n="th_currency">Currency</th><th data-i18n="th_value_cad">Value (CAD)</th><th data-i18n="th_weight">Weight</th><th>#</th><th></th></tr></thead>
<tbody>
"""

    max_cw = currency_data["Weight"].max() if not currency_data.empty else 1
    for _, row in currency_data.iterrows():
        bar_w = (row["Weight"] / max_cw * 100) if max_cw > 0 else 0
        html += f"""<tr>
    <td><strong>{row['Currency']}</strong></td>
    <td class="dollar-amount">${row['total_value']:,.0f}</td>
    <td>{row['Weight']:.1%}</td>
    <td>{row['num_positions']}</td>
    <td><div class="weight-bar currency-colors" style="width:{bar_w:.0f}%"></div></td>
</tr>\n"""

    html += """</tbody></table>
</div>
<div>
<h2 data-i18n="h2_account">Account Exposure</h2>
<table>
<thead><tr><th data-i18n="th_account">Account</th><th data-i18n="th_value_cad">Value (CAD)</th><th data-i18n="th_weight">Weight</th><th>#</th><th></th></tr></thead>
<tbody>
"""

    max_aw = account_data["Weight"].max() if not account_data.empty else 1
    for _, row in account_data.iterrows():
        bar_w = (row["Weight"] / max_aw * 100) if max_aw > 0 else 0
        html += f"""<tr>
    <td><strong>{row['Account']}</strong></td>
    <td class="dollar-amount">${row['total_value']:,.0f}</td>
    <td>{row['Weight']:.1%}</td>
    <td>{row['num_positions']}</td>
    <td><div class="weight-bar account-colors" style="width:{bar_w:.0f}%"></div></td>
</tr>\n"""

    html += f"""</tbody></table>
</div></div>
<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
</body></html>"""
    return html


def generate_index_html(portfolio_value, metrics, num_positions, num_options, usd_cad_rate=1.37):
    """Generate main dashboard page."""
    sharpe = metrics.get("Sharpe Ratio", 0)
    ann_ret = metrics.get("Annualized Return", 0)
    max_dd = metrics.get("Maximum Drawdown", 0)
    beta = metrics.get("Beta to SPY", "N/A")
    beta_str = f"{beta:.2f}" if isinstance(beta, float) else str(beta)
    delta_cad = metrics.get("Option Delta Exposure", 0)
    delta_usd = metrics.get("Net Delta (USD)", 0)
    portfolio_value_usd = portfolio_value / usd_cad_rate

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Stock Portfolio Analytics Dashboard</title>
<style>
    {COMMON_CSS}
    .header {{ background: linear-gradient(135deg, #1C2541 0%, #2A3F5F 100%); padding: 30px; border-radius: 12px; margin-bottom: 24px; }}
    .header h1 {{ margin: 0 0 8px 0; font-size: 28px; border: none; padding: 0; }}
    .header p {{ margin: 0; color: #8899AA; font-size: 14px; }}
    .kpi-strip {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 28px; }}
    .kpi-mini {{ background: #1C2541; border-radius: 8px; padding: 16px 20px; flex: 1; min-width: 180px; }}
    .kpi-mini .label {{ color: #8899AA; font-size: 11px; text-transform: uppercase; }}
    .kpi-mini .value {{ font-size: 22px; font-weight: bold; color: #D4A843; margin-top: 4px; }}
    .cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px; }}
    .card {{ background: #1C2541; border-radius: 10px; padding: 24px; transition: transform 0.2s, box-shadow 0.2s; cursor: pointer; text-decoration: none; color: inherit; display: block; border: 1px solid #2A3F5F; }}
    .card:hover {{ transform: translateY(-4px); box-shadow: 0 8px 25px rgba(0,0,0,0.3); border-color: #3A7BD5; }}
    .card h2 {{ margin: 0 0 10px 0; font-size: 18px; color: #E0E0E0; border: none; padding: 0; }}
    .card p {{ margin: 0; color: #8899AA; font-size: 13px; line-height: 1.5; }}
    .card .icon {{ font-size: 36px; margin-bottom: 12px; }}
    .disclaimer {{ margin-top: 18px; color: #8899AA; font-size: 12px; }}
</style>
</head>
<body>
{_nav("index.html")}
<div class="header">
    <h1 data-i18n="title_dashboard">Stock Portfolio Analytics Dashboard</h1>
    <p data-i18n="desc_dashboard">Comprehensive portfolio analysis with risk metrics, correlations, option hedging, and stress testing</p>
</div>

<div class="kpi-strip">
    <div class="kpi-mini"><div class="label" data-i18n="kpi_pv_cad">Portfolio Value (CAD)</div><div class="value dollar-amount">${portfolio_value:,.0f}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_pv_usd">Portfolio Value (USD)</div><div class="value dollar-amount">${portfolio_value_usd:,.0f}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_ann_ret">Annualized Return</div><div class="value positive">{ann_ret:.2%}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_sharpe">Sharpe Ratio</div><div class="value">{sharpe:.2f}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_max_dd">Max Drawdown</div><div class="value negative">{max_dd:.2%}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_beta">Beta to SPY</div><div class="value">{beta_str}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_pos_opt">Positions / Options</div><div class="value">{num_positions} / {num_options}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_delta_cad">Option Delta (CAD)</div><div class="value dollar-amount">${delta_cad:+,.0f}</div></div>
    <div class="kpi-mini"><div class="label" data-i18n="kpi_delta_usd">Option Delta (USD)</div><div class="value dollar-amount">${delta_usd:+,.0f}</div></div>
</div>

<div class="cards">
    <a class="card" href="positions.html">
        <div class="icon">&#128202;</div>
        <h2 data-i18n="card_positions">Positions</h2>
        <p data-i18n="card_positions_desc">All portfolio positions: stocks, ETFs, mutual funds, cash. Market values, weights, beta, and industry. Sortable columns.</p>
    </a>
    <a class="card" href="options.html">
        <div class="icon">&#128203;</div>
        <h2 data-i18n="card_options">Options</h2>
        <p data-i18n="card_options_desc">All option contracts with delta exposure analysis. Calls, puts, spreads, and their hedging impact on the portfolio.</p>
    </a>
    <a class="card" href="correlation_matrix.html">
        <div class="icon">&#128279;</div>
        <h2 data-i18n="card_correlation">Correlation Matrix</h2>
        <p data-i18n="card_correlation_desc">Pairwise return correlations with heatmap. Click tickers to sort. Hover cells for ticker pair details.</p>
    </a>
    <a class="card" href="risk_metrics.html">
        <div class="icon">&#9888;</div>
        <h2 data-i18n="card_risk">Risk Metrics</h2>
        <p data-i18n="card_risk_desc">VaR, Sharpe, Sortino, Calmar, Maximum Drawdown, Beta, option hedging impact. Hover cards for term explanations.</p>
    </a>
    <a class="card" href="stress_testing.html">
        <div class="icon">&#128293;</div>
        <h2 data-i18n="card_stress">Stress Testing</h2>
        <p data-i18n="card_stress_desc">Scenario analysis from -50% crash to +50% rally, showing both unhedged and option-hedged impacts with 1Y return context.</p>
    </a>
    <a class="card" href="sector_exposure.html">
        <div class="icon">&#127991;</div>
        <h2 data-i18n="card_exposure">Sector, Currency &amp; Account Exposure</h2>
        <p data-i18n="card_exposure_desc">Portfolio breakdown by sector allocation (incl. option notional), currency denomination, and brokerage account.</p>
    </a>
</div>
<p class="disclaimer" data-i18n="disclaimer">Disclaimer: This dashboard is for informational and educational purposes only and is not investment advice.</p>
<p class="timestamp"><span data-i18n="generated">Generated:</span> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
</body></html>"""
    return html


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("Stock Portfolio Analytics Report Generator v2")
    print("=" * 60)

    # 1. Read all data
    portfolio_df, opts_df, usd_cad_rate = read_portfolio(PORTFOLIO_FILE)
    print(f"  Loaded {len(portfolio_df)} portfolio positions")
    print(f"  Loaded {len(opts_df)} option contracts")
    print(f"  USD/CAD rate: {usd_cad_rate}")

    # 2. Portfolio value
    portfolio_value = portfolio_df["Mkt Value (CAD)"].sum()
    print(f"  Total portfolio value (CAD): ${portfolio_value:,.0f}")

    for ptype in portfolio_df["PositionType"].unique():
        sub = portfolio_df[portfolio_df["PositionType"] == ptype]["Mkt Value (CAD)"].sum()
        print(f"    {ptype}: ${sub:,.0f}")

    # 3. Collect all tradeable tickers (portfolio + option underlyings)
    non_tradeable = {"Cash", "Short Cash"}
    stock_tickers = [t for t in portfolio_df["Symbol"].unique() if t not in non_tradeable]
    option_underlyings = opts_df["Symbol"].unique().tolist()
    extra_tickers = [t for t in option_underlyings if t not in stock_tickers and t not in non_tradeable]
    all_tickers = sorted(set(stock_tickers + extra_tickers))
    print(f"\n  Unique tradeable tickers (stocks+options): {len(all_tickers)}")

    # 4. Fetch fundamentals (sector, beta, industry) from yfinance
    fund_df = fetch_fundamentals(all_tickers)

    # Back-fill Sector on portfolio_df and opts_df from fundamentals
    sector_map = dict(zip(fund_df["Symbol"], fund_df["Sector"]))
    portfolio_df["Sector"] = portfolio_df["Symbol"].map(sector_map).fillna("")
    opts_df["Sector"] = opts_df["Symbol"].map(sector_map).fillna("")

    # 5. Fetch live option prices and attach to opts_df
    opt_prices = fetch_option_prices(opts_df)
    opts_df["Opt Price"] = opt_prices

    # 6. Option delta exposure
    print("\nComputing option delta exposure...")
    option_delta_df, total_delta_usd = compute_option_delta_exposure(opts_df, usd_cad_rate=usd_cad_rate)
    print(f"  Net delta (USD): ${total_delta_usd:+,.0f}")
    print(f"  Net delta (CAD): ${total_delta_usd * usd_cad_rate:+,.0f}")

    # 7. Fetch price history
    prices = fetch_price_history(all_tickers)
    if prices.empty:
        print("ERROR: Could not fetch price data. Exiting.")
        sys.exit(1)

    # 8. Compute returns
    returns = compute_returns(prices)
    print(f"  Computed returns: {returns.shape[0]} days x {returns.shape[1]} tickers")

    # 9. Portfolio weights
    weight_series = pd.Series(0.0, index=returns.columns)
    for t in returns.columns:
        if t in portfolio_df["Symbol"].values:
            total_mkt = portfolio_df[portfolio_df["Symbol"] == t]["Mkt Value (CAD)"].sum()
            weight_series[t] = total_mkt

    if not option_delta_df.empty:
        for _, orow in option_delta_df.iterrows():
            sym = orow["Symbol"]
            if sym in weight_series.index:
                weight_series[sym] += orow["Notional Delta (CAD)"]

    total_weight = weight_series.sum()
    if total_weight != 0:
        weight_series = weight_series / total_weight

    # 10. Correlation matrix
    print("\nComputing correlation matrix...")
    corr = compute_correlation_matrix(returns)

    # 11. Risk metrics (with option hedging)
    print("Computing risk metrics...")
    metrics, portfolio_returns = compute_risk_metrics(
        returns, weight_series, portfolio_value, total_delta_usd, usd_cad_rate=usd_cad_rate
    )

    for key in ["Annualized Return", "Annualized Volatility", "Sharpe Ratio", "Sortino Ratio",
                 "Maximum Drawdown", "Beta to SPY", "VaR 95%", "Calmar Ratio"]:
        v = metrics.get(key)
        if isinstance(v, float):
            if key in ("Annualized Return", "Annualized Volatility", "Maximum Drawdown", "VaR 95%"):
                print(f"  {key}: {v:.2%}")
            else:
                print(f"  {key}: {v:.3f}")
        else:
            print(f"  {key}: {v}")

    # 12. Stress testing (with option hedging)
    beta_val = metrics.get("Beta to SPY", 1.0)
    if not isinstance(beta_val, (int, float)):
        beta_val = 1.0
    print("\nComputing stress testing scenarios...")
    stress_df = compute_stress_testing(
        portfolio_returns, weight_series, returns, portfolio_value,
        beta_val, total_delta_usd, usd_cad_rate=usd_cad_rate,
    )

    # 13. Individual risk (compute beta from SPY returns for tickers without it)
    print("Computing individual position risk...")
    spy_returns_for_beta = None
    if "SPY" in returns.columns:
        spy_returns_for_beta = returns["SPY"]
    else:
        try:
            spy_data = yf.download("SPY", period="1y", auto_adjust=True, progress=False)
            if not spy_data.empty:
                spy_close = spy_data["Close"]
                if isinstance(spy_close, pd.DataFrame):
                    spy_close = spy_close.iloc[:, 0]
                spy_returns_for_beta = np.log(spy_close / spy_close.shift(1)).dropna()
        except Exception:
            pass
    individual_risk = compute_individual_risk(returns, fund_df, spy_returns=spy_returns_for_beta)
    print(f"  Computed risk metrics for {len(individual_risk)} tickers")

    # 14. Generate HTML reports
    print("\nGenerating HTML reports...")

    reports = {
        "index.html": generate_index_html(portfolio_value, metrics, len(portfolio_df), len(opts_df), usd_cad_rate=usd_cad_rate),
        "positions.html": generate_html_positions(portfolio_df, opts_df, fund_df, portfolio_value, usd_cad_rate=usd_cad_rate),
        "options.html": generate_html_options(opts_df, option_delta_df, total_delta_usd, usd_cad_rate=usd_cad_rate),
        "correlation_matrix.html": generate_html_correlation(corr),
        "risk_metrics.html": generate_html_risk_metrics(metrics, individual_risk, portfolio_value, usd_cad_rate=usd_cad_rate),
        "stress_testing.html": generate_html_stress_testing(
            stress_df, portfolio_value, beta_val, total_delta_usd,
            usd_cad_rate=usd_cad_rate, ann_return=metrics.get("Annualized Return", 0),
        ),
        "sector_exposure.html": generate_html_sector_exposure(
            portfolio_df, opts_df, portfolio_value, usd_cad_rate=usd_cad_rate,
        ),
    }

    for filename, content in reports.items():
        filepath = OUTPUT_DIR / filename
        filepath.write_text(content, encoding="utf-8")
        size_kb = filepath.stat().st_size / 1024
        print(f"  Written: {filename} ({size_kb:.1f} KB)")

    # JSON metrics
    metrics_json = {}
    for k, v in metrics.items():
        if isinstance(v, (np.floating, float)):
            metrics_json[k] = round(float(v), 6)
        elif isinstance(v, (np.integer, int)):
            metrics_json[k] = int(v)
        else:
            metrics_json[k] = str(v)
    json_path = OUTPUT_DIR / "risk_metrics.json"
    json_path.write_text(json.dumps(metrics_json, indent=2), encoding="utf-8")
    print(f"  Written: risk_metrics.json")

    print("\n" + "=" * 60)
    print("BUILD COMPLETE - 7 HTML reports + 1 JSON file generated")
    print(f"Reports saved to: {OUTPUT_DIR}")
    print(f"Open index.html in a browser to view the dashboard.")
    print("=" * 60)


if __name__ == "__main__":
    main()
