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
from opencc import OpenCC

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


def fetch_latest_prices(price_history_df):
    """Extract the latest closing price for each ticker from price history.

    Uses the last available row of the already-fetched price history DataFrame
    so no additional API calls are needed.

    Returns a dict of {symbol: latest_price}.
    """
    if price_history_df.empty:
        return {}
    latest = price_history_df.iloc[-1]
    return {sym: float(latest[sym]) for sym in price_history_df.columns if pd.notna(latest[sym])}


def compute_returns(prices):
    """Compute daily log returns."""
    return np.log(prices / prices.shift(1)).dropna()


def compute_correlation_matrix(returns):
    """Compute correlation matrix of daily returns."""
    return returns.corr()


def compute_rrg_data(prices, weight_series, benchmark="SPY",
                     trail_days=63, history_days=252, warmup_days=180):
    """Compute Relative Rotation Graph (RRG) data for each holding vs the benchmark.

    Daily-sampled. ``trail_days`` ~ 3 months (63 trading days), ``history_days``
    ~ 1 year (252 trading days). ``warmup_days`` of extra history is fetched
    before the display window so the rolling RS / momentum windows are fully
    populated even at the earliest displayed date. Returns a dict suitable for
    JSON serialization::

        {
            "dates":   ["YYYY-MM-DD", ...],     # daily closing dates
            "tickers": ["AAPL", "MSFT", ...],
            "series":  {ticker: [[rsr, rsm], ...], ...},
            "trail_steps": 63,
            "benchmark": "SPY",
        }
    """
    empty = {"dates": [], "tickers": [], "series": {},
             "trail_steps": trail_days, "benchmark": benchmark}
    if prices is None or prices.empty:
        return empty

    # Identify which holdings have a positive weight (skip cash, benchmark, etc.).
    candidate_tickers = [
        t for t in prices.columns
        if t != benchmark
        and t in (weight_series.index if weight_series is not None else prices.columns)
        and (weight_series is None or float(weight_series.get(t, 0)) > 0)
    ]
    if not candidate_tickers:
        return empty

    # Fetch an extended history (display window + warmup) so rolling stats are
    # fully warmed up at the earliest visible date instead of bottoming out at
    # 100 for the first ~3 months.
    needed_calendar_days = int((history_days + warmup_days) * 1.45) + 30
    extended = None
    try:
        symbols = list({benchmark, *candidate_tickers})
        end_dt = datetime.now()
        start_dt = end_dt - timedelta(days=needed_calendar_days)
        ext_data = yf.download(
            symbols,
            start=start_dt.strftime("%Y-%m-%d"),
            end=end_dt.strftime("%Y-%m-%d"),
            auto_adjust=True,
            progress=False,
        )
        if not ext_data.empty:
            if isinstance(ext_data.columns, pd.MultiIndex):
                extended = ext_data["Close"].copy()
            else:
                extended = ext_data[["Close"]].copy()
                extended.columns = symbols
    except Exception:
        extended = None

    if extended is not None and not extended.empty:
        # Merge extended history with anything already in `prices` so we keep
        # any local overrides (e.g. backfilled cells) for the display window.
        merged = extended.copy()
        for col in prices.columns:
            if col in merged.columns:
                merged[col] = merged[col].combine_first(prices[col])
            else:
                merged[col] = prices[col]
        daily = merged.ffill().bfill()
    else:
        daily = prices.ffill().bfill()

    if benchmark not in daily.columns:
        try:
            bench_data = yf.download(benchmark, period="2y",
                                     auto_adjust=True, progress=False)
            if bench_data.empty:
                return empty
            bench_close = bench_data["Close"]
            if isinstance(bench_close, pd.DataFrame):
                bench_close = bench_close.iloc[:, 0]
            daily = daily.copy()
            daily[benchmark] = bench_close.reindex(daily.index).ffill().bfill()
        except Exception:
            return empty

    if benchmark not in daily.columns or len(daily) < 30:
        return empty

    bench = daily[benchmark]

    rs_window = 63    # ~3 months of trading days
    mom_window = 21   # ~1 month for momentum smoothing

    series = {}
    for ticker in candidate_tickers:
        if ticker not in daily.columns:
            continue
        ts = daily[ticker]
        if ts.isna().all():
            continue
        rs = (ts / bench) * 100.0
        sma_rs = rs.rolling(rs_window, min_periods=max(5, rs_window // 3)).mean()
        std_rs = rs.rolling(rs_window, min_periods=max(5, rs_window // 3)).std()
        rsr = 100 + (rs - sma_rs) / std_rs.replace(0, np.nan)

        sma_rsr = rsr.rolling(mom_window, min_periods=max(3, mom_window // 3)).mean()
        std_rsr = rsr.rolling(mom_window, min_periods=max(3, mom_window // 3)).std()
        rsm = 100 + (rsr - sma_rsr) / std_rsr.replace(0, np.nan)

        rsr = rsr.fillna(100).clip(93, 107)
        rsm = rsm.fillna(100).clip(93, 107)

        if len(rsr) == 0:
            continue
        series[ticker] = [
            [round(float(a), 3), round(float(b), 3)]
            for a, b in zip(rsr.values, rsm.values)
        ]

    if not series:
        return empty

    dates = [d.strftime("%Y-%m-%d") for d in daily.index]
    if len(dates) > history_days:
        dates = dates[-history_days:]
        for t in series:
            series[t] = series[t][-history_days:]

    tickers = sorted(series.keys())
    return {
        "dates": dates,
        "tickers": tickers,
        "series": series,
        "trail_steps": trail_days,
        "benchmark": benchmark,
    }


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
# HTML GENERATION — Single-Page Architecture with Tab Navigation
# ═══════════════════════════════════════════════════════════════════════════════

COMBINED_CSS = """
    * { box-sizing: border-box; }
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #0B1220; color: #E0E0E0; margin: 20px; }
    h1, h2 { color: #E0E0E0; border-bottom: 2px solid #2A3F5F; padding-bottom: 10px; }
    .info { color: #8899AA; margin-bottom: 20px; font-size: 14px; }
    .positive { color: #007A33; }
    .negative { color: #B81D13; }
    .timestamp { color: #8899AA; font-size: 11px; white-space: nowrap; margin: 0; }
    /* Tab system */
    .page { display: none; }
    .page.active { display: block; }
    /* Navigation */
    .nav { background: #1C2541; padding: 10px 16px; border-radius: 8px; margin-bottom: 20px; display: flex; gap: 16px; flex-wrap: wrap; align-items: center; position: sticky; top: 0; z-index: 50; }
    .nav a { color: #7AAFFF; text-decoration: none; font-size: 13px; padding: 4px 10px; border-radius: 4px; cursor: pointer; }
    .nav a:hover { background: #2A3F5F; }
    .nav a.active { background: #3A7BD5; color: white; }
    .nav .spacer { flex: 1; }
    .nav .privacy-toggle { background: #2A3F5F; color: #D4A843; border: 1px solid #3A7BD5; padding: 4px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: bold; }
    .nav .privacy-toggle:hover { background: #3A7BD5; color: white; }
    .nav .lang-toggle { background: #2A3F5F; color: #7AAFFF; border: 1px solid #3A7BD5; padding: 4px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: bold; margin-left: 8px; }
    .nav .lang-toggle:hover { background: #3A7BD5; color: white; }
    .nav .lang-flags { display: inline-flex; gap: 0; margin-left: 8px; align-items: center; background: #2A3F5F; border: 1px solid #3A7BD5; border-radius: 4px; overflow: hidden; padding: 0; }
    .nav .flag-btn { background: transparent; border: 0; padding: 4px 6px; cursor: pointer; display: inline-flex; align-items: center; line-height: 1; }
    .nav .flag-btn + .flag-btn { border-left: 1px solid #3A7BD5; }
    .nav .flag-btn:hover { background: #3A7BD5; }
    .nav .flag-btn.active { background: #3A7BD5; box-shadow: 0 0 0 1px #D4A843 inset; }
    .nav .flag-btn svg { display: block; width: 24px; height: 16px; border-radius: 2px; }
    /* Money toggle switch ($ visible / $ slashed) */
    .nav .money-toggle { display: inline-flex; align-items: center; gap: 8px; margin-left: 10px; cursor: pointer; user-select: none; }
    .nav .money-toggle input { display: none; }
    .nav .money-toggle .switch { position: relative; width: 42px; height: 22px; background: #2A3F5F; border: 1px solid #3A7BD5; border-radius: 22px; transition: background .15s; }
    .nav .money-toggle .switch::after { content: ''; position: absolute; top: 2px; left: 2px; width: 16px; height: 16px; background: #D4A843; border-radius: 50%; transition: transform .15s; }
    .nav .money-toggle input:checked + .switch { background: #3A7BD5; }
    .nav .money-toggle input:checked + .switch::after { transform: translateX(20px); background: #FFF; }
    .nav .money-icon { position: relative; width: 18px; height: 18px; display: inline-flex; align-items: center; justify-content: center; font-weight: bold; font-size: 14px; color: #D4A843; }
    .nav .money-icon.hide-icon::before { content: ''; position: absolute; left: -2px; top: 50%; width: 22px; height: 2px; background: #E04444; transform: translateY(-50%) rotate(-30deg); border-radius: 1px; }
    /* Theme toggle */
    .nav .theme-toggle { background: #2A3F5F; color: #D4A843; border: 1px solid #3A7BD5; width: 30px; height: 30px; border-radius: 50%; cursor: pointer; margin-left: 10px; display: inline-flex; align-items: center; justify-content: center; font-size: 14px; line-height: 1; padding: 0; }
    .nav .theme-toggle:hover { background: #3A7BD5; color: white; }
    /* Dashboard rotation panel */
    .dashboard-rrg { margin-top: 24px; padding: 18px; background: #131C2E; border: 1px solid #2A3F5F; border-radius: 8px; }
    .dashboard-rrg h2 { border-bottom: 1px solid #2A3F5F; margin-top: 0; }
    .rrg-mini canvas { max-height: 440px; }
    .rrg-legend-item { cursor: pointer; transition: opacity 0.15s; }
    .rrg-legend-item:hover { opacity: 0.85; }
    .rrg-legend-item.rrg-legend-off { opacity: 0.35; text-decoration: line-through; }
    /* Light theme overrides */
    body.light-theme { background: #F4F6FA; color: #1C2541; }
    body.light-theme h1, body.light-theme h2 { color: #1C2541; border-bottom-color: #C5CEDC; }
    body.light-theme .info { color: #5A6A82; }
    body.light-theme .timestamp { color: #5A6A82; }
    body.light-theme .nav { background: #FFFFFF; border: 1px solid #C5CEDC; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    body.light-theme .nav a { color: #2A5BB5; }
    body.light-theme .nav a:hover { background: #E5EBF5; }
    body.light-theme .nav a.active { background: #2A5BB5; color: #FFFFFF; }
    body.light-theme .nav .lang-flags,
    body.light-theme .nav .theme-toggle,
    body.light-theme .nav .money-toggle .switch { background: #F0F3F8; border-color: #2A5BB5; }
    body.light-theme .nav .flag-btn:hover { background: #D7E0F0; }
    body.light-theme .nav .flag-btn.active { background: #2A5BB5; }
    body.light-theme .nav .flag-btn + .flag-btn { border-left-color: #C5CEDC; }
    body.light-theme .nav .theme-toggle { color: #B5840F; }
    body.light-theme .nav .theme-toggle:hover { background: #2A5BB5; color: #FFF; }
    body.light-theme .nav .money-icon { color: #B5840F; }
    /* Privacy switch knob: on=blue→white knob, off=light bg→navy knob (visible). */
    body.light-theme .nav .money-toggle .switch::after { background: #2A5BB5; }
    body.light-theme .nav .money-toggle input:checked + .switch { background: #2A5BB5; }
    body.light-theme .nav .money-toggle input:checked + .switch::after { background: #FFFFFF; }
    body.light-theme .dashboard-rrg,
    body.light-theme .rrg-wrap { background: #FFFFFF; border-color: #C5CEDC; }
    body.light-theme .rrg-canvas-wrap canvas { background: #FFFFFF; }
    body.light-theme .rrg-date { color: #1C2541; }
    body.light-theme .rrg-legend-row { border-top-color: #C5CEDC; }
    body.light-theme .rrg-legend-item { color: #1C2541; }
    /* High-specificity overrides for dashboard / per-page boxes. */
    body.light-theme #page-dashboard .header,
    body.light-theme #page-dashboard .kpi-mini,
    body.light-theme #page-dashboard .card,
    body.light-theme #page-positions .kpi-card,
    body.light-theme #page-options .kpi-card,
    body.light-theme #page-risk .kpi-card,
    body.light-theme #page-stress .kpi-card,
    body.light-theme #page-exposure .exposure-box,
    body.light-theme .table-container { background: #FFFFFF; border-color: #C5CEDC; color: #1C2541; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
    body.light-theme #page-dashboard .header { background: linear-gradient(135deg, #FFFFFF 0%, #E5EBF5 100%); }
    body.light-theme #page-dashboard .header h1,
    body.light-theme #page-dashboard .card h2 { color: #1C2541; }
    body.light-theme #page-dashboard .header p,
    body.light-theme #page-dashboard .card p,
    body.light-theme #page-dashboard .kpi-mini .label,
    body.light-theme #page-positions .kpi-label,
    body.light-theme #page-options .kpi-label,
    body.light-theme #page-risk .kpi-label,
    body.light-theme #page-stress .kpi-label { color: #5A6A82; }
    body.light-theme #page-dashboard .card:hover { box-shadow: 0 8px 20px rgba(42,91,181,0.18); border-color: #2A5BB5; }
    body.light-theme #page-dashboard .kpi-mini .value,
    body.light-theme #page-positions .kpi-value,
    body.light-theme #page-options .kpi-value,
    body.light-theme #page-risk .kpi-value,
    body.light-theme #page-stress .kpi-value { color: #B5840F; }
    body.light-theme .disclaimer { color: #5A6A82; }
    /* Tables in light mode. */
    body.light-theme table { color: #1C2541; }
    body.light-theme th { background: #E5EBF5; color: #1C2541; }
    body.light-theme th:hover { background: #D7E0F0; }
    body.light-theme thead th { background: #E5EBF5 !important; color: #1C2541 !important; }
    body.light-theme tr:nth-child(even) { background: #F4F6FA; }
    body.light-theme tr:nth-child(odd) { background: #FFFFFF; }
    body.light-theme tr:hover { background: #E5EBF5; }
    body.light-theme td { border-bottom-color: #E1E6F0; }
    /* Correlation cells: leave inline-style backgrounds alone, but tighten frame */
    body.light-theme #page-correlation th.row-header,
    body.light-theme #page-correlation td.row-header { background: #E5EBF5; color: #1C2541; }
    body.light-theme #cell-tooltip { background: #FFFFFF; color: #1C2541; border-color: #2A5BB5; box-shadow: 0 4px 12px rgba(42,91,181,0.18); }
    body.light-theme .legend { color: #1C2541; }
    body.light-theme .opt-badge { background: #B5840F; color: #FFFFFF; }
    body.light-theme .opt-only-badge { background: #5A4BB5; color: #FFFFFF; }
    /* RRG (Relative Rotation Graph) */
    .rrg-wrap { background: #131C2E; border: 1px solid #2A3F5F; border-radius: 8px; padding: 16px; margin-top: 12px; }
    .rrg-canvas-wrap { position: relative; width: 100%; max-width: 900px; margin: 0 auto; }
    .rrg-canvas-wrap canvas { width: 100%; height: auto; display: block; background: #0F1729; border-radius: 6px; }
    .rrg-controls-row { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 14px; align-items: center; }
    .rrg-controls { display: flex; align-items: center; gap: 12px; flex: 1 1 320px; min-width: 0; }
    .rrg-play-btn { background: #3A7BD5; color: white; border: none; width: 36px; height: 36px; border-radius: 50%; cursor: pointer; font-size: 14px; flex: 0 0 auto; display: inline-flex; align-items: center; justify-content: center; }
    .rrg-play-btn:hover { background: #4A8CE5; }
    .rrg-slider { flex: 1; -webkit-appearance: none; appearance: none; height: 6px; background: #2A3F5F; border-radius: 3px; outline: none; }
    .rrg-slider::-webkit-slider-thumb { -webkit-appearance: none; appearance: none; width: 16px; height: 16px; border-radius: 50%; background: #D4A843; cursor: pointer; border: 2px solid #1C2541; }
    .rrg-slider::-moz-range-thumb { width: 16px; height: 16px; border-radius: 50%; background: #D4A843; cursor: pointer; border: 2px solid #1C2541; }
    .rrg-date { color: #D4D8E0; font-size: 13px; font-weight: bold; min-width: 80px; text-align: right; font-family: monospace; white-space: nowrap; }
    .rrg-trail-controls { flex: 0 1 280px; min-width: 180px; overflow: hidden; }
    .rrg-trail-label { font-size: 12px; color: #AABBCC; min-width: 70px; white-space: nowrap; flex: 0 0 auto; }
    .rrg-trail-slider { max-width: 120px; flex: 1 1 80px; }
    .rrg-trail-controls .rrg-date { min-width: 28px; text-align: center; }
    .rrg-legend-row { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin-top: 14px; padding-top: 12px; border-top: 1px solid #2A3F5F; }
    .rrg-legend { display: flex; flex-wrap: wrap; gap: 10px; flex: 1 1 0; }
    .rrg-legend-item { display: inline-flex; align-items: center; gap: 6px; font-size: 11px; color: #D4D8E0; }
    .rrg-legend-dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }
    .rrg-legend-actions { display: flex; gap: 8px; flex: 0 0 auto; }
    .rrg-mini-btn { background: #1F2D48; color: #D4D8E0; border: 1px solid #2A3F5F; border-radius: 4px; padding: 4px 12px; font-size: 11px; cursor: pointer; }
    .rrg-mini-btn:hover { background: #2A3F5F; color: #FFF; }
    body.light-theme .rrg-trail-label { color: #5A6A82; }
    body.light-theme .rrg-legend-row { border-top-color: #C5CEDC; }
    body.light-theme .rrg-mini-btn { background: #F0F3F8; color: #1C2541; border-color: #C5CEDC; }
    body.light-theme .rrg-mini-btn:hover { background: #2A5BB5; color: #FFFFFF; border-color: #2A5BB5; }
    @media (max-width: 700px) {
      .rrg-controls-row { flex-direction: column; }
      .rrg-trail-controls { flex: 1 1 100%; }
      .rrg-legend-actions { width: 100%; justify-content: flex-start; margin-top: 8px; }
    }
    /* Privacy */
    body.privacy-mode .dollar-amount { visibility: hidden; }
    body.privacy-mode .dollar-amount::after { content: '***'; visibility: visible; }
    /* Tables */
    table { border-collapse: collapse; width: 100%; }
    th { background: #1C2541; color: white; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; cursor: pointer; user-select: none; }
    th:hover { background: #2A3F5F; }
    td { padding: 8px 12px; border-bottom: 1px solid #1A2744; font-size: 13px; }
    tr:nth-child(even) { background: #111B2E; }
    tr:nth-child(odd) { background: #0D1526; }
    tr:hover { background: #1A2744; }
    /* Dashboard */
    #page-dashboard .header { background: linear-gradient(135deg, #1C2541 0%, #2A3F5F 100%); padding: 30px; border-radius: 12px; margin-bottom: 24px; }
    #page-dashboard .header h1 { margin: 0 0 8px 0; font-size: 28px; border: none; padding: 0; }
    #page-dashboard .header p { margin: 0; color: #8899AA; font-size: 14px; }
    #page-dashboard .kpi-strip { display: grid; grid-template-columns: repeat(5, 1fr); gap: 16px; margin-bottom: 28px; }
    @media (max-width: 1100px) { #page-dashboard .kpi-strip { grid-template-columns: repeat(3, 1fr); } }
    @media (max-width: 700px) { #page-dashboard .kpi-strip { grid-template-columns: repeat(2, 1fr); } }
    #page-dashboard .kpi-mini { background: #1C2541; border-radius: 8px; padding: 16px 20px; }
    #page-dashboard .kpi-mini .label { color: #8899AA; font-size: 11px; text-transform: uppercase; }
    #page-dashboard .kpi-mini .value { font-size: 22px; font-weight: bold; color: #D4A843; margin-top: 4px; }
    #page-dashboard .cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px; }
    #page-dashboard .card { background: #1C2541; border-radius: 10px; padding: 24px; transition: transform 0.2s, box-shadow 0.2s; cursor: pointer; text-decoration: none; color: inherit; display: block; border: 1px solid #2A3F5F; }
    #page-dashboard .card:hover { transform: translateY(-4px); box-shadow: 0 8px 25px rgba(0,0,0,0.3); border-color: #3A7BD5; }
    #page-dashboard .card h2 { margin: 0 0 10px 0; font-size: 18px; color: #E0E0E0; border: none; padding: 0; }
    #page-dashboard .card p { margin: 0; color: #8899AA; font-size: 13px; line-height: 1.5; }
    #page-dashboard .card .icon { font-size: 36px; margin-bottom: 12px; }
    .disclaimer { margin-top: 18px; color: #8899AA; font-size: 12px; }
    /* Positions */
    #page-positions .kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px; margin-bottom: 24px; }
    #page-positions .kpi-card { background: #1C2541; border-radius: 8px; padding: 14px; border-left: 4px solid #3A7BD5; }
    #page-positions .kpi-label { color: #8899AA; font-size: 11px; text-transform: uppercase; }
    #page-positions .kpi-value { font-size: 20px; font-weight: bold; color: #D4A843; }
    .weight-bar { height: 8px; background: #3A7BD5; border-radius: 4px; min-width: 2px; }
    .opt-badge { background: #D4A843; color: #0B1220; font-size: 10px; padding: 1px 5px; border-radius: 3px; margin-left: 4px; font-weight: bold; }
    .opt-only-badge { background: #7A5BD5; color: white; font-size: 10px; padding: 1px 5px; border-radius: 3px; margin-left: 4px; font-weight: bold; }
    /* Options */
    #page-options .kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 14px; margin-bottom: 24px; }
    #page-options .kpi-card { background: #1C2541; border-radius: 8px; padding: 14px; border-left: 4px solid #D4A843; }
    #page-options .kpi-label { color: #8899AA; font-size: 11px; text-transform: uppercase; }
    #page-options .kpi-value { font-size: 20px; font-weight: bold; color: #D4A843; }
    .call { color: #00C49A; }
    .put { color: #FF6B6B; }
    /* Correlation */
    #page-correlation .table-container { overflow-x: auto; position: relative; }
    #page-correlation table { font-size: 11px; width: auto; }
    #page-correlation th { padding: 6px 8px; white-space: nowrap; font-size: 11px; text-transform: none; letter-spacing: normal; }
    #page-correlation th.row-header { position: sticky; left: 0; z-index: 3; background: #1C2541; }
    #page-correlation td { padding: 5px 7px; text-align: center; border: 1px solid #1A2744; font-size: 11px; white-space: nowrap; }
    #page-correlation td.row-header { position: sticky; left: 0; background: #1C2541; color: white; font-weight: bold; z-index: 1; text-align: left; cursor: pointer; }
    #page-correlation td.row-header:hover { background: #2A3F5F; }
    .legend { margin-top: 20px; display: flex; gap: 20px; align-items: center; font-size: 13px; flex-wrap: wrap; }
    .legend-item { display: flex; align-items: center; gap: 6px; }
    .legend-box { width: 20px; height: 20px; border-radius: 3px; }
    #cell-tooltip { position: fixed; background: #1C2541; color: #E0E0E0; border: 1px solid #3A7BD5; padding: 8px 12px; border-radius: 6px; font-size: 12px; pointer-events: none; z-index: 100; display: none; box-shadow: 0 4px 12px rgba(0,0,0,0.5); }
    /* Risk Metrics */
    #page-risk .kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; margin-bottom: 30px; }
    #page-risk .kpi-card { background: #1C2541; border-radius: 8px; padding: 18px; border-left: 4px solid #3A7BD5; position: relative; cursor: help; }
    #page-risk .kpi-card.pos { border-left-color: #007A33; }
    #page-risk .kpi-card.neg { border-left-color: #B81D13; }
    #page-risk .kpi-card.neut { border-left-color: #D4A843; }
    #page-risk .kpi-label { color: #8899AA; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
    #page-risk .kpi-value { font-size: 24px; font-weight: bold; }
    #page-risk .kpi-value.pos { color: #007A33; }
    #page-risk .kpi-value.neg { color: #B81D13; }
    #page-risk .kpi-value.neut { color: #D4A843; }
    .kpi-tooltip { display: none; position: absolute; bottom: calc(100% + 8px); left: 0; right: 0; background: #0D1526; border: 1px solid #3A7BD5; border-radius: 6px; padding: 10px 14px; font-size: 12px; color: #CCDDEE; line-height: 1.5; z-index: 20; box-shadow: 0 4px 16px rgba(0,0,0,0.5); pointer-events: none; }
    #page-risk .kpi-card:hover .kpi-tooltip { display: block; }
    .help-icon { font-size: 11px; color: #556677; margin-left: 4px; }
    .section-label { background: #2A3F5F; color: #AAC0DD; padding: 8px 14px; border-radius: 6px; font-size: 13px; margin: 24px 0 12px 0; display: inline-block; }
    /* Stress Testing */
    #page-stress .summary-box { background: #1C2541; padding: 16px 24px; border-radius: 8px; margin-bottom: 24px; display: flex; gap: 30px; flex-wrap: wrap; font-size: 14px; align-items: center; }
    #page-stress .summary-box .label { color: #8899AA; font-size: 11px; text-transform: uppercase; }
    #page-stress .summary-box .value { color: #D4A843; font-weight: bold; font-size: 18px; }
    .ccy-switch-wrap { display: flex; align-items: center; gap: 8px; margin-left: auto; font-size: 13px; font-weight: bold; white-space: nowrap; }
    .ccy-switch-wrap .ccy-label { color: #556677; transition: color .2s; }
    .ccy-switch-wrap .ccy-label.active { color: #D4A843; }
    .ccy-switch { position: relative; width: 44px; height: 24px; flex-shrink: 0; }
    .ccy-switch input { opacity: 0; width: 0; height: 0; }
    .ccy-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background: #3A7BD5; border-radius: 24px; transition: background .3s; }
    .ccy-slider::before { content: ''; position: absolute; height: 18px; width: 18px; left: 3px; bottom: 3px; background: white; border-radius: 50%; transition: transform .3s; }
    .ccy-switch input:checked + .ccy-slider { background: #D4A843; }
    .ccy-switch input:checked + .ccy-slider::before { transform: translateX(20px); }
    /* Exposure */
    .grid-3col { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 30px; }
    @media (max-width: 1200px) { .grid-3col { grid-template-columns: 1fr 1fr; } }
    @media (max-width: 800px) { .grid-3col { grid-template-columns: 1fr; } }
    #page-exposure .weight-bar { height: 10px; border-radius: 5px; min-width: 2px; }
    .sector-colors { background: linear-gradient(90deg, #3A7BD5, #00C49A); }
    .currency-colors { background: linear-gradient(90deg, #D4A843, #E07A3A); }
    .account-colors { background: linear-gradient(90deg, #7A5BD5, #BD5BA8); }
"""

TAB_PAGES = [
    ("dashboard", "Dashboard", "nav_dashboard"),
    ("positions", "Positions", "nav_positions"),
    ("options", "Options", "nav_options"),
    ("correlation", "Correlation", "nav_correlation"),
    ("risk", "Risk Metrics", "nav_risk_metrics"),
    ("stress", "Stress Testing", "nav_stress_testing"),
    ("exposure", "Exposure", "nav_exposure"),
    ("rotation", "Rotation", "nav_rotation"),
]

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
    "opt_value_cad": {"en": "Options Value (CAD)", "zh": "期權市值 (加元)"},
    "th_opt_price": {"en": "Opt Price", "zh": "期權價格"},
    "th_contract_value": {"en": "Contract Value", "zh": "合約價值"},
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
    "m_total_pv_cad": {"en": "Total Portfolio Value (CAD)", "zh": "投資組合總值 (加元)"},
    "m_total_pv_usd": {"en": "Total Portfolio Value (USD)", "zh": "投資組合總值 (美元)"},
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
    "m_var95_cad": {"en": "VaR 95% (CAD)", "zh": "風險價值 95% (加元)"},
    "m_var95_usd": {"en": "VaR 95% (USD)", "zh": "風險價值 95% (美元)"},
    "m_var99d": {"en": "VaR 99% ($)", "zh": "風險價值 99% ($)"},
    "m_var99_cad": {"en": "VaR 99% (CAD)", "zh": "風險價值 99% (加元)"},
    "m_var99_usd": {"en": "VaR 99% (USD)", "zh": "風險價值 99% (美元)"},
    "m_skew": {"en": "Skewness", "zh": "偏度"},
    "m_kurt": {"en": "Kurtosis", "zh": "峰度"},
    "m_net_delta": {"en": "Net Delta (USD)", "zh": "淨Delta (美元)"},
    "m_net_delta_usd": {"en": "Net Delta (USD)", "zh": "淨Delta (美元)"},
    "m_net_delta_cad": {"en": "Net Delta (CAD)", "zh": "淨Delta (加元)"},
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
    "st_pv_cad": {"en": "Portfolio Value (CAD)", "zh": "投資組合價值 (加元)"},
    "st_pv_usd": {"en": "Portfolio Value (USD)", "zh": "投資組合價值 (美元)"},
    "st_delta_cad": {"en": "Option Delta (CAD)", "zh": "期權Delta (加元)"},
    "st_delta_usd": {"en": "Option Delta (USD)", "zh": "期權Delta (美元)"},
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
    "sc_mild_pullback": {"en": "Mild Pullback (-3%)", "zh": "輕微回調 (-3%)"},
    "sc_bubble": {"en": "Bubble (+40%)", "zh": "泡沫 (+40%)"},
    "sc_euphoric": {"en": "Euphoric Rally (+50%)", "zh": "狂熱上升 (+50%)"},
    "title_exposure": {"en": "Portfolio Exposure Analysis", "zh": "投資組合風險敞口分析"},
    "info_exposure": {"en": "Breakdown by sector (including option notional), currency, and brokerage account.", "zh": "按行業（含期權名義值）、貨幣及券商帳戶的分佈。"},
    "h2_sector": {"en": "Sector Exposure", "zh": "行業風險敞口"},
    "h2_currency": {"en": "Currency Exposure", "zh": "貨幣風險敞口"},
    "h2_account": {"en": "Account Exposure", "zh": "帳戶風險敞口"},
    "th_value_cad": {"en": "Value (CAD)", "zh": "價值 (加元)"},
    "th_value_usd": {"en": "Value (USD)", "zh": "價值 (美元)"},
    # Sector Rotation (RRG) tab
    "nav_rotation": {"en": "Rotation", "zh": "輪動"},
    "card_rotation": {"en": "Sector Rotation (RRG)", "zh": "板塊輪動 (RRG)"},
    "card_rotation_desc": {"en": "Relative Rotation Graph showing each holding's relative strength and momentum versus SPY. Drag the timeline to replay history.", "zh": "相對輪動圖：顯示每項持倉相對 SPY 的強度與動量。拖動時間軸可回放歷史。"},
    "title_rotation": {"en": "Relative Rotation Graph", "zh": "相對輪動圖"},
    "info_rotation": {"en": "Each dot is a holding plotted by its JdK RS-Ratio (x) and JdK RS-Momentum (y) versus SPY. Trail length is 3 months (~63 trading days). Press play or drag the slider to animate. Click a ticker in the legend to hide / show it.", "zh": "每個圓點代表一項持倉，依其相對 SPY 的 JdK RS 比率 (x 軸) 與 RS 動量 (y 軸) 繪製。尾跡長度為 3 個月 (約 63 個交易日)。按播放鍵或拖動滑桿可進行動畫。點擊圖例中的代號可隱藏／顯示對應的圓點。"},
    "rrg_leading": {"en": "Leading", "zh": "領先"},
    "rrg_weakening": {"en": "Weakening", "zh": "轉弱"},
    "rrg_lagging": {"en": "Lagging", "zh": "落後"},
    "rrg_improving": {"en": "Improving", "zh": "改善"},
    "rrg_x_axis": {"en": "JdK RS-Ratio", "zh": "JdK RS 比率"},
    "rrg_y_axis": {"en": "JdK RS-Momentum", "zh": "JdK RS 動量"},
    "rrg_play": {"en": "Play", "zh": "播放"},
    "rrg_pause": {"en": "Pause", "zh": "暫停"},
    "rrg_show_all": {"en": "Show All", "zh": "顯示全部"},
    "rrg_hide_all": {"en": "Hide All", "zh": "隱藏全部"},
    "rrg_trail_label": {"en": "Trail (days)", "zh": "尾跡長度（天）"},
    "rrg_benchmark": {"en": "Benchmark: SPY", "zh": "基準：SPY"},
    "rrg_no_data": {"en": "Insufficient price history to compute the rotation graph.", "zh": "價格歷史資料不足，無法計算輪動圖。"},
}


# ─── Multi-language helpers (HK / TW / CN) ──────────────────────────────────
# The TRANSLATIONS dict above stores Hong Kong Traditional Chinese under "zh".
# We programmatically derive Taiwan Traditional ("zh-tw") and Mainland
# Simplified ("zh-cn") variants using OpenCC + curated vocabulary swaps so
# investment terminology is locale-appropriate.

_CC_HK_TO_S = OpenCC("hk2s")  # HK Traditional -> Simplified Chinese (CN)

# Vocabulary swaps applied AFTER the HK source for Taiwan Traditional output.
# Characters stay Traditional; only investment-term wording is localized.
_HK_TO_TW_VOCAB = [
    ("互惠基金", "共同基金"),
    ("認購期權", "買權"),
    ("認沽期權", "賣權"),
    ("認購", "買權"),
    ("認沽", "賣權"),
    ("期權", "選擇權"),
    ("沽空", "放空"),
    ("對沖", "避險"),
    ("行業分類", "產業分類"),
    ("行業", "產業"),
    ("帳戶", "帳戶"),
    ("券商", "券商"),
    ("敞口", "曝險"),
    ("組合", "組合"),
    ("資訊", "資訊"),
    ("回撤", "回檔"),
    ("回報", "報酬"),
    ("互惠", "共同"),
    ("名義", "名目"),
    ("標的", "標的"),
    ("行情", "行情"),
    ("情景", "情境"),
]

# Vocabulary swaps applied AFTER OpenCC simplification for CN output.
_HK_TO_CN_VOCAB = [
    ("互惠基金", "共同基金"),
    ("认购期权", "看涨期权"),
    ("认沽期权", "看跌期权"),
    ("认购", "看涨期权"),
    ("认沽", "看跌期权"),
    ("沽空", "做空"),
    ("对冲", "对冲"),
    ("帐户", "账户"),
    ("行业分类", "行业分类"),
    ("券商", "券商"),
    ("敞口", "敞口"),
    ("回撤", "回撤"),
    ("回报", "收益"),
    ("名义", "名义"),
    ("情景", "情景"),
]


def _hk_to_tw(text: str) -> str:
    out = text
    for a, b in _HK_TO_TW_VOCAB:
        out = out.replace(a, b)
    return out


def _hk_to_cn(text: str) -> str:
    out = _CC_HK_TO_S.convert(text)
    for a, b in _HK_TO_CN_VOCAB:
        out = out.replace(a, b)
    return out


def _augment_translations():
    """Populate zh-tw and zh-cn variants for every translation key."""
    for key, langs in TRANSLATIONS.items():
        hk = langs.get("zh", langs.get("en", ""))
        if "zh-tw" not in langs:
            langs["zh-tw"] = _hk_to_tw(hk)
        if "zh-cn" not in langs:
            langs["zh-cn"] = _hk_to_cn(hk)


_augment_translations()


# ─── Flag toggle buttons (US / HK / TW / CN) ────────────────────────────────
# Inline SVG flags so the HTML is self-contained and renders consistently
# across platforms (Windows emoji fonts don't include flag glyphs).
_FLAG_SVG = {
    "en":    "<svg viewBox='0 0 60 40' xmlns='http://www.w3.org/2000/svg'><rect width='60' height='40' fill='#B22234'/><g fill='white'><rect y='3.08' width='60' height='3.08'/><rect y='9.23' width='60' height='3.08'/><rect y='15.38' width='60' height='3.08'/><rect y='21.54' width='60' height='3.08'/><rect y='27.69' width='60' height='3.08'/><rect y='33.85' width='60' height='3.08'/></g><rect width='24' height='21.54' fill='#3C3B6E'/><text x='12' y='13' fill='white' font-family='Arial' font-size='5' text-anchor='middle'>★★★</text><text x='12' y='19' fill='white' font-family='Arial' font-size='5' text-anchor='middle'>★★</text></svg>",
    "zh-HK": "<svg viewBox='0 0 60 40' xmlns='http://www.w3.org/2000/svg'><rect width='60' height='40' fill='#DE2910'/><g transform='translate(30 20)' fill='white'><circle r='6'/><g fill='#DE2910'><circle cx='0' cy='-3' r='1.6'/><circle cx='2.85' cy='-0.93' r='1.6'/><circle cx='1.76' cy='2.43' r='1.6'/><circle cx='-1.76' cy='2.43' r='1.6'/><circle cx='-2.85' cy='-0.93' r='1.6'/></g></g></svg>",
    "zh-TW": "<svg viewBox='0 0 60 40' xmlns='http://www.w3.org/2000/svg'><rect width='60' height='40' fill='#FE0000'/><rect width='30' height='20' fill='#000095'/><g transform='translate(15 10)'><circle r='5.5' fill='white'/><circle r='2.7' fill='#000095'/><polygon fill='white' points='0,-5.5 1,-1.4 5.2,-1.7 1.6,0.5 2.6,4.6 0,2.0 -2.6,4.6 -1.6,0.5 -5.2,-1.7 -1,-1.4'/></g></svg>",
    "zh-CN": "<svg viewBox='0 0 60 40' xmlns='http://www.w3.org/2000/svg'><rect width='60' height='40' fill='#DE2910'/><g fill='#FFDE00'><polygon points='10,4 11.8,9.5 17.5,9.5 12.9,12.9 14.7,18.4 10,15 5.3,18.4 7.1,12.9 2.5,9.5 8.2,9.5'/><circle cx='20' cy='4' r='1.2'/><circle cx='24' cy='8' r='1.2'/><circle cx='24' cy='13' r='1.2'/><circle cx='20' cy='17' r='1.2'/></g></svg>",
}
_FLAG_LABEL = {"en": "EN", "zh-HK": "港", "zh-TW": "台", "zh-CN": "简"}
_FLAG_TITLE = {
    "en": "English",
    "zh-HK": "Hong Kong (Traditional)",
    "zh-TW": "Taiwan (Traditional)",
    "zh-CN": "Mainland China (Simplified)",
}
FLAG_BUTTONS_HTML = '<span class="lang-flags">' + ''.join(
    f'<button type="button" class="flag-btn" data-lang="{code}" '
    f'title="{_FLAG_TITLE[code]}" onclick="setLanguage(\'{code}\')">'
    f'{svg}</button>'
    for code, svg in _FLAG_SVG.items()
) + '</span>'

PRIVACY_SWITCH_HTML = (
    '<label class="money-toggle" title="Show / hide dollar amounts">'
    '<input type="checkbox" id="privacy-switch" onchange="togglePrivacy()">'
    '<span class="switch"></span>'
    '<span id="privacy-icon" class="money-icon">$</span>'
    '</label>'
)

THEME_TOGGLE_HTML = (
    '<button type="button" id="theme-btn" class="theme-toggle" '
    'title="Toggle dark / light theme" onclick="toggleTheme()">&#9790;</button>'
)

SORTABLE_JS = """
function sortTable(tableId, colIdx, isNumeric) {
    const table = document.getElementById(tableId);
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const header = table.querySelectorAll('thead th')[colIdx];
    const curDir = header.getAttribute('data-sort-dir') || 'none';
    const newDir = curDir === 'asc' ? 'desc' : 'asc';
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


# ─── Section Generators ─────────────────────────────────────────────────────
# Each returns HTML content for a <div class="page"> section.
# Correlation and Stress also return page-specific JS code.


def _generate_dashboard_section(portfolio_value, metrics, num_positions, num_options, usd_cad_rate=1.37):
    """Generate dashboard tab content."""
    sharpe = metrics.get("Sharpe Ratio", 0)
    ann_ret = metrics.get("Annualized Return", 0)
    max_dd = metrics.get("Maximum Drawdown", 0)
    beta = metrics.get("Beta to SPY", "N/A")
    beta_str = f"{beta:.2f}" if isinstance(beta, float) else str(beta)
    delta_cad = metrics.get("Option Delta Exposure", 0)
    delta_usd = metrics.get("Net Delta (USD)", 0)
    portfolio_value_usd = portfolio_value / usd_cad_rate

    return f"""<div id="page-dashboard" class="page active">
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
    <div class="kpi-mini" style="visibility:hidden;"></div>
</div>
<div class="cards">
    <a class="card" onclick="showPage('positions')">
        <div class="icon">&#128202;</div>
        <h2 data-i18n="card_positions">Positions</h2>
        <p data-i18n="card_positions_desc">All portfolio positions: stocks, ETFs, mutual funds, cash. Market values, weights, beta, and industry. Sortable columns.</p>
    </a>
    <a class="card" onclick="showPage('options')">
        <div class="icon">&#128203;</div>
        <h2 data-i18n="card_options">Options</h2>
        <p data-i18n="card_options_desc">All option contracts with delta exposure analysis. Calls, puts, spreads, and their hedging impact on the portfolio.</p>
    </a>
    <a class="card" onclick="showPage('correlation')">
        <div class="icon">&#128279;</div>
        <h2 data-i18n="card_correlation">Correlation Matrix</h2>
        <p data-i18n="card_correlation_desc">Pairwise return correlations with heatmap. Click tickers to sort. Hover cells for ticker pair details.</p>
    </a>
    <a class="card" onclick="showPage('risk')">
        <div class="icon">&#9888;</div>
        <h2 data-i18n="card_risk">Risk Metrics</h2>
        <p data-i18n="card_risk_desc">VaR, Sharpe, Sortino, Calmar, Maximum Drawdown, Beta, option hedging impact. Hover cards for term explanations.</p>
    </a>
    <a class="card" onclick="showPage('stress')">
        <div class="icon">&#128293;</div>
        <h2 data-i18n="card_stress">Stress Testing</h2>
        <p data-i18n="card_stress_desc">Scenario analysis from -50% crash to +50% rally, showing both unhedged and option-hedged impacts with 1Y return context.</p>
    </a>
    <a class="card" onclick="showPage('exposure')">
        <div class="icon">&#127991;</div>
        <h2 data-i18n="card_exposure">Sector, Currency &amp; Account Exposure</h2>
        <p data-i18n="card_exposure_desc">Portfolio breakdown by sector allocation (incl. option notional), currency denomination, and brokerage account.</p>
    </a>
    <a class="card" onclick="showPage('rotation')">
        <div class="icon">&#128260;</div>
        <h2 data-i18n="card_rotation">Sector Rotation (RRG)</h2>
        <p data-i18n="card_rotation_desc">Relative Rotation Graph showing each holding's relative strength and momentum versus SPY. Drag the timeline to replay history.</p>
    </a>
</div>
<p class="disclaimer" data-i18n="disclaimer">Disclaimer: This dashboard is for informational and educational purposes only and is not investment advice.</p>
</div>"""


def _generate_positions_section(portfolio_df, opts_df, fund_df, portfolio_value, usd_cad_rate=1.37):
    """Generate positions tab content."""
    fund_cols = ["Symbol", "Type", "Beta", "Industry"]
    if "P/E" in fund_df.columns:
        fund_cols.append("P/E")
    available_cols = [c for c in fund_cols if c in fund_df.columns]
    merged = portfolio_df.merge(fund_df[available_cols], on="Symbol", how="left")
    merged.loc[merged["PositionType"] == "Cash", "Type"] = "Cash"
    merged.loc[merged["PositionType"] == "Cash", "Beta"] = 0.0
    merged["Weight"] = merged["Mkt Value (CAD)"] / portfolio_value if portfolio_value else 0

    option_counts = opts_df.groupby("Symbol").size().to_dict()

    # Option-only tickers
    portfolio_symbols = set(merged["Symbol"].unique())
    option_only_symbols = [s for s in opts_df["Symbol"].unique() if s not in portfolio_symbols]
    if option_only_symbols:
        opt_only_rows = []
        for sym in option_only_symbols:
            opt_rows = opts_df[opts_df["Symbol"] == sym]
            sector = opt_rows["Sector"].iloc[0] if "Sector" in opt_rows.columns and not opt_rows["Sector"].isna().all() else "-"
            currency = opt_rows["Currency"].iloc[0] if "Currency" in opt_rows.columns and not opt_rows["Currency"].isna().all() else "USD"
            fund_row = fund_df[fund_df["Symbol"] == sym]
            beta = fund_row["Beta"].values[0] if len(fund_row) > 0 and "Beta" in fund_row.columns and pd.notna(fund_row["Beta"].values[0]) else None
            industry = fund_row["Industry"].values[0] if len(fund_row) > 0 and "Industry" in fund_row.columns and pd.notna(fund_row["Industry"].values[0]) else "-"
            ftype = fund_row["Type"].values[0] if len(fund_row) > 0 and "Type" in fund_row.columns and pd.notna(fund_row["Type"].values[0]) else "-"
            opt_only_rows.append({
                "Symbol": sym, "Shares": 0, "Price": 0, "Currency": currency,
                "Mkt Value": 0, "Mkt Value (CAD)": 0, "Sector": sector,
                "Account": "Options Only", "PositionType": "Options Only",
                "Type": ftype, "Beta": beta, "Industry": industry, "Weight": 0,
            })
        opt_only_df = pd.DataFrame(opt_only_rows)
        merged = pd.concat([merged, opt_only_df], ignore_index=True)

    num_positions = len(merged)
    total_value = merged["Mkt Value (CAD)"].sum()
    num_stocks = len(merged[merged.get("Type", pd.Series(dtype=str)) == "Stock"]) if "Type" in merged.columns else 0
    num_etfs = len(merged[merged.get("Type", pd.Series(dtype=str)) == "ETF"]) if "Type" in merged.columns else 0
    num_mfs = len(merged[merged["PositionType"] == "Mutual Fund"])
    num_options = len(opts_df)
    cash_total = merged[merged["PositionType"] == "Cash"]["Mkt Value (CAD)"].sum()

    html = '<div id="page-positions" class="page">\n'
    html += '<h1 data-i18n="title_positions">Portfolio Positions</h1>\n'
    html += '<p class="info" data-i18n="info_positions">All positions including stocks, ETFs, mutual funds, and cash. Click column headers to sort.</p>\n'
    html += '<div class="kpi-grid">\n'

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

    html += '</tbody></table>\n</div>'
    return html


def _generate_options_section(opts_df, option_delta_df, total_delta_usd, usd_cad_rate=1.37):
    """Generate options tab content."""
    total_delta_cad = total_delta_usd * usd_cad_rate
    total_contracts = len(opts_df)
    calls = len(opts_df[opts_df["Type"] == "CALL"]) if "Type" in opts_df.columns else 0
    puts = len(opts_df[opts_df["Type"] == "PUT"]) if "Type" in opts_df.columns else 0

    total_opt_value_cad = 0.0
    if "Opt Price" in opts_df.columns:
        for _, row in opts_df.iterrows():
            opt_price = row.get("Opt Price", 0) or 0
            shares = row.get("Shares", 0) or 0
            currency = row.get("Currency", "USD")
            fx = usd_cad_rate if currency == "USD" else 1.0
            total_opt_value_cad += opt_price * shares * fx

    html = '<div id="page-options" class="page">\n'
    html += '<h1 data-i18n="title_options">Options Positions &amp; Delta Exposure</h1>\n'
    html += '<p class="info" data-i18n="info_options">All option contracts with live prices and estimated delta exposure. Negative shares = short position. Click headers to sort.</p>\n'
    html += '<div class="kpi-grid">\n'

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

    html += '</div>'
    return html


def _generate_correlation_section(corr_matrix):
    """Generate correlation matrix tab content and its JS code."""
    tickers = list(corr_matrix.columns)

    corr_json = {}
    for t1 in tickers:
        corr_json[t1] = {}
        for t2 in tickers:
            v = corr_matrix.loc[t1, t2]
            corr_json[t1][t2] = round(float(v), 4) if not np.isnan(v) else None

    header_cells = ""
    for t in tickers:
        header_cells += f'<th onclick="corrSortByCol(\'{t}\')">{t}</th>'

    body_rows = ""
    for t1 in tickers:
        cells = f'<td class="row-header" onclick="corrSortByRow(\'{t1}\')">{t1}</td>'
        for t2 in tickers:
            val = corr_matrix.loc[t1, t2]
            if np.isnan(val):
                cells += f'<td data-r="{t1}" data-c="{t2}" style="background:#1A2744;">-</td>'
            else:
                bg, fg = _corr_color(val, t1 == t2)
                cells += f'<td data-r="{t1}" data-c="{t2}" style="background:{bg};color:{fg};">{val:.2f}</td>'
        body_rows += f"<tr>{cells}</tr>\n"

    html = f"""<div id="page-correlation" class="page">
<h1 data-i18n="title_correlation">Portfolio Correlation Matrix</h1>
<p class="info" data-i18n="info_correlation">Correlation of daily log returns over the past 12 months. Click any ticker header or row label to sort. Hover cells to see ticker pair.</p>
<div class="table-container">
<table id="corr-table">
<thead><tr><th class="row-header" onclick="corrResetSort()" data-i18n="th_ticker">Ticker</th>
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
</div>"""

    js = f"""
const CORR = {json.dumps(corr_json)};
const TICKERS = {json.dumps(tickers)};
let corrCurrentSort = null;
let corrSortAsc = true;

function corrAttachTooltips() {{
    document.querySelectorAll('#corr-table td[data-r]').forEach(function(td) {{
        td.addEventListener('mouseenter', function(e) {{
            var r = this.getAttribute('data-r');
            var c = this.getAttribute('data-c');
            var v = this.textContent;
            var isZh = ((localStorage.getItem('portfolio_language') || 'en') + '').indexOf('zh') === 0;
            var tooltip = document.getElementById('cell-tooltip');
            tooltip.innerHTML = '<strong>' + r + '</strong> ' + (isZh ? '對' : 'vs') + ' <strong>' + c + '</strong><br>' + (isZh ? '相關性: ' : 'Correlation: ') + v;
            tooltip.style.display = 'block';
        }});
        td.addEventListener('mousemove', function(e) {{
            var tooltip = document.getElementById('cell-tooltip');
            tooltip.style.left = (e.clientX + 14) + 'px';
            tooltip.style.top = (e.clientY + 14) + 'px';
        }});
        td.addEventListener('mouseleave', function() {{ document.getElementById('cell-tooltip').style.display = 'none'; }});
    }});
}}
document.addEventListener('DOMContentLoaded', corrAttachTooltips);

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

function corrRebuildRow(row, sorted) {{
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

function corrSortByCol(ticker) {{
    if (corrCurrentSort === 'col_' + ticker) {{ corrSortAsc = !corrSortAsc; }} else {{ corrCurrentSort = 'col_' + ticker; corrSortAsc = false; }}
    var tbody = document.querySelector('#corr-table tbody');
    var rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort(function(a, b) {{
        var aLabel = a.querySelector('td.row-header').textContent;
        var bLabel = b.querySelector('td.row-header').textContent;
        var aVal = CORR[aLabel] && CORR[aLabel][ticker] != null ? CORR[aLabel][ticker] : -999;
        var bVal = CORR[bLabel] && CORR[bLabel][ticker] != null ? CORR[bLabel][ticker] : -999;
        return corrSortAsc ? aVal - bVal : bVal - aVal;
    }});
    rows.forEach(function(r) {{ tbody.appendChild(r); }});
    corrAttachTooltips();
}}

function corrSortByRow(ticker) {{
    if (corrCurrentSort === 'row_' + ticker) {{ corrSortAsc = !corrSortAsc; }} else {{ corrCurrentSort = 'row_' + ticker; corrSortAsc = false; }}
    var sorted = TICKERS.slice().sort(function(a, b) {{
        var aVal = CORR[ticker] && CORR[ticker][a] != null ? CORR[ticker][a] : -999;
        var bVal = CORR[ticker] && CORR[ticker][b] != null ? CORR[ticker][b] : -999;
        return corrSortAsc ? aVal - bVal : bVal - aVal;
    }});
    var thead = document.querySelector('#corr-table thead tr');
    while (thead.children.length > 1) thead.removeChild(thead.lastChild);
    sorted.forEach(function(t) {{
        var th = document.createElement('th');
        th.textContent = t;
        th.onclick = function() {{ corrSortByCol(t); }};
        thead.appendChild(th);
    }});
    var rows = document.querySelectorAll('#corr-table tbody tr');
    rows.forEach(function(row) {{ corrRebuildRow(row, sorted); }});
    corrAttachTooltips();
}}

function corrResetSort() {{
    showPage('correlation');
}}
"""
    return html, js


def _generate_risk_section(metrics, individual_risk_df, portfolio_value, usd_cad_rate=1.37):
    """Generate risk metrics tab content."""
    portfolio_value_usd = portfolio_value / usd_cad_rate

    # Set derived metrics
    metrics["Total Portfolio Value (CAD)"] = portfolio_value
    metrics["Total Portfolio Value (USD)"] = portfolio_value_usd
    metrics["VaR 95% (CAD)"] = abs(metrics.get("VaR 95%", 0)) * portfolio_value
    metrics["VaR 95% (USD)"] = abs(metrics.get("VaR 95%", 0)) * portfolio_value_usd
    metrics["VaR 99% (CAD)"] = abs(metrics.get("VaR 99%", 0)) * portfolio_value
    metrics["VaR 99% (USD)"] = abs(metrics.get("VaR 99%", 0)) * portfolio_value_usd
    net_delta_usd = metrics.get("Net Delta (USD)", 0)
    metrics["Net Delta (CAD)"] = net_delta_usd * usd_cad_rate if isinstance(net_delta_usd, (int, float)) else 0

    dollar_keys = {
        "Total Portfolio Value (CAD)", "Total Portfolio Value (USD)",
        "VaR 95% (CAD)", "VaR 95% (USD)", "VaR 99% (CAD)", "VaR 99% (USD)",
        "Net Delta (USD)", "Net Delta (CAD)",
    }

    def _fmt_metric(key):
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

    groups = [
        ("Portfolio Overview", ["Total Portfolio Value (CAD)", "Total Portfolio Value (USD)", "Annualized Return", "Annualized Volatility"]),
        ("Risk-Adjusted Returns", ["Sharpe Ratio", "Sortino Ratio", "Calmar Ratio"]),
        ("Drawdown & Market Risk", ["Maximum Drawdown", "Beta to SPY"]),
        ("Value at Risk", ["VaR 95%", "VaR 99%", "CVaR 95%", "VaR 95% (CAD)", "VaR 95% (USD)", "VaR 99% (CAD)", "VaR 99% (USD)"]),
        ("Distribution Shape", ["Skewness", "Kurtosis"]),
        ("Option Hedging", ["Net Delta (USD)", "Net Delta (CAD)", "Option Hedging Impact", "Hedged VaR 95%", "Hedged VaR 99%"]),
    ]

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

    html = f'<div id="page-risk" class="page">\n'
    html += f'<h1 data-i18n="title_risk">Portfolio Risk Metrics</h1>\n'
    html += f'<p class="info" data-i18n="info_risk">Risk analytics based on 1-year daily return history. Risk-free rate: {RISK_FREE_RATE:.1%}. Hover KPI cards for explanations.</p>\n'

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

    html += '</tbody></table>\n</div>'
    return html


def _generate_stress_section(stress_df, portfolio_value, beta, option_delta_usd=0, usd_cad_rate=1.37, ann_return=0):
    """Generate stress testing tab content and its JS code."""
    option_delta_cad = option_delta_usd * usd_cad_rate
    portfolio_value_usd = portfolio_value / usd_cad_rate
    option_delta_usd_val = option_delta_usd

    html = f"""<div id="page-stress" class="page">
<h1 data-i18n="title_stress">Portfolio Stress Testing</h1>
<p class="info" data-i18n="info_stress">Simulated impact of market-wide moves on portfolio value using beta, including option hedging effects.</p>
<div class="summary-box">
    <div class="item"><div class="label" data-i18n="st_pv_cad">Portfolio Value (CAD)</div><div class="value dollar-amount">${portfolio_value:,.0f}</div></div>
    <div class="item"><div class="label" data-i18n="st_pv_usd">Portfolio Value (USD)</div><div class="value dollar-amount">${portfolio_value_usd:,.0f}</div></div>
    <div class="item"><div class="label" data-i18n="st_beta">Portfolio Beta</div><div class="value">{beta:.3f}</div></div>
    <div class="item"><div class="label" data-i18n="st_delta_cad">Option Delta (CAD)</div><div class="value dollar-amount">${option_delta_cad:+,.0f}</div></div>
    <div class="item"><div class="label" data-i18n="st_delta_usd">Option Delta (USD)</div><div class="value dollar-amount">${option_delta_usd_val:+,.0f}</div></div>
    <div class="item"><div class="label" data-i18n="st_1y_ret">1Y Portfolio Return</div><div class="value {'positive' if ann_return > 0 else 'negative'}">{ann_return:.2%}</div></div>
    <div class="ccy-switch-wrap">
        <span class="ccy-label active" id="ccy-lbl-cad">CAD</span>
        <label class="ccy-switch"><input type="checkbox" id="ccy-toggle" onchange="stressToggleCurrency()"><span class="ccy-slider"></span></label>
        <span class="ccy-label" id="ccy-lbl-usd">USD</span>
    </div>
</div>
<table id="stress-table">
<thead><tr>
    <th data-i18n="th_scenario">Scenario</th>
    <th data-i18n="th_mkt_move">Market Move</th>
    <th data-i18n="th_unhedged_pct">Unhedged Impact (%)</th>
    <th data-i18n="th_unhedged_d">Unhedged Impact ($)</th>
    <th data-i18n="th_opt_pnl">Option Hedge P&amp;L ($)</th>
    <th data-i18n="th_hedged_pct">Hedged Impact (%)</th>
    <th data-i18n="th_hedged_d">Hedged Impact ($)</th>
    <th data-i18n="th_est_nav">Estimated NAV</th>
</tr></thead>
<tbody>
"""

    _sc_i18n = {
        "Depression (-50%)": "sc_bear", "Severe Bear (-40%)": "sc_deep",
        "Bear Market (-30%)": "sc_major_crash", "Market Crash (-20%)": "sc_crash",
        "Severe Correction (-15%)": "sc_severe", "Correction (-10%)": "sc_major_corr",
        "Flash Crash (-5%)": "sc_minor",
        "Mild Pullback (-3%)": "sc_mild_pullback",
        "Rally (+5%)": "sc_modest", "Strong Rally (+10%)": "sc_moderate",
        "Bull Run (+20%)": "sc_bull", "Euphoria (+30%)": "sc_surge",
        "Bubble (+40%)": "sc_bubble", "Mania (+50%)": "sc_euphoric",
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

    html += """</tbody></table>
</div>"""

    js = """
var stressShowCAD = true;
function stressToggleCurrency() {
    var cb = document.getElementById('ccy-toggle');
    stressShowCAD = !cb.checked;
    document.getElementById('ccy-lbl-cad').classList.toggle('active', stressShowCAD);
    document.getElementById('ccy-lbl-usd').classList.toggle('active', !stressShowCAD);
    document.querySelectorAll('.ccy-cell').forEach(function(td) {
        var raw = stressShowCAD ? parseFloat(td.getAttribute('data-cad')) : parseFloat(td.getAttribute('data-usd'));
        if (isNaN(raw)) return;
        var formatted = Math.abs(raw).toLocaleString('en-US', {maximumFractionDigits: 0});
        var origText = td.textContent.trim();
        if (origText.charAt(0) === '$' && (origText.charAt(1) === '+' || origText.charAt(1) === '-')) {
            td.textContent = '$' + (raw >= 0 ? '+' : '-') + formatted;
        } else {
            td.textContent = '$' + (raw < 0 ? '-' : '') + formatted;
        }
    });
}
"""
    return html, js


def _generate_exposure_section(portfolio_df, opts_df, portfolio_value, usd_cad_rate=1.37):
    """Generate exposure analysis tab content."""
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

    sector_data = all_positions.groupby("Sector").agg(
        total_value=("Mkt Value (CAD)", "sum"),
        total_value_usd=("Mkt Value (USD)", "sum"),
        num_positions=("Symbol", "count"),
    ).reset_index()
    total_sector = sector_data["total_value"].sum()
    sector_data["Weight"] = sector_data["total_value"] / total_sector if total_sector else 0
    sector_data = sector_data.sort_values("total_value", ascending=False)

    all_with_currency = all_positions[all_positions["Currency"].notna()].copy()
    currency_data = all_with_currency.groupby("Currency").agg(
        total_value=("Mkt Value (CAD)", "sum"),
        num_positions=("Symbol", "count"),
    ).reset_index()
    total_cur = currency_data["total_value"].sum()
    currency_data["Weight"] = currency_data["total_value"] / total_cur if total_cur else 0
    currency_data = currency_data.sort_values("total_value", ascending=False)

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

    html = """<div id="page-exposure" class="page">
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

    html += """</tbody></table>
</div></div>
</div>"""
    return html


# ─── Single-Page Assembler ──────────────────────────────────────────────────

def _rrg_widget_html(prefix, *, height=600, with_legend=True, classes=""):
    """Return RRG widget markup for a given DOM id prefix."""
    legend_html = (
        f'<div id="{prefix}-legend" class="rrg-legend"></div>'
        if with_legend else ''
    )
    return f"""<div class="rrg-wrap {classes}">
  <div class="rrg-canvas-wrap">
    <canvas id="{prefix}-canvas" width="900" height="{height}"></canvas>
  </div>
  <div class="rrg-controls-row">
    <div class="rrg-controls">
      <button id="{prefix}-play" class="rrg-play-btn" type="button" aria-label="Play" title="Play">&#9654;</button>
      <input id="{prefix}-slider" class="rrg-slider" type="range" min="0" max="0" value="0" step="1">
      <div id="{prefix}-date" class="rrg-date">--</div>
    </div>
    <div class="rrg-controls rrg-trail-controls">
      <span class="rrg-trail-label" data-i18n="rrg_trail_label">Trail (days)</span>
      <input id="{prefix}-trail" class="rrg-slider rrg-trail-slider" type="range" min="1" max="126" value="63" step="1">
      <div id="{prefix}-trail-value" class="rrg-date">63</div>
    </div>
  </div>
  <div class="rrg-legend-row">
    {legend_html}
    <div class="rrg-legend-actions">
      <button id="{prefix}-show-all" type="button" class="rrg-mini-btn" data-i18n="rrg_show_all">Show All</button>
      <button id="{prefix}-hide-all" type="button" class="rrg-mini-btn" data-i18n="rrg_hide_all">Hide All</button>
    </div>
  </div>
</div>"""


def _generate_rotation_section(rrg_data):
    """Generate the Sector Rotation (RRG) tab content + JS.

    Returns (page_html, dashboard_widget_html, js).
    The JS exposes a global ``RRG_DATA`` and ``initRrg(prefix)`` so multiple
    widgets (e.g. the full tab + a smaller dashboard preview) can share data.
    """
    has_data = bool(rrg_data and rrg_data.get("dates") and rrg_data.get("series"))
    rrg_json = json.dumps(rrg_data if has_data else {
        "dates": [], "tickers": [], "series": {}, "trail_weeks": 12, "benchmark": "SPY"
    })

    no_data_html = (
        ''
        if has_data
        else '<p class="info" data-i18n="rrg_no_data">Insufficient price history to compute the rotation graph.</p>'
    )

    page_html = f"""<div id="page-rotation" class="page">
<h1 data-i18n="title_rotation">Relative Rotation Graph</h1>
<p class="info" data-i18n="info_rotation">Each dot is a holding plotted by its JdK RS-Ratio (x) and JdK RS-Momentum (y) versus SPY. Trail length is 3 months (~63 trading days). Press play or drag the slider to animate. Click a ticker in the legend to hide / show it.</p>
{no_data_html}
{_rrg_widget_html("rrg", height=600, with_legend=True)}
</div>"""

    dashboard_widget = ""

    js = f"""
window.RRG_DATA = {rrg_json};

window.initRrg = function(prefix) {{
    const RRG_DATA = window.RRG_DATA;
    const canvas = document.getElementById(prefix + '-canvas');
    if (!canvas || !RRG_DATA || !RRG_DATA.dates || RRG_DATA.dates.length === 0) return;
    const ctx = canvas.getContext('2d');
    const slider = document.getElementById(prefix + '-slider');
    const playBtn = document.getElementById(prefix + '-play');
    const dateLbl = document.getElementById(prefix + '-date');
    const legend = document.getElementById(prefix + '-legend');

    const PALETTE = [
        '#3A7BD5', '#D4A843', '#7A5BD5', '#00C49A', '#E07A3A',
        '#BD5BA8', '#3AB5D5', '#D5483A', '#9CC23A', '#E8A0BF',
        '#5BB7D5', '#A87A3A', '#D58FA8', '#7DC2A8', '#C28F3A',
        '#5BD5A8', '#D55B7A', '#7AD55B', '#3AD5BD', '#D53A7A'
    ];
    const colors = {{}};
    RRG_DATA.tickers.forEach(function(t, i) {{ colors[t] = PALETTE[i % PALETTE.length]; }});

    // Compute auto-zoomed bounds from actual data so the dots use the full grid.
    function computeBounds() {{
        let lo = Infinity, hi = -Infinity;
        RRG_DATA.tickers.forEach(function(t) {{
            const s = RRG_DATA.series[t];
            if (!s) return;
            for (let i = 0; i < s.length; i++) {{
                const a = s[i][0], b = s[i][1];
                if (a < lo) lo = a; if (a > hi) hi = a;
                if (b < lo) lo = b; if (b > hi) hi = b;
            }}
        }});
        if (!isFinite(lo) || !isFinite(hi)) {{ lo = 98; hi = 102; }}
        // Symmetric span around 100, with padding.
        const span = Math.max(Math.abs(hi - 100), Math.abs(100 - lo)) * 1.10;
        const minSpan = 0.6;
        const finalSpan = Math.max(span, minSpan);
        return {{ min: 100 - finalSpan, max: 100 + finalSpan }};
    }}
    const BOUNDS = computeBounds();
    const X_MIN = BOUNDS.min, X_MAX = BOUNDS.max;
    const Y_MIN = BOUNDS.min, Y_MAX = BOUNDS.max;
    const PAD = {{ left: 56, right: 24, top: 30, bottom: 44 }};

    // Per-ticker visibility (toggled via legend clicks).
    const visible = {{}};
    RRG_DATA.tickers.forEach(function(t) {{ visible[t] = true; }});
    let trailLen = RRG_DATA.trail_steps || RRG_DATA.trail_weeks || 63;

    function isLight() {{ return document.body.classList.contains('light-theme'); }}
    function getQuadrantLabel(key) {{
        var T = window.__T_RRG || {{}};
        return T[key] || key;
    }}
    function project(rsr, rsm, w, h) {{
        const px = PAD.left + (rsr - X_MIN) / (X_MAX - X_MIN) * (w - PAD.left - PAD.right);
        const py = PAD.top + (1 - (rsm - Y_MIN) / (Y_MAX - Y_MIN)) * (h - PAD.top - PAD.bottom);
        return [px, py];
    }}
    function hexToRgba(hex, a) {{
        const h = hex.replace('#', '');
        const r = parseInt(h.substring(0, 2), 16);
        const g = parseInt(h.substring(2, 4), 16);
        const b = parseInt(h.substring(4, 6), 16);
        return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
    }}

    // Generate "nice" axis tick values (5-7 ticks across the range).
    function niceTicks(lo, hi, target) {{
        target = target || 6;
        const range = hi - lo;
        const rough = range / target;
        const pow10 = Math.pow(10, Math.floor(Math.log10(rough)));
        const candidates = [1, 2, 2.5, 5, 10];
        let step = pow10 * candidates[0];
        for (let i = 0; i < candidates.length; i++) {{
            const s = pow10 * candidates[i];
            if (range / s <= target * 1.3) {{ step = s; break; }}
        }}
        const start = Math.ceil(lo / step) * step;
        const ticks = [];
        for (let v = start; v <= hi + step * 0.001; v += step) {{
            ticks.push(Number(v.toFixed(6)));
        }}
        return {{ step: step, ticks: ticks }};
    }}

    function draw(idx) {{
        const w = canvas.width, h = canvas.height;
        const light = isLight();
        ctx.clearRect(0, 0, w, h);
        ctx.fillStyle = light ? '#FFFFFF' : '#0F1729';
        ctx.fillRect(0, 0, w, h);

        const cx = PAD.left + (100 - X_MIN) / (X_MAX - X_MIN) * (w - PAD.left - PAD.right);
        const cy = PAD.top + (1 - (100 - Y_MIN) / (Y_MAX - Y_MIN)) * (h - PAD.top - PAD.bottom);
        const x0 = PAD.left, y0 = PAD.top;
        const x1 = w - PAD.right, y1 = h - PAD.bottom;

        // Quadrant tints.
        const qa = light ? 0.18 : 0.10;
        ctx.fillStyle = 'rgba(58, 123, 213, ' + qa + ')';
        ctx.fillRect(x0, y0, cx - x0, cy - y0);
        ctx.fillStyle = 'rgba(0, 196, 154, ' + qa + ')';
        ctx.fillRect(cx, y0, x1 - cx, cy - y0);
        ctx.fillStyle = 'rgba(216, 50, 50, ' + qa + ')';
        ctx.fillRect(x0, cy, cx - x0, y1 - cy);
        ctx.fillStyle = 'rgba(212, 168, 67, ' + qa + ')';
        ctx.fillRect(cx, cy, x1 - cx, y1 - cy);

        // Faint gridlines + tick marks.
        const gridColor  = light ? '#E1E6F0' : '#1F2D48';
        const axisColor  = light ? '#A0AEC0' : '#2A3F5F';
        const tickColor  = light ? '#6A7990' : '#7A8AA0';
        const xTicks = niceTicks(X_MIN, X_MAX, 6).ticks;
        const yTicks = niceTicks(Y_MIN, Y_MAX, 6).ticks;
        ctx.strokeStyle = gridColor; ctx.lineWidth = 1;
        ctx.beginPath();
        xTicks.forEach(function(v) {{
            const px = project(v, 100, w, h)[0];
            ctx.moveTo(px, y0); ctx.lineTo(px, y1);
        }});
        yTicks.forEach(function(v) {{
            const py = project(100, v, w, h)[1];
            ctx.moveTo(x0, py); ctx.lineTo(x1, py);
        }});
        ctx.stroke();

        // Outer frame + center crosshair.
        ctx.strokeStyle = axisColor; ctx.lineWidth = 1;
        ctx.strokeRect(x0, y0, x1 - x0, y1 - y0);
        ctx.beginPath();
        ctx.moveTo(cx, y0); ctx.lineTo(cx, y1);
        ctx.moveTo(x0, cy); ctx.lineTo(x1, cy);
        ctx.stroke();

        // Tick labels.
        ctx.fillStyle = tickColor;
        ctx.font = '10px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        xTicks.forEach(function(v) {{
            const px = project(v, 100, w, h)[0];
            ctx.fillText(v.toFixed(1), px, y1 + 14);
        }});
        ctx.textAlign = 'right';
        yTicks.forEach(function(v) {{
            const py = project(100, v, w, h)[1];
            ctx.fillText(v.toFixed(1), x0 - 6, py + 4);
        }});

        // Quadrant labels.
        ctx.fillStyle = light ? '#3A4A60' : '#8899AA';
        ctx.font = 'bold 12px Segoe UI, sans-serif';
        ctx.textAlign = 'right';
        ctx.fillText(getQuadrantLabel('rrg_improving'), cx - 6, y0 + 16);
        ctx.textAlign = 'left';
        ctx.fillText(getQuadrantLabel('rrg_leading'),   cx + 6, y0 + 16);
        ctx.textAlign = 'right';
        ctx.fillText(getQuadrantLabel('rrg_lagging'),   cx - 6, y1 - 8);
        ctx.textAlign = 'left';
        ctx.fillText(getQuadrantLabel('rrg_weakening'), cx + 6, y1 - 8);

        // Axis titles.
        ctx.fillStyle = light ? '#3A4A60' : '#AABBCC';
        ctx.font = '11px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(getQuadrantLabel('rrg_x_axis'), (x0 + x1) / 2, h - 8);
        ctx.save();
        ctx.translate(14, (y0 + y1) / 2);
        ctx.rotate(-Math.PI / 2);
        ctx.textAlign = 'center';
        ctx.fillText(getQuadrantLabel('rrg_y_axis'), 0, 0);
        ctx.restore();

        // Trails + dots. Trail line tapers (fat & opaque near dot, thin & faded at tail).
        const trail = trailLen;
        const labelColor = light ? '#1C2541' : '#E8ECF3';
        const dotStroke  = light ? '#FFFFFF' : '#0F1729';
        RRG_DATA.tickers.forEach(function(t) {{
            if (!visible[t]) return;
            const series = RRG_DATA.series[t];
            if (!series || idx >= series.length) return;
            const color = colors[t];
            // Draw trail segments newest -> oldest so newer (fatter, brighter) sit on top.
            for (let k = 1; k <= trail && idx - k >= 0; k++) {{
                const a = series[idx - k];
                const b = series[idx - k + 1];
                if (!a || !b) continue;
                const pa = project(a[0], a[1], w, h);
                const pb = project(b[0], b[1], w, h);
                const t01 = 1 - (k / (trail + 1));   // 1 near dot, 0 at tail
                const lw = 0.6 + 3.0 * t01;          // thicker near dot
                const alpha = 0.10 + 0.85 * t01;
                ctx.strokeStyle = hexToRgba(color, alpha);
                ctx.lineWidth = lw;
                ctx.lineCap = 'round';
                ctx.beginPath();
                ctx.moveTo(pa[0], pa[1]); ctx.lineTo(pb[0], pb[1]);
                ctx.stroke();
            }}
            const cur = series[idx];
            if (cur) {{
                const p = project(cur[0], cur[1], w, h);
                ctx.fillStyle = color;
                ctx.beginPath();
                ctx.arc(p[0], p[1], 7, 0, Math.PI * 2);
                ctx.fill();
                ctx.strokeStyle = dotStroke;
                ctx.lineWidth = 2;
                ctx.stroke();
                ctx.fillStyle = labelColor;
                ctx.font = 'bold 11px Segoe UI, sans-serif';
                ctx.textAlign = 'left';
                ctx.fillText(t, p[0] + 10, p[1] + 4);
            }}
        }});
    }}

    if (legend) {{
        legend.innerHTML = '';
        RRG_DATA.tickers.forEach(function(t) {{
            const item = document.createElement('span');
            item.className = 'rrg-legend-item';
            item.setAttribute('data-ticker', t);
            item.innerHTML = '<span class="rrg-legend-dot" style="background:' + colors[t] + '"></span>' + t;
            item.addEventListener('click', function() {{
                visible[t] = !visible[t];
                item.classList.toggle('rrg-legend-off', !visible[t]);
                draw(curIdx);
            }});
            legend.appendChild(item);
        }});
    }}

    slider.max = RRG_DATA.dates.length - 1;
    slider.value = RRG_DATA.dates.length - 1;
    let curIdx = RRG_DATA.dates.length - 1;
    let timer = null;

    function update(idx) {{
        curIdx = idx;
        slider.value = idx;
        if (dateLbl) dateLbl.textContent = RRG_DATA.dates[idx];
        draw(idx);
    }}
    slider.addEventListener('input', function() {{ update(parseInt(slider.value, 10)); }});

    const trailSlider = document.getElementById(prefix + '-trail');
    const trailValueLbl = document.getElementById(prefix + '-trail-value');
    if (trailSlider) {{
        trailSlider.value = trailLen;
        if (trailValueLbl) trailValueLbl.textContent = trailLen;
        trailSlider.addEventListener('input', function() {{
            trailLen = parseInt(trailSlider.value, 10) || 1;
            if (trailValueLbl) trailValueLbl.textContent = trailLen;
            draw(curIdx);
        }});
    }}

    const showAllBtn = document.getElementById(prefix + '-show-all');
    const hideAllBtn = document.getElementById(prefix + '-hide-all');
    function setAllVisible(v) {{
        RRG_DATA.tickers.forEach(function(t) {{ visible[t] = v; }});
        if (legend) {{
            const items = legend.querySelectorAll('.rrg-legend-item');
            items.forEach(function(el) {{ el.classList.toggle('rrg-legend-off', !v); }});
        }}
        draw(curIdx);
    }}
    if (showAllBtn) showAllBtn.addEventListener('click', function() {{ setAllVisible(true); }});
    if (hideAllBtn) hideAllBtn.addEventListener('click', function() {{ setAllVisible(false); }});

    function stop() {{
        if (timer) {{ clearInterval(timer); timer = null; }}
        playBtn.innerHTML = '&#9654;';
        playBtn.title = getQuadrantLabel('rrg_play');
        playBtn.setAttribute('aria-label', getQuadrantLabel('rrg_play'));
    }}
    function play() {{
        if (timer) return;
        if (curIdx >= RRG_DATA.dates.length - 1) {{
            // Restart shortly before the end so the trail is already visible.
            update(Math.max(0, RRG_DATA.dates.length - 1 - 90));
        }}
        playBtn.innerHTML = '&#10073;&#10073;';
        playBtn.title = getQuadrantLabel('rrg_pause');
        playBtn.setAttribute('aria-label', getQuadrantLabel('rrg_pause'));
        timer = setInterval(function() {{
            if (curIdx >= RRG_DATA.dates.length - 1) {{ stop(); return; }}
            update(curIdx + 1);
        }}, 60);  // ~16 fps with daily samples
    }}
    playBtn.addEventListener('click', function() {{
        if (timer) stop(); else play();
    }});

    // Register redraw hook (used by language + theme toggles).
    window.__rrgRedraws = window.__rrgRedraws || [];
    window.__rrgRedraws.push(function() {{ draw(curIdx); }});
    update(curIdx);
}};

document.addEventListener('DOMContentLoaded', function() {{
    if (!window.RRG_DATA || !window.RRG_DATA.dates || window.RRG_DATA.dates.length === 0) return;
    window.initRrg('rrg');
}});
"""
    return page_html, dashboard_widget, js


def generate_single_html(
    portfolio_value, metrics, num_positions, num_options,
    portfolio_df, opts_df, fund_df,
    option_delta_df, total_delta_usd,
    corr_matrix, individual_risk_df,
    stress_df, beta_val,
    rrg_data=None,
    usd_cad_rate=1.37,
):
    """Assemble all sections into a single HTML file with tab navigation."""
    # Generate all sections
    dashboard_html = _generate_dashboard_section(
        portfolio_value, metrics, num_positions, num_options, usd_cad_rate)
    positions_html = _generate_positions_section(
        portfolio_df, opts_df, fund_df, portfolio_value, usd_cad_rate)
    options_html = _generate_options_section(
        opts_df, option_delta_df, total_delta_usd, usd_cad_rate)
    correlation_html, correlation_js = _generate_correlation_section(corr_matrix)
    risk_html = _generate_risk_section(
        metrics, individual_risk_df, portfolio_value, usd_cad_rate)
    stress_html, stress_js = _generate_stress_section(
        stress_df, portfolio_value, beta_val, total_delta_usd,
        usd_cad_rate=usd_cad_rate, ann_return=metrics.get("Annualized Return", 0))
    exposure_html = _generate_exposure_section(
        portfolio_df, opts_df, portfolio_value, usd_cad_rate)
    rotation_html, _dashboard_rrg_unused, rotation_js = _generate_rotation_section(rrg_data)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Build nav HTML
    nav_links = []
    for page_id, label, i18n_key in TAB_PAGES:
        nav_links.append(
            f'<a data-page="{page_id}" onclick="showPage(\'{page_id}\')" data-i18n="{i18n_key}">{label}</a>')
    nav_html = '<div class="nav">' + ''.join(nav_links)
    nav_html += '<span class="spacer"></span>'
    nav_html += f'<span class="timestamp"><span data-i18n="generated">Generated:</span> {timestamp}</span>'
    nav_html += FLAG_BUTTONS_HTML
    nav_html += PRIVACY_SWITCH_HTML
    nav_html += THEME_TOGGLE_HTML
    nav_html += '</div>'

    # Translations JSON for JS
    translations_json = json.dumps(TRANSLATIONS, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Stock Portfolio Analytics Dashboard</title>
<style>
{COMBINED_CSS}
</style>
</head>
<body>
{nav_html}
{dashboard_html}
{positions_html}
{options_html}
{correlation_html}
{risk_html}
{stress_html}
{exposure_html}
{rotation_html}
<div id="cell-tooltip"></div>

<script>
// ─── Tab Navigation ─────────────────────────────────────────────────────────
function showPage(pageId) {{
    document.querySelectorAll('.page').forEach(function(p) {{ p.classList.remove('active'); }});
    document.querySelectorAll('.nav a[data-page]').forEach(function(a) {{ a.classList.remove('active'); }});
    var page = document.getElementById('page-' + pageId);
    if (page) page.classList.add('active');
    var link = document.querySelector('.nav a[data-page="' + pageId + '"]');
    if (link) link.classList.add('active');
    history.replaceState(null, '', '#' + pageId);
    window.scrollTo(0, 0);
}}
window.addEventListener('hashchange', function() {{
    var page = location.hash.replace('#', '') || 'dashboard';
    showPage(page);
}});
document.addEventListener('DOMContentLoaded', function() {{
    var page = location.hash.replace('#', '') || 'dashboard';
    showPage(page);
}});

// ─── Privacy Toggle (switch + icon: $ visible / $ slashed) ─────────────────
(function() {{
    var key = 'portfolio_privacy_mode';
    function isHidden() {{ return document.body.classList.contains('privacy-mode'); }}
    function syncUi() {{
        var sw = document.getElementById('privacy-switch');
        if (sw) sw.checked = isHidden();
        var icon = document.getElementById('privacy-icon');
        if (icon) icon.classList.toggle('hide-icon', isHidden());
    }}
    if (localStorage.getItem(key) === 'true') {{
        document.body.classList.add('privacy-mode');
    }}
    window.togglePrivacy = function() {{
        document.body.classList.toggle('privacy-mode');
        localStorage.setItem(key, isHidden());
        syncUi();
    }};
    document.addEventListener('DOMContentLoaded', syncUi);
}})();

// ─── Theme Toggle (dark / light) ────────────────────────────────────────────
(function() {{
    var key = 'portfolio_theme';
    function isLight() {{ return document.body.classList.contains('light-theme'); }}
    function syncUi() {{
        var btn = document.getElementById('theme-btn');
        if (btn) btn.innerHTML = isLight() ? '&#9728;' : '&#9790;'; // sun / moon
    }}
    if (localStorage.getItem(key) === 'light') {{
        document.body.classList.add('light-theme');
    }}
    window.toggleTheme = function() {{
        document.body.classList.toggle('light-theme');
        localStorage.setItem(key, isLight() ? 'light' : 'dark');
        syncUi();
        if (window.__rrgRedraws) {{
            window.__rrgRedraws.forEach(function(fn) {{ try {{ fn(); }} catch(_) {{}} }});
        }}
    }};
    document.addEventListener('DOMContentLoaded', syncUi);
}})();

// ─── Language Toggle (4 locales: en, zh-HK, zh-TW, zh-CN) ──────────────────
(function() {{
    var T = {translations_json};
    var langKey = 'portfolio_language';
    // Map UI lang code -> key inside the translation dict.
    var LANG_KEYS = {{
        'en': ['en'],
        'zh-HK': ['zh', 'zh-tw', 'zh-cn', 'en'],
        'zh-TW': ['zh-tw', 'zh', 'zh-cn', 'en'],
        'zh-CN': ['zh-cn', 'zh', 'zh-tw', 'en']
    }};
    function getLang() {{
        var v = localStorage.getItem(langKey) || 'en';
        // Migrate legacy 'zh' value to 'zh-HK'.
        if (v === 'zh') v = 'zh-HK';
        return v;
    }}
    function lookup(k, lang) {{
        var entry = T[k]; if (!entry) return null;
        var keys = LANG_KEYS[lang] || ['en'];
        for (var i = 0; i < keys.length; i++) {{
            if (entry[keys[i]]) return entry[keys[i]];
        }}
        return null;
    }}
    function buildRrgLabelMap(lang) {{
        var keys = ['rrg_leading','rrg_weakening','rrg_lagging','rrg_improving',
                    'rrg_x_axis','rrg_y_axis','rrg_play','rrg_pause',
                    'rrg_show_all','rrg_hide_all','rrg_trail_label'];
        var m = {{}};
        keys.forEach(function(k) {{ var v = lookup(k, lang); if (v) m[k] = v; }});
        return m;
    }}
    window.applyLanguage = function() {{
        var lang = getLang();
        document.querySelectorAll('[data-i18n]').forEach(function(el) {{
            var v = lookup(el.getAttribute('data-i18n'), lang);
            if (v != null) el.textContent = v;
        }});
        document.querySelectorAll('[data-i18n-html]').forEach(function(el) {{
            var v = lookup(el.getAttribute('data-i18n-html'), lang);
            if (v != null) el.innerHTML = v;
        }});
        document.querySelectorAll('.flag-btn').forEach(function(b) {{
            b.classList.toggle('active', b.getAttribute('data-lang') === lang);
        }});
        window.__T_RRG = buildRrgLabelMap(lang);
        if (window.__rrgRedraws) window.__rrgRedraws.forEach(function(fn) {{ try {{ fn(); }} catch(_) {{}} }});
    }};
    window.setLanguage = function(lang) {{
        if (!LANG_KEYS[lang]) return;
        localStorage.setItem(langKey, lang);
        applyLanguage();
    }};
    document.addEventListener('DOMContentLoaded', applyLanguage);
}})();

// ─── Sortable Tables ────────────────────────────────────────────────────────
{SORTABLE_JS}

// ─── Correlation Matrix JS ──────────────────────────────────────────────────
{correlation_js}

// ─── Stress Testing Currency Toggle ─────────────────────────────────────────
{stress_js}

// ─── Sector Rotation (RRG) ──────────────────────────────────────────────────
{rotation_js}
</script>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("Stock Portfolio Analytics Report Generator v3 (Single-Page)")
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

    # 3. Collect all tradeable tickers
    non_tradeable = {"Cash", "Short Cash"}
    stock_tickers = [t for t in portfolio_df["Symbol"].unique() if t not in non_tradeable]
    option_underlyings = opts_df["Symbol"].unique().tolist()
    extra_tickers = [t for t in option_underlyings if t not in stock_tickers and t not in non_tradeable]
    all_tickers = sorted(set(stock_tickers + extra_tickers))
    print(f"\n  Unique tradeable tickers (stocks+options): {len(all_tickers)}")

    # 4. Fetch fundamentals
    fund_df = fetch_fundamentals(all_tickers)
    sector_map = dict(zip(fund_df["Symbol"], fund_df["Sector"]))
    portfolio_df["Sector"] = portfolio_df["Symbol"].map(sector_map).fillna("")
    opts_df["Sector"] = opts_df["Symbol"].map(sector_map).fillna("")

    # 5. Fetch live option prices
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

    # 7b. Update portfolio positions with latest live prices from yfinance
    latest_prices = fetch_latest_prices(prices)
    if latest_prices:
        print(f"\nUpdating portfolio prices with latest live data ({len(latest_prices)} tickers)...")
        non_tradeable = {"Cash", "Short Cash"}
        updated_count = 0
        for idx, row in portfolio_df.iterrows():
            sym = row["Symbol"]
            if sym in non_tradeable:
                continue
            if sym in latest_prices:
                old_price = row["Price"]
                new_price = latest_prices[sym]
                currency = row.get("Currency", "USD")
                shares = row["Shares"]
                portfolio_df.at[idx, "Price"] = new_price
                new_mkt_value = shares * new_price
                portfolio_df.at[idx, "Mkt Value"] = new_mkt_value
                if currency == "CAD":
                    portfolio_df.at[idx, "Mkt Value (CAD)"] = new_mkt_value
                else:
                    portfolio_df.at[idx, "Mkt Value (CAD)"] = new_mkt_value * usd_cad_rate
                if abs(old_price - new_price) > 0.005:
                    print(f"    {sym}: ${old_price:,.2f} -> ${new_price:,.2f}")
                    updated_count += 1
        portfolio_value = portfolio_df["Mkt Value (CAD)"].sum()
        print(f"  Updated {updated_count} prices. New portfolio value (CAD): ${portfolio_value:,.0f}")
    else:
        print("\n  WARNING: No live prices available; using spreadsheet prices.")

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

    # 11. Risk metrics
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

    # 12. Stress testing
    beta_val = metrics.get("Beta to SPY", 1.0)
    if not isinstance(beta_val, (int, float)):
        beta_val = 1.0
    print("\nComputing stress testing scenarios...")
    stress_df = compute_stress_testing(
        portfolio_returns, weight_series, returns, portfolio_value,
        beta_val, total_delta_usd, usd_cad_rate=usd_cad_rate,
    )

    # 13. Individual risk
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

    # 13b. Sector rotation (RRG) data
    print("Computing relative rotation graph (RRG) data...")
    rrg_data = compute_rrg_data(prices, weight_series, benchmark="SPY")
    print(f"  RRG: {len(rrg_data.get('tickers', []))} tickers, "
          f"{len(rrg_data.get('dates', []))} weekly points")

    # 14. Generate single HTML report
    print("\nGenerating single-page HTML report...")
    html = generate_single_html(
        portfolio_value=portfolio_value,
        metrics=metrics,
        num_positions=len(portfolio_df),
        num_options=len(opts_df),
        portfolio_df=portfolio_df,
        opts_df=opts_df,
        fund_df=fund_df,
        option_delta_df=option_delta_df,
        total_delta_usd=total_delta_usd,
        corr_matrix=corr,
        individual_risk_df=individual_risk,
        stress_df=stress_df,
        beta_val=beta_val,
        rrg_data=rrg_data,
        usd_cad_rate=usd_cad_rate,
    )

    filepath = OUTPUT_DIR / "index.html"
    filepath.write_text(html, encoding="utf-8")
    size_kb = filepath.stat().st_size / 1024
    print(f"  Written: index.html ({size_kb:.1f} KB)")

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
    print("BUILD COMPLETE - 1 HTML report + 1 JSON file generated")
    print(f"Open index.html in a browser to view the dashboard.")
    print("=" * 60)


if __name__ == "__main__":
    main()
