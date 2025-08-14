import pandas as pd
import numpy as np


def calculate_performance_metrics(
    results_df: pd.DataFrame,
    market_data: pd.DataFrame,
    initial_cash: float = 10_000
) -> dict:
    if results_df.empty:
        return {
            "Total Return [%]": 0,
            "Sharpe Ratio": 0,
            "# Trades": 0,
            "Win Rate [%]": 0,
            "SqrtMSE": 0
        }

    # Use EntryDate and ExitDate for multi-period trades
    results_df["EntryDate"] = pd.to_datetime(results_df["EntryDate"])
    results_df["ExitDate"] = pd.to_datetime(results_df["ExitDate"])
    results_df.sort_values("EntryDate", inplace=True)
    results_df.reset_index(drop=True, inplace=True)

    # Build equity curve
    equity_curve = [initial_cash]
    for pnl in results_df["PnL"]:
        equity_curve.append(equity_curve[-1] + pnl)
    equity_curve = equity_curve[1:]
    results_df["Equity"] = equity_curve
    results_df["Returns"] = results_df["PnL"] / initial_cash
    returns = results_df["Returns"]

    # Drawdown Calculation
    equity_series = pd.Series(equity_curve, index=results_df["ExitDate"])
    running_max = equity_series.cummax()
    drawdown = equity_series / running_max - 1.0
    max_drawdown = drawdown.min() * 100

    # Other Metrics
    total_return = (equity_curve[-1] - initial_cash) / initial_cash * 100
    trades = len(results_df)
    wins = sum(results_df["PnL"] > 0)
    win_rate = wins / trades * 100 if trades > 0 else 0
    avg_trade = returns.mean()
    volatility = returns.std()
    sharpe_ratio = avg_trade / volatility * (252 ** 0.5) if volatility > 0 else 0
    duration = (results_df["ExitDate"].iloc[-1] - results_df["EntryDate"].iloc[0]) + pd.Timedelta(days=1)

    # SqrtMSE between Pt and Close from full market_data
    if "Pt" in market_data.columns and "Close" in market_data.columns:
        mse = np.mean((market_data["Pt"] - market_data["Close"]) ** 2)
        sqrt_mse = np.sqrt(mse)
    else:
        sqrt_mse = 0.0

    metrics = {
        "Start": results_df["EntryDate"].iloc[0],
        "End": results_df["ExitDate"].iloc[-1],
        "Duration": duration,
        "Equity Final [$]": round(equity_curve[-1], 2),
        "Return [%]": round(total_return, 4),
        "# Trades": trades,
        "Win Rate [%]": round(win_rate, 2),
        "Avg. Trade [%]": round(avg_trade * 100, 4),
        "Sharpe Ratio": round(sharpe_ratio, 4),
        "Max Drawdown [%]": round(max_drawdown, 4),
        "SqrtMSE": round(sqrt_mse, 4),
    }
    return metrics
 


def compute_optimization_metrics(metrics: dict) -> dict:
    """
    Map metrics dict to optimization keys for scoring.
    """
    return {
        "AccReturn": metrics.get("Return [%]", 0),
        "Sharpe": metrics.get("Sharpe Ratio", 0),
        "Max Drawdown": metrics.get("Max Drawdown [%]", 0),
        "Accuracy": metrics.get("Win Rate [%]", 0),
        "SqrtMSE": metrics.get("SqrtMSE", 0),
    }


