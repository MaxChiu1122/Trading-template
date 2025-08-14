# Excel Trading Optimizer

A comprehensive Python-based trading strategy backtester and optimizer that uses Excel for configuration and provides robust performance analysis with visualization capabilities.

## 🚀 Features

- **Excel-Driven Configuration**: Define strategies, indicators, and parameters directly in Excel
- **Dynamic Indicator Building**: Support for TA-Lib indicators and custom arithmetic combinations
- **Rolling Window Optimization**: Hyperopt-powered parameter optimization with train/test splits
- **Multi-Period Position Holding**: Support for both single-day and multi-day position management
- **Comprehensive Performance Metrics**: 15+ performance metrics including Sharpe ratio, max drawdown, win rate
- **Interactive Visualizations**: Equity curves, trade markers, and performance charts
- **Modular Architecture**: Clean, maintainable code with separated concerns

## 📊 Sample Results

The system generates comprehensive trading analysis including:
- Performance metrics (returns, Sharpe ratio, drawdowns)
- Trade-by-trade analysis with entry/exit tracking
- Visual equity curves and trade markers
- Parameter optimization results across rolling windows

## 🛠️ Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excel-trading-optimizer.git
cd excel-trading-optimizer
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Ensure you have TA-Lib installed. For Windows:
```bash
# Download the appropriate .whl file from https://www.lfd.uci.edu/~gohlke/pythonlibs/#ta-lib
pip install TA_Lib-0.4.24-cp39-cp39-win_amd64.whl  # adjust for your Python version
```

## 📈 Quick Start

### Basic Backtesting

1. Open `excel/trading_template.xlsx`
2. Configure your strategy in the "Strategy Logic Builder" sheet
3. Set up indicators in the "Indicator Builder" sheet
4. Run the backtest:

```python
from main import main
main(optimize=False)  # Run basic backtest
```

### Parameter Optimization

1. Define parameter ranges in the "Dashboard" sheet
2. Set optimization objectives and weights
3. Run optimization:

```python
from main import main
main(optimize=True)  # Run optimization
```

## 📁 Project Structure

```
excel-trading-optimizer/
├── main.py                    # Main entry point and workflow orchestration
├── strategy.py               # Strategy logic evaluation and trade generation
├── optimizer.py              # Rolling window optimization with Hyperopt
├── indicator_builder.py      # Dynamic indicator construction
├── performance_metrics.py    # Performance calculation and analysis
├── excel_io.py              # Excel file I/O operations
├── generate_visuals.py       # Visualization and plotting
├── update_data.py           # Market data updating utilities
├── requirements.txt         # Python dependencies
├── excel/
│   └── trading_template.xlsx # Excel configuration template
└── images/                  # Generated visualization outputs
```

## 🔧 Configuration

### Strategy Logic Builder

Define your trading rules in Excel using this format:

| Rule Type | Column A | Operator | Column B/Value | Action at | Logic Type |
|-----------|----------|----------|----------------|-----------|------------|
| Enter-Buy | RSI | < | 30 | Open | END |
| Exit-long | RSI | > | 70 | Close | END |

### Indicator Builder

Create custom indicators with parameter support:

| Name | Function | Param1 | Param2 | Combination |
|------|----------|--------|--------|-------------|
| RSI_14 | RSI | Close | 14 | |
| MA_Fast | SMA | Close | fast_period | |
| Signal | | | | RSI_14 + MA_Fast |

### Parameter Optimization

Set parameter ranges for optimization:

| Parameter | Min | Max | Step |
|-----------|-----|-----|------|
| fast_period | 5 | 20 | 1 |
| slow_period | 20 | 50 | 5 |

## 📊 Performance Metrics

The system calculates comprehensive performance metrics:

- **Returns**: Total and annualized returns
- **Risk Metrics**: Sharpe ratio, Sortino ratio, max drawdown
- **Trade Statistics**: Win rate, profit factor, average trade
- **Risk Management**: VaR, CVaR, volatility measures

## 🎯 Strategy Types Supported

- **Trend Following**: Moving average crossovers, momentum strategies
- **Mean Reversion**: RSI, Bollinger Bands, oversold/overbought
- **Breakout**: Price level breaks, volatility breakouts
- **Multi-Factor**: Complex combinations of technical indicators

## 🔄 Optimization Features

- **Rolling Window Backtesting**: Time-series cross-validation
- **Hyperparameter Optimization**: Bayesian optimization with Hyperopt
- **Multiple Objectives**: Multi-objective optimization with custom weights
- **Robust Validation**: Separate train/test periods for each window

## 📝 Example Usage

```python
# Load configuration from Excel
from excel_io import read_dashboard_inputs

config = read_dashboard_inputs("excel/trading_template.xlsx")

# Run strategy backtest
from strategy import parse_strategy_logic, strategy_from_logic

rules = parse_strategy_logic(config["logic_table"])
trades = strategy_from_logic(config["market_data"], rules)

# Calculate performance metrics
from performance_metrics import calculate_performance_metrics

metrics = calculate_performance_metrics(trades, config["market_data"])
print(f"Total Return: {metrics['Return [%]']:.2f}%")
print(f"Sharpe Ratio: {metrics['Sharpe Ratio']:.2f}")
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This software is for educational and research purposes only. Past performance does not guarantee future results. Always conduct thorough testing before using any trading strategy with real money.

## 🙋‍♂️ Support

If you encounter any issues or have questions:
1. Check the [Issues](https://github.com/yourusername/excel-trading-optimizer/issues) page
2. Create a new issue with detailed description
3. Include your Excel configuration and error messages

## 🔗 Related Projects

- [TA-Lib](https://github.com/mrjbq7/ta-lib) - Technical Analysis Library
- [Hyperopt](https://github.com/hyperopt/hyperopt) - Hyperparameter Optimization
- [Pandas](https://github.com/pandas-dev/pandas) - Data Analysis Library

---

**Built with ❤️ for algorithmic trading enthusiasts**
