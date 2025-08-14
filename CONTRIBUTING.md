# Contributing to Excel Trading Optimizer

Thank you for your interest in contributing to Excel Trading Optimizer! This document provides guidelines and information for contributors.

## 🤝 How to Contribute

### Reporting Issues

1. **Search existing issues** before creating a new one
2. **Use the issue template** and provide detailed information:
   - Python version and operating system
   - Full error traceback
   - Steps to reproduce the issue
   - Expected vs actual behavior
   - Excel configuration (if relevant)

### Suggesting Features

1. **Check existing feature requests** to avoid duplicates
2. **Describe the use case** and why it would be valuable
3. **Provide examples** of how the feature would be used
4. **Consider implementation complexity** and backwards compatibility

### Code Contributions

#### Getting Started

1. Fork the repository
2. Clone your fork locally:
   ```bash
   git clone https://github.com/yourusername/excel-trading-optimizer.git
   cd excel-trading-optimizer
   ```
3. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   pip install -r requirements-dev.txt  # If available
   ```

#### Development Workflow

1. **Create a feature branch**:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes**:
   - Follow the existing code style
   - Add tests for new functionality
   - Update documentation as needed

3. **Test your changes**:
   ```bash
   python -m pytest tests/  # If tests exist
   python main.py  # Test basic functionality
   ```

4. **Commit your changes**:
   ```bash
   git add .
   git commit -m "Add: brief description of your changes"
   ```

5. **Push and create a Pull Request**:
   ```bash
   git push origin feature/your-feature-name
   ```

## 📝 Code Style Guidelines

### Python Code Style

- Follow **PEP 8** style guidelines
- Use **meaningful variable and function names**
- Add **docstrings** to all functions and classes
- Keep functions **focused and small** (ideally < 50 lines)
- Use **type hints** where appropriate

### Example Function Format

```python
def calculate_performance_metrics(trades_df: pd.DataFrame, market_data: pd.DataFrame) -> Dict[str, float]:
    """
    Calculate comprehensive performance metrics for a trading strategy.
    
    Args:
        trades_df: DataFrame containing trade data with Entry, Exit, PnL columns
        market_data: DataFrame containing market price data
        
    Returns:
        Dictionary containing performance metrics (Sharpe ratio, returns, etc.)
        
    Raises:
        ValueError: If required columns are missing from input DataFrames
    """
    # Implementation here
    pass
```

### Excel Integration

- **Maintain backwards compatibility** with existing Excel templates
- **Document new Excel sheet requirements** in both code and README
- **Provide clear error messages** for Excel format issues
- **Test with different Excel versions** when possible

## 🧪 Testing Guidelines

### Manual Testing

Before submitting a PR, test the following scenarios:

1. **Basic backtest** with sample data
2. **Parameter optimization** with small parameter ranges
3. **Excel I/O operations** (reading/writing different sheets)
4. **Edge cases** (empty data, invalid parameters, etc.)

### Automated Testing (Future)

We plan to add automated testing. When contributing:

- **Write tests for new functions**
- **Test error handling and edge cases**
- **Mock external dependencies** (Excel files, market data)

## 📚 Documentation

### Code Documentation

- **Add docstrings** to all public functions
- **Comment complex algorithms** and business logic
- **Explain Excel sheet relationships** and data flow
- **Document parameter meanings** and valid ranges

### README Updates

Update the README when adding:
- New features or capabilities
- New Excel sheet requirements
- New dependencies
- Breaking changes

## 🏗️ Architecture Guidelines

### Modular Design

- **Keep modules focused** on single responsibilities
- **Minimize dependencies** between modules
- **Use clear interfaces** between components
- **Separate business logic** from I/O operations

### Current Module Structure

```
main.py              # Workflow orchestration
├── excel_io.py      # Excel file operations
├── strategy.py      # Trading logic evaluation
├── optimizer.py     # Parameter optimization
├── indicators.py    # Technical indicator calculations
├── metrics.py       # Performance calculations
└── visuals.py       # Chart generation
```

### Adding New Modules

When adding new modules:
1. **Follow the existing naming convention**
2. **Add clear module docstring** explaining purpose
3. **Update imports** in relevant files
4. **Document new dependencies** in requirements.txt

## 🐛 Debugging Guidelines

### Common Issues

1. **Excel File Corruption**: Always backup Excel files before modification
2. **Data Type Mismatches**: Check pandas DataFrame dtypes
3. **Missing Columns**: Validate Excel sheet structure
4. **Memory Issues**: Use chunking for large datasets

### Debugging Tools

- **Print debugging**: Use clear, prefixed debug messages
- **Pandas inspection**: `.info()`, `.describe()`, `.head()`
- **Excel validation**: Check sheet names, column headers
- **Performance profiling**: Use `time.time()` for bottlenecks

## 📋 Pull Request Checklist

Before submitting a PR, ensure:

- [ ] **Code follows style guidelines**
- [ ] **All functions have docstrings**
- [ ] **Manual testing completed**
- [ ] **README updated** (if applicable)
- [ ] **No breaking changes** (or clearly documented)
- [ ] **Excel template compatibility maintained**
- [ ] **Error handling implemented**
- [ ] **Performance impact considered**

## 🚀 Release Process

### Version Numbering

We follow semantic versioning (SemVer):
- **Major version** (X.0.0): Breaking changes
- **Minor version** (0.X.0): New features, backwards compatible
- **Patch version** (0.0.X): Bug fixes, backwards compatible

### Release Checklist

1. Update version numbers
2. Update CHANGELOG.md
3. Test with multiple Excel templates
4. Create release notes
5. Tag the release

## 💬 Community Guidelines

### Communication

- **Be respectful** and constructive in discussions
- **Provide context** when asking questions
- **Help others** when you can
- **Share your use cases** and experiences

### Code of Conduct

- **Be inclusive** and welcoming to all contributors
- **Focus on constructive feedback**
- **Respect different perspectives** and experience levels
- **Help create a positive learning environment**

## 📞 Getting Help

### Where to Ask Questions

1. **GitHub Issues**: For bugs and feature requests
2. **Discussions**: For general questions and ideas
3. **Pull Request comments**: For code-specific questions

### Providing Help

Include in your question:
- **Python version** and operating system
- **Complete error message** and traceback
- **Minimal example** that reproduces the issue
- **What you've already tried**

## 🎯 Priority Areas for Contribution

We especially welcome contributions in these areas:

1. **Testing framework** setup and test cases
2. **Performance optimization** for large datasets
3. **Additional technical indicators**
4. **Enhanced visualization options**
5. **Documentation improvements**
6. **Excel template enhancements**
7. **Error handling and validation**

Thank you for contributing to Excel Trading Optimizer! 🙏
