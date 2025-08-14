"""
Setup script for Excel Trading Optimizer

A comprehensive Python-based trading strategy backtester and optimizer
that uses Excel for configuration and provides robust performance analysis.
"""

from setuptools import setup, find_packages
import os

# Read README file
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

# Read requirements
with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="excel-trading-optimizer",
    version="1.0.0",
    author="Excel Trading Optimizer Project",
    author_email="your.email@example.com",  # Update with your email
    description="Excel-driven trading strategy backtester and optimizer",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/excel-trading-optimizer",  # Update with your repo
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Financial and Insurance Industry",
        "Intended Audience :: Developers",
        "Topic :: Office/Business :: Financial :: Investment",
        "Topic :: Scientific/Engineering :: Information Analysis",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
            "black>=21.0",
            "flake8>=3.8",
            "mypy>=0.800",
        ],
        "docs": [
            "sphinx>=4.0",
            "sphinx-rtd-theme>=0.5",
        ],
    },
    entry_points={
        "console_scripts": [
            "excel-trading-optimizer=main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["excel/*.xlsx", "*.md", "*.txt"],
    },
    keywords=[
        "trading",
        "backtesting",
        "optimization",
        "excel",
        "technical-analysis",
        "finance",
        "quantitative",
        "strategy",
        "hyperopt",
        "ta-lib",
    ],
    project_urls={
        "Bug Reports": "https://github.com/yourusername/excel-trading-optimizer/issues",
        "Source": "https://github.com/yourusername/excel-trading-optimizer",
        "Documentation": "https://github.com/yourusername/excel-trading-optimizer#readme",
    },
)
