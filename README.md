# Cost-Benefit Analysis Tool

A Python-based cost-benefit analysis model that uses real-world data from government APIs to evaluate projects and investments.

## Features

- Fetches real data from U.S. government APIs (EIA, BLS, FRED)
- Calculates NPV, ROI, payback period, and benefit-cost ratio
- Generates comprehensive Excel reports with multiple sheets
- Creates professional visualizations
- Supports custom scenarios and what-if analysis

## Data Sources

- U.S. Energy Information Administration (EIA): Industrial electricity prices
- Bureau of Labor Statistics (BLS): Manufacturing wage data
- Federal Reserve Economic Data (FRED): Inflation rates

## Installation

1. Clone the repository:
git clone https://github.com/YOUR_USERNAME/cost-benefit-analysis.git
cd cost-benefit-analysis

2. Install required packages:
pip install -r requirements.txt

## Usage

Run the analysis:
python cost_benefit_analysis.py

The script will fetch real-world data from government APIs, perform cost-benefit analysis, display results in console, generate visualizations, and export detailed Excel report.

## Example Output

The tool analyzes an Energy Efficiency and Automation Initiative with 5-year analysis period, real electricity pricing and wage data, NPV, ROI, and payback calculations, and comprehensive Excel report with 5 sheets.

## Requirements

- Python 3.7+
- pandas
- numpy
- matplotlib
- requests
- openpyxl

## License

MIT License
"@ | Out-File -FilePath README.md -Encoding UTF8
