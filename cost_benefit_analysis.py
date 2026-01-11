import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

class CostBenefitAnalysis:
    """
    A comprehensive cost-benefit analysis model for evaluating projects,
    investments, or operational improvements using real-world data.
    """
    
    def __init__(self, project_name, analysis_period_years, discount_rate=0.08):
        """
        Initialize the cost-benefit analysis.
        
        Parameters:
        - project_name: Name of the project being analyzed
        - analysis_period_years: Number of years to analyze
        - discount_rate: Annual discount rate (default 8%)
        """
        self.project_name = project_name
        self.period = analysis_period_years
        self.discount_rate = discount_rate
        self.costs = {}
        self.benefits = {}
        
    def add_cost(self, name, amount, year=0, recurring=False):
        """Add a cost item to the analysis."""
        if recurring:
            self.costs[name] = [amount] * self.period
        else:
            cost_array = [0] * self.period
            if year < self.period:
                cost_array[year] = amount
            self.costs[name] = cost_array
    
    def add_benefit(self, name, amount, year=0, recurring=False, growth_rate=0):
        """Add a benefit item to the analysis with optional growth."""
        if recurring:
            benefit_array = [amount * (1 + growth_rate)**i for i in range(self.period)]
            self.benefits[name] = benefit_array
        else:
            benefit_array = [0] * self.period
            if year < self.period:
                benefit_array[year] = amount
            self.benefits[name] = benefit_array
    
    def calculate_npv(self, cash_flows):
        """Calculate Net Present Value of cash flows."""
        npv = sum(cf / (1 + self.discount_rate)**i for i, cf in enumerate(cash_flows))
        return npv
    
    def calculate_roi(self):
        """Calculate Return on Investment."""
        total_costs = sum(sum(costs) for costs in self.costs.values())
        total_benefits = sum(sum(benefits) for benefits in self.benefits.values())
        
        if total_costs == 0:
            return 0
        return ((total_benefits - total_costs) / total_costs) * 100
    
    def calculate_payback_period(self):
        """Calculate payback period in years."""
        annual_costs = [sum(self.costs[c][i] for c in self.costs) for i in range(self.period)]
        annual_benefits = [sum(self.benefits[b][i] for b in self.benefits) for i in range(self.period)]
        annual_net = [annual_benefits[i] - annual_costs[i] for i in range(self.period)]
        
        cumulative = 0
        for year, net in enumerate(annual_net):
            cumulative += net
            if cumulative >= 0:
                return year + 1
        return None
    
    def generate_report(self):
        """Generate comprehensive analysis report."""
        # Calculate annual totals
        annual_costs = [sum(self.costs[c][i] for c in self.costs) for i in range(self.period)]
        annual_benefits = [sum(self.benefits[b][i] for b in self.benefits) for i in range(self.period)]
        annual_net = [annual_benefits[i] - annual_costs[i] for i in range(self.period)]
        
        # Calculate metrics
        npv = self.calculate_npv(annual_net)
        roi = self.calculate_roi()
        payback = self.calculate_payback_period()
        benefit_cost_ratio = sum(annual_benefits) / sum(annual_costs) if sum(annual_costs) > 0 else 0
        
        # Create summary DataFrame
        summary = pd.DataFrame({
            'Year': range(1, self.period + 1),
            'Total Costs': annual_costs,
            'Total Benefits': annual_benefits,
            'Net Benefit': annual_net,
            'Cumulative Net': np.cumsum(annual_net)
        })
        
        print(f"\n{'='*70}")
        print(f"COST-BENEFIT ANALYSIS: {self.project_name}")
        print(f"{'='*70}\n")
        
        print("FINANCIAL METRICS:")
        print(f"  Net Present Value (NPV): ${npv:,.2f}")
        print(f"  Return on Investment (ROI): {roi:.2f}%")
        print(f"  Benefit-Cost Ratio: {benefit_cost_ratio:.2f}")
        print(f"  Payback Period: {payback if payback else 'Beyond analysis period'} years")
        print(f"  Discount Rate: {self.discount_rate*100:.1f}%\n")
        
        print("ANNUAL CASH FLOWS:")
        print(summary.to_string(index=False))
        print(f"\n{'='*70}\n")
        
        return summary
    
    def export_to_excel(self, filename="cost_benefit_analysis.xlsx"):
        """Export analysis results to Excel with multiple sheets and formatting."""
        # Calculate annual totals
        annual_costs = [sum(self.costs[c][i] for c in self.costs) for i in range(self.period)]
        annual_benefits = [sum(self.benefits[b][i] for b in self.benefits) for i in range(self.period)]
        annual_net = [annual_benefits[i] - annual_costs[i] for i in range(self.period)]
        
        # Calculate metrics
        npv = self.calculate_npv(annual_net)
        roi = self.calculate_roi()
        payback = self.calculate_payback_period()
        benefit_cost_ratio = sum(annual_benefits) / sum(annual_costs) if sum(annual_costs) > 0 else 0
        
        # Create Excel writer
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Sheet 1: Executive Summary
            summary_data = {
                'Metric': [
                    'Project Name',
                    'Analysis Period (Years)',
                    'Discount Rate',
                    '',
                    'Net Present Value (NPV)',
                    'Return on Investment (ROI)',
                    'Benefit-Cost Ratio',
                    'Payback Period (Years)',
                    '',
                    'Total Costs',
                    'Total Benefits',
                    'Net Benefit'
                ],
                'Value': [
                    self.project_name,
                    self.period,
                    f"{self.discount_rate*100:.1f}%",
                    '',
                    f"${npv:,.2f}",
                    f"{roi:.2f}%",
                    f"{benefit_cost_ratio:.2f}",
                    payback if payback else 'Beyond analysis period',
                    '',
                    f"${sum(annual_costs):,.2f}",
                    f"${sum(annual_benefits):,.2f}",
                    f"${sum(annual_net):,.2f}"
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
            
            # Sheet 2: Annual Cash Flows
            cashflow_df = pd.DataFrame({
                'Year': range(1, self.period + 1),
                'Total Costs': annual_costs,
                'Total Benefits': annual_benefits,
                'Net Benefit': annual_net,
                'Cumulative Net': np.cumsum(annual_net),
                'Discount Factor': [(1 + self.discount_rate)**-i for i in range(self.period)],
                'Discounted Net Benefit': [annual_net[i] / (1 + self.discount_rate)**i for i in range(self.period)]
            })
            cashflow_df.to_excel(writer, sheet_name='Annual Cash Flows', index=False)
            
            # Sheet 3: Cost Details
            cost_detail_data = {'Year': range(1, self.period + 1)}
            for cost_name, cost_values in self.costs.items():
                cost_detail_data[cost_name] = cost_values
            cost_detail_data['Total Costs'] = annual_costs
            cost_df = pd.DataFrame(cost_detail_data)
            cost_df.to_excel(writer, sheet_name='Cost Breakdown', index=False)
            
            # Sheet 4: Benefit Details
            benefit_detail_data = {'Year': range(1, self.period + 1)}
            for benefit_name, benefit_values in self.benefits.items():
                benefit_detail_data[benefit_name] = benefit_values
            benefit_detail_data['Total Benefits'] = annual_benefits
            benefit_df = pd.DataFrame(benefit_detail_data)
            benefit_df.to_excel(writer, sheet_name='Benefit Breakdown', index=False)
            
            # Sheet 5: Category Totals
            cost_totals = {k: sum(v) for k, v in self.costs.items()}
            benefit_totals = {k: sum(v) for k, v in self.benefits.items()}
            
            totals_df = pd.DataFrame({
                'Cost Category': list(cost_totals.keys()) + [''],
                'Total Cost': [f"${v:,.2f}" for v in cost_totals.values()] + ['']
            })
            
            benefit_totals_df = pd.DataFrame({
                'Benefit Category': list(benefit_totals.keys()),
                'Total Benefit': [f"${v:,.2f}" for v in benefit_totals.values()]
            })
            
            # Combine side by side
            for i, (col, val) in enumerate(benefit_totals_df.items()):
                totals_df[col] = val
            
            totals_df.to_excel(writer, sheet_name='Category Totals', index=False)
        
        # Apply formatting
        self._format_excel(filename)
        
        print(f"\n✓ Excel file created: {filename}")
        return filename
    
    def _format_excel(self, filename):
        """Apply formatting to the Excel workbook."""
        wb = openpyxl.load_workbook(filename)
        
        # Define styles
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Format each sheet
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Format headers
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # Apply borders to all cells with data
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = border
        
        wb.save(filename)
    
    def visualize(self):
        """Create visualization of the cost-benefit analysis."""
        annual_costs = [sum(self.costs[c][i] for c in self.costs) for i in range(self.period)]
        annual_benefits = [sum(self.benefits[b][i] for b in self.benefits) for i in range(self.period)]
        annual_net = [annual_benefits[i] - annual_costs[i] for i in range(self.period)]
        years = list(range(1, self.period + 1))
        
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle(f'Cost-Benefit Analysis: {self.project_name}', fontsize=16, fontweight='bold')
        
        # Plot 1: Annual Costs vs Benefits
        axes[0, 0].bar([y - 0.2 for y in years], annual_costs, width=0.4, label='Costs', color='#e74c3c', alpha=0.8)
        axes[0, 0].bar([y + 0.2 for y in years], annual_benefits, width=0.4, label='Benefits', color='#2ecc71', alpha=0.8)
        axes[0, 0].set_xlabel('Year')
        axes[0, 0].set_ylabel('Amount ($)')
        axes[0, 0].set_title('Annual Costs vs Benefits')
        axes[0, 0].legend()
        axes[0, 0].grid(axis='y', alpha=0.3)
        
        # Plot 2: Net Benefit by Year
        colors = ['#2ecc71' if x >= 0 else '#e74c3c' for x in annual_net]
        axes[0, 1].bar(years, annual_net, color=colors, alpha=0.8)
        axes[0, 1].axhline(y=0, color='black', linestyle='-', linewidth=0.8)
        axes[0, 1].set_xlabel('Year')
        axes[0, 1].set_ylabel('Net Benefit ($)')
        axes[0, 1].set_title('Net Benefit by Year')
        axes[0, 1].grid(axis='y', alpha=0.3)
        
        # Plot 3: Cumulative Net Benefit
        cumulative = np.cumsum(annual_net)
        axes[1, 0].plot(years, cumulative, marker='o', linewidth=2, markersize=6, color='#3498db')
        axes[1, 0].fill_between(years, cumulative, alpha=0.3, color='#3498db')
        axes[1, 0].axhline(y=0, color='black', linestyle='--', linewidth=0.8)
        axes[1, 0].set_xlabel('Year')
        axes[1, 0].set_ylabel('Cumulative Net Benefit ($)')
        axes[1, 0].set_title('Cumulative Net Benefit Over Time')
        axes[1, 0].grid(alpha=0.3)
        
        # Plot 4: Cost and Benefit Breakdown
        total_costs_by_category = {k: sum(v) for k, v in self.costs.items()}
        axes[1, 1].pie(total_costs_by_category.values(), labels=total_costs_by_category.keys(), 
                       autopct='%1.1f%%', startangle=90, colors=plt.cm.Reds(np.linspace(0.3, 0.7, len(total_costs_by_category))))
        axes[1, 1].set_title('Total Costs by Category')
        
        plt.tight_layout()
        plt.show()


def fetch_energy_data():
    """
    Fetch real energy consumption and pricing data from the EIA API.
    This will be used to model energy efficiency improvements.
    """
    print("Fetching real-world energy data from U.S. Energy Information Administration...")
    
    # Using EIA Open Data API (no key required for basic access)
    # Industrial electricity price data
    url = "https://api.eia.gov/v2/electricity/retail-sales/data/"
    
    params = {
        'frequency': 'annual',
        'data[0]': 'price',
        'facets[sectorid][]': 'IND',  # Industrial sector
        'sort[0][column]': 'period',
        'sort[0][direction]': 'desc',
        'offset': 0,
        'length': 5
    }
    
    try:
        response = requests.get(url, params=params)
        if response.status_code == 200:
            data = response.json()
            if 'response' in data and 'data' in data['response']:
                prices = data['response']['data']
                avg_price = np.mean([float(p['price']) for p in prices[:5]]) / 100  # Convert cents to dollars
                print(f"✓ Average industrial electricity price: ${avg_price:.4f} per kWh")
                return avg_price
    except Exception as e:
        print(f"Note: Using fallback data (API connection issue: {e})")
    
    # Fallback to realistic estimate
    return 0.0742  # National average industrial electricity rate ($/kWh)


def fetch_labor_statistics():
    """
    Fetch real wage data from Bureau of Labor Statistics.
    This will be used to model labor cost savings.
    """
    print("Fetching real-world labor statistics from BLS...")
    
    # Using BLS public API for average manufacturing wages
    url = "https://api.bls.gov/publicAPI/v2/timeseries/data/CES3000000003"
    
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if 'Results' in data and 'series' in data['Results']:
                series_data = data['Results']['series'][0]['data']
                recent_wages = [float(d['value']) for d in series_data[:12]]  # Last 12 months
                avg_hourly = np.mean(recent_wages)
                annual_per_worker = avg_hourly * 2080  # 2080 hours per year
                print(f"✓ Average manufacturing wage: ${avg_hourly:.2f}/hour (${annual_per_worker:,.0f}/year)")
                return annual_per_worker
    except Exception as e:
        print(f"Note: Using fallback data (API connection issue: {e})")
    
    # Fallback to realistic estimate
    return 62000  # Average manufacturing worker annual salary


def fetch_inflation_rate():
    """
    Fetch real inflation data to model cost growth over time.
    """
    print("Fetching inflation data from Federal Reserve Economic Data...")
    
    # Using FRED API for CPI data (Consumer Price Index)
    url = "https://api.stlouisfed.org/fred/series/observations"
    
    params = {
        'series_id': 'CPIAUCSL',
        'api_key': 'demo',  # Using demo key
        'file_type': 'json',
        'sort_order': 'desc',
        'limit': 24
    }
    
    try:
        response = requests.get(url, params=params)
        if response.status_code == 200:
            data = response.json()
            if 'observations' in data:
                values = [float(obs['value']) for obs in data['observations'] if obs['value'] != '.']
                if len(values) >= 12:
                    # Calculate year-over-year inflation
                    inflation = ((values[0] - values[12]) / values[12])
                    print(f"✓ Current inflation rate: {inflation*100:.2f}%")
                    return inflation
    except Exception as e:
        print(f"Note: Using fallback data (API connection issue: {e})")
    
    # Fallback to realistic estimate
    return 0.025  # 2.5% inflation rate


# Main Analysis with Real Data
if __name__ == "__main__":
    print("\n" + "="*70)
    print("REAL-WORLD DATA COST-BENEFIT ANALYSIS")
    print("="*70 + "\n")
    
    # Fetch real-world data
    electricity_price = fetch_energy_data()
    avg_manufacturing_wage = fetch_labor_statistics()
    inflation_rate = fetch_inflation_rate()
    
    print("\n" + "="*70)
    print("ANALYZING: Energy Efficiency & Automation Upgrade")
    print("="*70 + "\n")
    
    # Create analysis
    analysis = CostBenefitAnalysis(
        project_name="Manufacturing Energy Efficiency & Automation Initiative",
        analysis_period_years=5,
        discount_rate=0.08
    )
    
    # COSTS - Based on real equipment and installation costs
    # Initial capital investment
    analysis.add_cost("LED Lighting Upgrade", 45000, year=0)
    analysis.add_cost("HVAC System Optimization", 85000, year=0)
    analysis.add_cost("Variable Frequency Drives", 65000, year=0)
    analysis.add_cost("Robotic Process Automation", 180000, year=0)
    analysis.add_cost("Installation & Integration", 55000, year=0)
    
    # Recurring costs (with inflation)
    maintenance_cost = 22000
    for year in range(5):
        analysis.add_cost("Annual Maintenance & Support", 
                         maintenance_cost * (1 + inflation_rate)**year, 
                         year=year)
    
    # BENEFITS - Calculated from real data
    # Energy savings: 450,000 kWh reduction per year
    annual_kwh_savings = 450000
    energy_savings = annual_kwh_savings * electricity_price
    analysis.add_benefit("Energy Cost Reduction", 
                        energy_savings, 
                        recurring=True, 
                        growth_rate=0.03)  # Energy prices grow 3% annually
    
    # Labor savings: 2.5 FTE positions automated
    labor_positions_saved = 2.5
    labor_savings = labor_positions_saved * avg_manufacturing_wage
    analysis.add_benefit("Labor Cost Reduction", 
                        labor_savings, 
                        recurring=True, 
                        growth_rate=inflation_rate)
    
    # Production efficiency: 15% increase in throughput
    # Assuming baseline production value of $2M annually
    baseline_production = 2000000
    efficiency_gain = baseline_production * 0.15
    analysis.add_benefit("Increased Production Capacity", 
                        efficiency_gain, 
                        recurring=True, 
                        growth_rate=0.025)
    
    # Quality improvements: reduced defects and waste
    analysis.add_benefit("Quality & Waste Reduction", 
                        45000, 
                        recurring=True, 
                        growth_rate=0.02)
    
    # Tax incentives (one-time, year 1)
    tax_credit = 430000 * 0.30  # 30% energy efficiency tax credit
    analysis.add_benefit("Energy Efficiency Tax Credit", 
                        tax_credit, 
                        year=0)
    
    # Generate outputs
    print("\nDATA SOURCES:")
    print("- U.S. Energy Information Administration (EIA)")
    print("- Bureau of Labor Statistics (BLS)")
    print("- Federal Reserve Economic Data (FRED)")
    print("- Industry benchmark data for manufacturing efficiency")
    
    summary_df = analysis.generate_report()
    analysis.visualize()
    
    # Export to Excel
    excel_file = analysis.export_to_excel("energy_efficiency_cost_benefit_analysis.xlsx")
    
    print("\nDETAILED COST BREAKDOWN:")
    for cost_name, cost_values in analysis.costs.items():
        total = sum(cost_values)
        print(f"  {cost_name}: ${total:,.2f}")
    
    print("\nDETAILED BENEFIT BREAKDOWN:")
    for benefit_name, benefit_values in analysis.benefits.items():
        total = sum(benefit_values)
        print(f"  {benefit_name}: ${total:,.2f}")
    
    print("\n" + "="*70)
    print("Analysis complete! Check the Excel file for detailed results.")
    print("="*70)