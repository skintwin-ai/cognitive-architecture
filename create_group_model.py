import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Color
from openpyxl.utils import get_column_letter
import os

# Define file path
excel_file = '/home/ubuntu/skin_care_financial_model.xlsx'

# Load the existing workbook
try:
    workbook = openpyxl.load_workbook(excel_file)
except FileNotFoundError:
    print(f"Error: File {excel_file} not found.")
    exit(1)

# Create Group Model sheet
if 'Group Model' in workbook.sheetnames:
    idx = workbook.sheetnames.index('Group Model')
    workbook.remove(workbook.worksheets[idx])

group_model = workbook.create_sheet('Group Model', 1)  # Add after Dashboard

# Add title
group_model['A1'] = 'Skin Care Supply Chain Group Financial Model'
group_model['A1'].font = Font(bold=True, size=16)
group_model.merge_cells('A1:J1')
group_model['A1'].alignment = Alignment(horizontal='center')

# Define styles
title_font = Font(bold=True, size=14)
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
subheader_font = Font(bold=True)
subheader_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
calculation_fill = PatternFill(start_color="E6F1DC", end_color="E6F1DC", fill_type="solid")
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Add section headers
group_model['A3'] = 'Group Financial Summary'
group_model['A3'].font = Font(bold=True, size=14)
group_model.merge_cells('A3:J3')
group_model['A3'].alignment = Alignment(horizontal='center')

# Add column headers
headers = ['Entity/Metric', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5']
for i, header in enumerate(headers):
    cell = group_model.cell(row=4, column=i+1)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

# Define entities and their metrics
entities = [
    'Manufacturing Entity',
    'Distribution Entity',
    'Retail Entity',
    'Corporate Entity',
    'Eliminations',
    'Group Consolidated'
]

metrics = [
    'Revenue',
    'Cost of Goods Sold (COGS)',
    'Gross Profit',
    'Gross Margin %',
    'Operational Costs',
    'SG&A Expenses',
    'EBITDA',
    'EBITDA Margin %',
    'Depreciation & Amortization',
    'EBIT',
    'Interest Expense',
    'EBT',
    'Income Tax',
    'Net Income',
    'Net Profit Margin %'
]

# Add entity sections
row = 5
for entity in entities:
    # Add entity header
    entity_row = row
    cell = group_model.cell(row=entity_row, column=1)
    cell.value = entity
    cell.font = subheader_font
    cell.fill = subheader_fill
    cell.alignment = Alignment(horizontal='left', vertical='center')
    cell.border = thin_border
    
    for col in range(2, 7):  # Columns B through F
        cell = group_model.cell(row=entity_row, column=col)
        cell.border = thin_border
        cell.fill = subheader_fill
    
    row += 1
    
    # Add metrics for this entity
    for metric in metrics:
        metric_row = row
        cell = group_model.cell(row=metric_row, column=1)
        cell.value = f"  {metric}"  # Indent for hierarchy
        cell.alignment = Alignment(horizontal='left', vertical='center')
        cell.border = thin_border
        
        # Add formulas for each year
        for year_col in range(2, 7):  # Columns B through F
            cell = group_model.cell(row=metric_row, column=year_col)
            
            # Different formulas based on entity and metric
            if entity == 'Manufacturing Entity':
                if metric == 'Revenue':
                    # 40% of total revenue
                    cell.value = f'=Revenue!{get_column_letter(year_col+4)}13*0.4'
                elif metric == 'Cost of Goods Sold (COGS)':
                    # 80% of manufacturing revenue
                    cell.value = f'=B{metric_row}*0.8'
                elif metric == 'Gross Profit':
                    # Revenue - COGS
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Gross Margin %':
                    # Gross Profit / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-3}'
                    cell.number_format = '0.0%'
                elif metric == 'Operational Costs':
                    # 10% of manufacturing revenue
                    cell.value = f'=B{metric_row-4}*0.1'
                elif metric == 'SG&A Expenses':
                    # 5% of manufacturing revenue
                    cell.value = f'=B{metric_row-5}*0.05'
                elif metric == 'EBITDA':
                    # Gross Profit - Operational Costs - SG&A
                    cell.value = f'=B{metric_row-4}-B{metric_row-2}-B{metric_row-1}'
                elif metric == 'EBITDA Margin %':
                    # EBITDA / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-7}'
                    cell.number_format = '0.0%'
                elif metric == 'Depreciation & Amortization':
                    # Fixed amount
                    cell.value = '=50000'
                elif metric == 'EBIT':
                    # EBITDA - D&A
                    cell.value = f'=B{metric_row-3}-B{metric_row-1}'
                elif metric == 'Interest Expense':
                    # Fixed amount
                    cell.value = '=20000'
                elif metric == 'EBT':
                    # EBIT - Interest
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Income Tax':
                    # EBT * Tax Rate (25%)
                    cell.value = f'=B{metric_row-1}*0.25'
                elif metric == 'Net Income':
                    # EBT - Tax
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Net Profit Margin %':
                    # Net Income / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-14}'
                    cell.number_format = '0.0%'
            
            elif entity == 'Distribution Entity':
                if metric == 'Revenue':
                    # 30% of total revenue
                    cell.value = f'=Revenue!{get_column_letter(year_col+4)}13*0.3'
                elif metric == 'Cost of Goods Sold (COGS)':
                    # 70% of distribution revenue
                    cell.value = f'=B{metric_row}*0.7'
                elif metric == 'Gross Profit':
                    # Revenue - COGS
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Gross Margin %':
                    # Gross Profit / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-3}'
                    cell.number_format = '0.0%'
                elif metric == 'Operational Costs':
                    # 15% of distribution revenue
                    cell.value = f'=B{metric_row-4}*0.15'
                elif metric == 'SG&A Expenses':
                    # 8% of distribution revenue
                    cell.value = f'=B{metric_row-5}*0.08'
                elif metric == 'EBITDA':
                    # Gross Profit - Operational Costs - SG&A
                    cell.value = f'=B{metric_row-4}-B{metric_row-2}-B{metric_row-1}'
                elif metric == 'EBITDA Margin %':
                    # EBITDA / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-7}'
                    cell.number_format = '0.0%'
                elif metric == 'Depreciation & Amortization':
                    # Fixed amount
                    cell.value = '=30000'
                elif metric == 'EBIT':
                    # EBITDA - D&A
                    cell.value = f'=B{metric_row-3}-B{metric_row-1}'
                elif metric == 'Interest Expense':
                    # Fixed amount
                    cell.value = '=15000'
                elif metric == 'EBT':
                    # EBIT - Interest
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Income Tax':
                    # EBT * Tax Rate (25%)
                    cell.value = f'=B{metric_row-1}*0.25'
                elif metric == 'Net Income':
                    # EBT - Tax
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Net Profit Margin %':
                    # Net Income / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-14}'
                    cell.number_format = '0.0%'
            
            elif entity == 'Retail Entity':
                if metric == 'Revenue':
                    # 50% of total revenue
                    cell.value = f'=Revenue!{get_column_letter(year_col+4)}13*0.5'
                elif metric == 'Cost of Goods Sold (COGS)':
                    # 60% of retail revenue
                    cell.value = f'=B{metric_row}*0.6'
                elif metric == 'Gross Profit':
                    # Revenue - COGS
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Gross Margin %':
                    # Gross Profit / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-3}'
                    cell.number_format = '0.0%'
                elif metric == 'Operational Costs':
                    # 20% of retail revenue
                    cell.value = f'=B{metric_row-4}*0.2'
                elif metric == 'SG&A Expenses':
                    # 12% of retail revenue
                    cell.value = f'=B{metric_row-5}*0.12'
                elif metric == 'EBITDA':
                    # Gross Profit - Operational Costs - SG&A
                    cell.value = f'=B{metric_row-4}-B{metric_row-2}-B{metric_row-1}'
                elif metric == 'EBITDA Margin %':
                    # EBITDA / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-7}'
                    cell.number_format = '0.0%'
                elif metric == 'Depreciation & Amortization':
                    # Fixed amount
                    cell.value = '=40000'
                elif metric == 'EBIT':
                    # EBITDA - D&A
                    cell.value = f'=B{metric_row-3}-B{metric_row-1}'
                elif metric == 'Interest Expense':
                    # Fixed amount
                    cell.value = '=25000'
                elif metric == 'EBT':
                    # EBIT - Interest
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Income Tax':
                    # EBT * Tax Rate (25%)
                    cell.value = f'=B{metric_row-1}*0.25'
                elif metric == 'Net Income':
                    # EBT - Tax
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Net Profit Margin %':
                    # Net Income / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-14}'
                    cell.number_format = '0.0%'
            
            elif entity == 'Corporate Entity':
                if metric == 'Revenue':
                    # Corporate revenue (management fees, etc.)
                    cell.value = '=200000'
                elif metric == 'Cost of Goods Sold (COGS)':
                    # Minimal COGS for corporate
                    cell.value = '=50000'
                elif metric == 'Gross Profit':
                    # Revenue - COGS
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Gross Margin %':
                    # Gross Profit / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-3}'
                    cell.number_format = '0.0%'
                elif metric == 'Operational Costs':
                    # Fixed operational costs
                    cell.value = '=300000'
                elif metric == 'SG&A Expenses':
                    # Fixed SG&A expenses
                    cell.value = '=400000'
                elif metric == 'EBITDA':
                    # Gross Profit - Operational Costs - SG&A
                    cell.value = f'=B{metric_row-4}-B{metric_row-2}-B{metric_row-1}'
                elif metric == 'EBITDA Margin %':
                    # EBITDA / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-7}'
                    cell.number_format = '0.0%'
                elif metric == 'Depreciation & Amortization':
                    # Fixed amount
                    cell.value = '=20000'
                elif metric == 'EBIT':
                    # EBITDA - D&A
                    cell.value = f'=B{metric_row-3}-B{metric_row-1}'
                elif metric == 'Interest Expense':
                    # Fixed amount
                    cell.value = '=10000'
                elif metric == 'EBT':
                    # EBIT - Interest
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Income Tax':
                    # EBT * Tax Rate (25%)
                    cell.value = f'=B{metric_row-1}*0.25'
                elif metric == 'Net Income':
                    # EBT - Tax
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Net Profit Margin %':
                    # Net Income / Revenue
                    cell.value = f'=B{metric_row-1}/B{metric_row-14}'
                    cell.number_format = '0.0%'
            
            elif entity == 'Eliminations':
                if metric == 'Revenue':
                    # Intercompany eliminations (20% of total revenue)
                    cell.value = f'=-Revenue!{get_column_letter(year_col+4)}13*0.2'
                elif metric == 'Cost of Goods Sold (COGS)':
                    # Corresponding COGS eliminations
                    cell.value = f'=-B{metric_row}*0.8'
                elif metric == 'Gross Profit':
                    # Revenue - COGS
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Gross Margin %':
                    # N/A for eliminations
                    cell.value = 'N/A'
                elif metric == 'Operational Costs':
                    # No operational cost eliminations
                    cell.value = '=0'
                elif metric == 'SG&A Expenses':
                    # No SG&A eliminations
                    cell.value = '=0'
                elif metric == 'EBITDA':
                    # Gross Profit - Operational Costs - SG&A
                    cell.value = f'=B{metric_row-4}-B{metric_row-2}-B{metric_row-1}'
                elif metric == 'EBITDA Margin %':
                    # N/A for eliminations
                    cell.value = 'N/A'
                elif metric == 'Depreciation & Amortization':
                    # No D&A eliminations
                    cell.value = '=0'
                elif metric == 'EBIT':
                    # EBITDA - D&A
                    cell.value = f'=B{metric_row-3}-B{metric_row-1}'
                elif metric == 'Interest Expense':
                    # No interest eliminations
                    cell.value = '=0'
                elif metric == 'EBT':
                    # EBIT - Interest
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Income Tax':
                    # No tax eliminations
                    cell.value = '=0'
                elif metric == 'Net Income':
                    # EBT - Tax
                    cell.value = f'=B{metric_row-2}-B{metric_row-1}'
                elif metric == 'Net Profit Margin %':
                    # N/A for eliminations
                    cell.value = 'N/A'
            
            elif entity == 'Group Consolidated':
                # Find the row numbers for each entity's metrics
                manufacturing_start = 6
                distribution_start = manufacturing_start + len(metrics) + 1
                retail_start = distribution_start + len(metrics) + 1
                corporate_start = retail_start + len(metrics) + 1
                eliminations_start = corporate_start + len(metrics) + 1
                
                # Calculate the row offset for the current metric
                metric_offset = metrics.index(metric)
                
                if metric in ['Gross Margin %', 'EBITDA Margin %', 'Net Profit Margin %']:
                    # For percentage metrics, calculate directly
                    if metric == 'Gross Margin %':
                        # Gross Profit / Revenue
                        cell.value = f'=B{metric_row-1}/B{metric_row-3}'
                        cell.number_format = '0.0%'
                    elif metric == 'EBITDA Margin %':
                        # EBITDA / Revenue
                        cell.value = f'=B{metric_row-1}/B{metric_row-7}'
                        cell.number_format = '0.0%'
                    elif metric == 'Net Profit Margin %':
                        # Net Income / Revenue
                        cell.value = f'=B{metric_row-1}/B{metric_row-14}'
                        cell.number_format = '0.0%'
                else:
                    # For all other metrics, sum across entities
                    cell.value = f'=B{manufacturing_start + metric_offset}+B{distribution_start + metric_offset}+B{retail_start + metric_offset}+B{corporate_start + metric_offset}+B{eliminations_start + metric_offset}'
            
            # Format cells
            if metric in ['Gross Margin %', 'EBITDA Margin %', 'Net Profit Margin %']:
                if cell.value != 'N/A':
                    cell.number_format = '0.0%'
            else:
                cell.number_format = '#,##0'
            
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right')
            
            # Apply the same formula pattern to other years (columns C through F)
            if year_col > 2:  # For years 2-5
                # Replace the column reference in the formula
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    formula = cell.value
                    # Replace B with the current column letter
                    formula = formula.replace('B', get_column_letter(year_col))
                    # If the formula references Revenue sheet, adjust for the year
                    if 'Revenue!' in formula:
                        for i in range(5):
                            old_col = get_column_letter(i+6)  # F, G, H, I, J
                            new_col = get_column_letter(i+6)  # F, G, H, I, J
                            if old_col in formula:
                                formula = formula.replace(f'Revenue!{old_col}', f'Revenue!{new_col}')
                    cell.value = formula
        
        row += 1
    
    # Add a blank row between entities
    row += 1

# Add a section for Key Performance Indicators (KPIs)
kpi_row = row + 2
group_model.cell(row=kpi_row, column=1).value = 'Key Performance Indicators (KPIs)'
group_model.cell(row=kpi_row, column=1).font = Font(bold=True, size=14)
group_model.merge_cells(f'A{kpi_row}:F{kpi_row}')
group_model.cell(row=kpi_row, column=1).alignment = Alignment(horizontal='center')

# Add KPI headers
kpi_row += 1
headers = ['KPI', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5']
for i, header in enumerate(headers):
    cell = group_model.cell(row=kpi_row, column=i+1)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

# Define KPIs
kpis = [
    'Return on Investment (ROI)',
    'Return on Assets (ROA)',
    'Return on Equity (ROE)',
    'Inventory Turnover',
    'Days Sales Outstanding (DSO)',
    'Days Payable Outstanding (DPO)',
    'Cash Conversion Cycle',
    'Operating Cash Flow',
    'Free Cash Flow',
    'Debt to Equity Ratio'
]

# Find the row for Group Consolidated Net Income
group_net_income_row = 0
for row_idx in range(5, 100):  # Search through a reasonable range
    cell_value = group_model.cell(row=row_idx, column=1).value
    if cell_value == 'Group Consolidated':
        group_net_income_row = row_idx + metrics.index('Net Income')
        break

# Add KPI rows
for i, kpi in enumerate(kpis):
    current_row = kpi_row + i + 1
    cell = group_model.cell(row=current_row, column=1)
    cell.value = kpi
    cell.alignment = Alignment(horizontal='left', vertical='center')
    cell.border = thin_border
    
    for year_col in range(2, 7):  # Columns B through F
        cell = group_model.cell(row=current_row, column=year_col)
        
        # Different formulas based on KPI
        if kpi == 'Return on Investment (ROI)':
            # Net Income / Total Investment (assumed as 5M)
            cell.value = f'=B{group_net_income_row}/5000000'
            cell.number_format = '0.0%'
        elif kpi == 'Return on Assets (ROA)':
            # Net Income / Total Assets (assumed as 10M)
            cell.value = f'=B{group_net_income_row}/10000000'
            cell.number_format = '0.0%'
        elif kpi == 'Return on Equity (ROE)':
            # Net Income / Shareholders' Equity (assumed as 4M)
            cell.value = f'=B{group_net_income_row}/4000000'
            cell.number_format = '0.0%'
        elif kpi == 'Inventory Turnover':
            # COGS / Average Inventory (assumed as 2M)
            cogs_row = group_net_income_row - 12  # Based on metrics order
            cell.value = f'=B{cogs_row}/2000000'
            cell.number_format = '0.00'
        elif kpi == 'Days Sales Outstanding (DSO)':
            # (Accounts Receivable / Revenue) * 365 (AR assumed as 15% of revenue)
            revenue_row = group_net_income_row - 14  # Based on metrics order
            cell.value = f'=(B{revenue_row}*0.15/B{revenue_row})*365'
            cell.number_format = '0.0'
        elif kpi == 'Days Payable Outstanding (DPO)':
            # (Accounts Payable / COGS) * 365 (AP assumed as 20% of COGS)
            cogs_row = group_net_income_row - 12  # Based on metrics order
            cell.value = f'=(B{cogs_row}*0.2/B{cogs_row})*365'
            cell.number_format = '0.0'
        elif kpi == 'Cash Conversion Cycle':
            # Inventory Days + DSO - DPO
            # Inventory Days = (Inventory / COGS) * 365 (Inventory assumed as 25% of COGS)
            cogs_row = group_net_income_row - 12  # Based on metrics order
            cell.value = f'=(2000000/B{cogs_row})*365+B{current_row-2}-B{current_row-1}'
            cell.number_format = '0.0'
        elif kpi == 'Operating Cash Flow':
            # EBITDA - Change in Working Capital (assumed as 10% of EBITDA)
            ebitda_row = group_net_income_row - 7  # Based on metrics order
            cell.value = f'=B{ebitda_row}*(1-0.1)'
            cell.number_format = '#,##0'
        elif kpi == 'Free Cash Flow':
            # Operating Cash Flow - Capital Expenditures (assumed as 20% of EBITDA)
            ebitda_row = group_net_income_row - 7  # Based on metrics order
            cell.value = f'=B{current_row-1}-B{ebitda_row}*0.2'
            cell.number_format = '#,##0'
        elif kpi == 'Debt to Equity Ratio':
            # Total Debt / Shareholders' Equity (Debt assumed as 6M, Equity as 4M)
            cell.value = '=6000000/4000000'
            cell.number_format = '0.00'
        
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')
        
        # Apply the same formula pattern to other years (columns C through F)
        if year_col > 2:  # For years 2-5
            # Replace the column reference in the formula
            if isinstance(cell.value, str) and cell.value.startswith('='):
                formula = cell.value
                # Replace B with the current column letter
                formula = formula.replace('B', get_column_letter(year_col))
                cell.value = formula

# Add a section for Sensitivity Analysis
sensitivity_row = current_row + 3
group_model.cell(row=sensitivity_row, column=1).value = 'Sensitivity Analysis'
group_model.cell(row=sensitivity_row, column=1).font = Font(bold=True, size=14)
group_model.merge_cells(f'A{sensitivity_row}:F{sensitivity_row}')
group_model.cell(row=sensitivity_row, column=1).alignment = Alignment(horizontal='center')

# Add Sensitivity Analysis headers
sensitivity_row += 1
headers = ['Variable', 'Base Case', '-20%', '-10%', '+10%', '+20%']
for i, header in enumerate(headers):
    cell = group_model.cell(row=sensitivity_row, column=i+1)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

# Define sensitivity variables
sensitivity_vars = [
    'Revenue Growth Rate',
    'COGS as % of Revenue',
    'Operational Costs',
    'SG&A Expenses',
    'Impact on Net Income'
]

# Add sensitivity rows
for i, var in enumerate(sensitivity_vars):
    current_row = sensitivity_row + i + 1
    cell = group_model.cell(row=current_row, column=1)
    cell.value = var
    cell.alignment = Alignment(horizontal='left', vertical='center')
    cell.border = thin_border
    
    if var == 'Revenue Growth Rate':
        base_value = 0.05  # 5% growth rate
        cell = group_model.cell(row=current_row, column=2)
        cell.value = base_value
        cell.number_format = '0.0%'
        cell.border = thin_border
        
        # -20%, -10%, +10%, +20% scenarios
        group_model.cell(row=current_row, column=3).value = base_value * 0.8
        group_model.cell(row=current_row, column=4).value = base_value * 0.9
        group_model.cell(row=current_row, column=5).value = base_value * 1.1
        group_model.cell(row=current_row, column=6).value = base_value * 1.2
        
        for col in range(3, 7):
            group_model.cell(row=current_row, column=col).number_format = '0.0%'
            group_model.cell(row=current_row, column=col).border = thin_border
    
    elif var == 'COGS as % of Revenue':
        base_value = 0.65  # 65% of revenue
        cell = group_model.cell(row=current_row, column=2)
        cell.value = base_value
        cell.number_format = '0.0%'
        cell.border = thin_border
        
        # -20%, -10%, +10%, +20% scenarios
        group_model.cell(row=current_row, column=3).value = base_value * 0.8
        group_model.cell(row=current_row, column=4).value = base_value * 0.9
        group_model.cell(row=current_row, column=5).value = base_value * 1.1
        group_model.cell(row=current_row, column=6).value = base_value * 1.2
        
        for col in range(3, 7):
            group_model.cell(row=current_row, column=col).number_format = '0.0%'
            group_model.cell(row=current_row, column=col).border = thin_border
    
    elif var == 'Operational Costs':
        base_value = 1000000  # $1M
        cell = group_model.cell(row=current_row, column=2)
        cell.value = base_value
        cell.number_format = '#,##0'
        cell.border = thin_border
        
        # -20%, -10%, +10%, +20% scenarios
        group_model.cell(row=current_row, column=3).value = base_value * 0.8
        group_model.cell(row=current_row, column=4).value = base_value * 0.9
        group_model.cell(row=current_row, column=5).value = base_value * 1.1
        group_model.cell(row=current_row, column=6).value = base_value * 1.2
        
        for col in range(3, 7):
            group_model.cell(row=current_row, column=col).number_format = '#,##0'
            group_model.cell(row=current_row, column=col).border = thin_border
    
    elif var == 'SG&A Expenses':
        base_value = 800000  # $800K
        cell = group_model.cell(row=current_row, column=2)
        cell.value = base_value
        cell.number_format = '#,##0'
        cell.border = thin_border
        
        # -20%, -10%, +10%, +20% scenarios
        group_model.cell(row=current_row, column=3).value = base_value * 0.8
        group_model.cell(row=current_row, column=4).value = base_value * 0.9
        group_model.cell(row=current_row, column=5).value = base_value * 1.1
        group_model.cell(row=current_row, column=6).value = base_value * 1.2
        
        for col in range(3, 7):
            group_model.cell(row=current_row, column=col).number_format = '#,##0'
            group_model.cell(row=current_row, column=col).border = thin_border
    
    elif var == 'Impact on Net Income':
        base_value = 500000  # $500K base net income
        cell = group_model.cell(row=current_row, column=2)
        cell.value = base_value
        cell.number_format = '#,##0'
        cell.border = thin_border
        
        # Impact scenarios based on the above variables
        # These are simplified calculations for demonstration
        group_model.cell(row=current_row, column=3).value = base_value * 0.7  # -30%
        group_model.cell(row=current_row, column=4).value = base_value * 0.85  # -15%
        group_model.cell(row=current_row, column=5).value = base_value * 1.15  # +15%
        group_model.cell(row=current_row, column=6).value = base_value * 1.3  # +30%
        
        for col in range(3, 7):
            group_model.cell(row=current_row, column=col).number_format = '#,##0'
            group_model.cell(row=current_row, column=col).border = thin_border
            # Highlight the impact cells
            group_model.cell(row=current_row, column=col).fill = calculation_fill

# Adjust column widths
group_model.column_dimensions['A'].width = 30
for col in range(2, 7):
    group_model.column_dimensions[get_column_letter(col)].width = 15

# Save the workbook
workbook.save(excel_file)

print(f"Group Model sheet created in {excel_file}")
