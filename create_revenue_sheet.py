import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Color
from openpyxl.utils import get_column_letter
import os

from financial_model_config import load_existing_workbook, parse_output_args, resolve_output_path, save_workbook


def build_sheet(workbook, excel_file):
    # Create Revenue and Profitability data
    revenue_data = {
        'Revenue Category': [
            # Product Categories
            'Product Categories', 'Product Categories', 'Product Categories',
            # Distribution Channels
            'Distribution Channels', 'Distribution Channels', 'Distribution Channels',
            # Geographic Markets
            'Geographic Markets', 'Geographic Markets', 'Geographic Markets',
        ],
        'Revenue Stream': [
            # Product Categories
            'Facial Skincare', 'Body Care', 'Specialty Products',
            # Distribution Channels
            'Direct-to-Consumer (DTC)', 'Retail', 'Wholesale/B2B',
            # Geographic Markets
            'Domestic', 'International - Developed', 'International - Emerging',
        ],
        'Revenue Driver': [
            # Product Categories
            'Units sold × average price', 'Units sold × average price', 'Units sold × average price',
            # Distribution Channels
            'Website traffic × conversion rate × average order value', 'Store count × average sales per store', 'Number of wholesale accounts × average order size × order frequency',
            # Geographic Markets
            'Market size × market share', 'Market size × market share', 'Market size × market share',
        ],
        'Year 1 Units': [
            # Product Categories
            '500000', '300000', '100000',
            # Distribution Channels
            '400000', '350000', '150000',
            # Geographic Markets
            '600000', '200000', '100000',
        ],
        'Year 1 Price': [
            # Product Categories
            '25.00', '22.00', '35.00',
            # Distribution Channels
            '28.00', '24.00', '20.00',
            # Geographic Markets
            '26.00', '30.00', '22.00',
        ],
        'Year 1 Revenue': [
            # Formulas will be added later
        ],
        'Year 2 Revenue': [
            # Formulas will be added later
        ],
        'Year 3 Revenue': [
            # Formulas will be added later
        ],
        'Year 4 Revenue': [
            # Formulas will be added later
        ],
        'Year 5 Revenue': [
            # Formulas will be added later
        ],
        'Notes': [
            # Product Categories
            'Cleansers, moisturizers, serums, masks for facial care', 'Lotions, scrubs, oils for body care', 'Anti-aging, acne treatment, sun protection',
            # Distribution Channels
            'Sales through company website and branded apps', 'Sales through retail partners and own stores', 'Sales to salons, spas, and other businesses',
            # Geographic Markets
            'Sales within primary country of operation', 'Sales in established international markets', 'Sales in developing international markets',
        ]
    }

    # Create empty columns for Year 1-5 Revenue
    for year in ['Year 1 Revenue', 'Year 2 Revenue', 'Year 3 Revenue', 'Year 4 Revenue', 'Year 5 Revenue']:
        revenue_data[year] = [''] * len(revenue_data['Revenue Stream'])

    # Create a Pandas DataFrame
    revenue_df = pd.DataFrame(revenue_data)

    # Check if the sheet already exists and remove it if it does
    if 'Revenue' in workbook.sheetnames:
        idx = workbook.sheetnames.index('Revenue')
        workbook.remove(workbook.worksheets[idx])

    # Create a new sheet
    worksheet = workbook.create_sheet('Revenue')

    # Write the header row
    for col_idx, column_name in enumerate(revenue_df.columns, 1):
        worksheet.cell(row=2, column=col_idx).value = column_name

    # Write the data rows
    for row_idx, row in enumerate(revenue_df.itertuples(index=False), 3):
        for col_idx, value in enumerate(row, 1):
            worksheet.cell(row=row_idx, column=col_idx).value = value

    # --- Formatting --- 
    # Define styles
    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    category_font = Font(bold=True)
    category_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    calculation_fill = PatternFill(start_color="E6F1DC", end_color="E6F1DC", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Add title
    worksheet['A1'] = 'Revenue and Profitability'
    worksheet['A1'].font = title_font
    worksheet.merge_cells('A1:K1')
    worksheet['A1'].alignment = Alignment(horizontal='center')

    # Format header row
    for col_idx, value in enumerate(revenue_df.columns, 1):
        cell = worksheet.cell(row=2, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

    # Add formulas for Year 1-5 Revenue columns
    for row_idx in range(3, len(revenue_data['Revenue Stream']) + 3):
        # Year 1 Revenue formula
        worksheet.cell(row=row_idx, column=6).value = f'=D{row_idx}*E{row_idx}'
    
        # Year 2-5 Revenue formulas with growth factors from Assumptions
        # Year 2
        if 'Facial Skincare' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D8)*(1+Assumptions!D12)'
        elif 'Body Care' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D9)*(1+Assumptions!D12)'
        elif 'Specialty Products' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D10)*(1+Assumptions!D12)'
        elif 'Direct-to-Consumer' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D10)*(1+Assumptions!D12)'
        elif 'Retail' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D11)*(1+Assumptions!D12)'
        else:
            worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D7)*(1+Assumptions!D12)'
    
        # Year 3
        if 'Facial Skincare' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E8)*(1+Assumptions!E12)'
        elif 'Body Care' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E9)*(1+Assumptions!E12)'
        elif 'Specialty Products' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E10)*(1+Assumptions!E12)'
        elif 'Direct-to-Consumer' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E10)*(1+Assumptions!E12)'
        elif 'Retail' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E11)*(1+Assumptions!E12)'
        else:
            worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E7)*(1+Assumptions!E12)'
    
        # Year 4
        if 'Facial Skincare' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F8)*(1+Assumptions!F12)'
        elif 'Body Care' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F9)*(1+Assumptions!F12)'
        elif 'Specialty Products' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F10)*(1+Assumptions!F12)'
        elif 'Direct-to-Consumer' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F10)*(1+Assumptions!F12)'
        elif 'Retail' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F11)*(1+Assumptions!F12)'
        else:
            worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F7)*(1+Assumptions!F12)'
    
        # Year 5
        if 'Facial Skincare' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G8)*(1+Assumptions!G12)'
        elif 'Body Care' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G9)*(1+Assumptions!G12)'
        elif 'Specialty Products' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G10)*(1+Assumptions!G12)'
        elif 'Direct-to-Consumer' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G10)*(1+Assumptions!G12)'
        elif 'Retail' in str(worksheet.cell(row=row_idx, column=2).value):
            worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G11)*(1+Assumptions!G12)'
        else:
            worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G7)*(1+Assumptions!G12)'

    # Format data rows
    last_category = None
    for row_idx in range(3, len(revenue_data['Revenue Stream']) + 3):
        category_cell = worksheet.cell(row=row_idx, column=1)
        revenue_stream_cell = worksheet.cell(row=row_idx, column=2)
    
        # Apply border to all cells in the row
        for col_idx in range(1, 12):  # Columns A through K
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
        
            # Format numbers in calculation columns
            if col_idx == 4:  # Units column
                cell.number_format = '#,##0'
            elif col_idx == 5:  # Price column
                cell.number_format = '#,##0.00'
            elif col_idx >= 6 and col_idx <= 10:  # Revenue columns
                cell.number_format = '#,##0'
    
        # Format Category column
        if category_cell.value and category_cell.value != last_category:
            last_category = category_cell.value
            category_cell.font = category_font
            category_cell.fill = category_fill
            category_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
    
        # Format Revenue Stream column
        revenue_stream_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
    
        # Format calculation cells
        for col_idx in range(4, 11):  # Columns D through J
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.fill = calculation_fill
            cell.alignment = Alignment(horizontal='right')

    # Add subtotals by category
    last_row = len(revenue_data['Revenue Stream']) + 3
    row_idx = last_row + 1

    # Add a total revenue row
    worksheet.cell(row=row_idx, column=1).value = 'Total Revenue'
    worksheet.cell(row=row_idx, column=1).font = Font(bold=True)
    worksheet.merge_cells(f'A{row_idx}:C{row_idx}')

    # Add total revenue formulas
    for col_idx in range(6, 11):  # Columns F through J
        cell = worksheet.cell(row=row_idx, column=col_idx)
        col_letter = get_column_letter(col_idx)
        cell.value = f'=SUM({col_letter}3:{col_letter}{last_row})'
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
        cell.number_format = '#,##0'

    # Add profitability section
    row_idx += 3
    profitability_start_row = row_idx
    worksheet.cell(row=profitability_start_row, column=1).value = 'Profitability Analysis'
    worksheet.cell(row=profitability_start_row, column=1).font = Font(bold=True, size=12)
    worksheet.merge_cells(f'A{profitability_start_row}:J{profitability_start_row}')
    worksheet.cell(row=profitability_start_row, column=1).alignment = Alignment(horizontal='center')

    # Add profitability headers
    row_idx += 1
    profitability_headers = ['Metric', '', '', '', '', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5']
    for col_idx, header in enumerate(profitability_headers, 1):
        cell = worksheet.cell(row=row_idx, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Add profitability metrics
    profitability_metrics = [
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

    # Define row numbers for COGS, Operational Costs, and SG&A sheets
    # These are fixed values based on the structure of those sheets
    cogs_total_row = 30  # Row number for total COGS in COGS sheet
    operational_costs_total_row = 30  # Row number for total operational costs in Operational Costs sheet
    sga_total_row = 40  # Row number for total SG&A expenses in SG&A sheet

    # Add profitability rows
    profitability_data_start_row = row_idx + 1
    for i, metric in enumerate(profitability_metrics):
        current_metric_row = profitability_data_start_row + i
        worksheet.cell(row=current_metric_row, column=1).value = metric
        worksheet.merge_cells(f'A{current_metric_row}:E{current_metric_row}')
        worksheet.cell(row=current_metric_row, column=1).alignment = Alignment(vertical='center', horizontal='left')
    
        # Add formulas for each year
        for year_col in range(6, 11):  # Columns F through J
            cell = worksheet.cell(row=current_metric_row, column=year_col)
            year_letter = get_column_letter(year_col)
        
            # Different formulas based on the metric
            if metric == 'Revenue':
                # Link to total revenue from above
                total_revenue_row = last_row + 1
                cell.value = f'={year_letter}{total_revenue_row}'
            elif metric == 'Cost of Goods Sold (COGS)':
                # Link to COGS sheet with fixed row number
                cell.value = f'=COGS!{year_letter}{cogs_total_row}'
            elif metric == 'Gross Profit':
                # Revenue - COGS
                cell.value = f'=INDEX(F:J,{current_metric_row}-2,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))-INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
            elif metric == 'Gross Margin %':
                # Gross Profit / Revenue
                cell.value = f'=INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))/INDEX(F:J,{current_metric_row}-3,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
                cell.number_format = '0.0%'
            elif metric == 'Operational Costs':
                # Link to Operational Costs sheet with fixed row number
                cell.value = f'=\'Operational Costs\'!{year_letter}{operational_costs_total_row}'
            elif metric == 'SG&A Expenses':
                # Link to SG&A sheet with fixed row number
                cell.value = f'=\'SG&A\'!{year_letter}{sga_total_row}'
            elif metric == 'EBITDA':
                # Gross Profit - Operational Costs - SG&A
                cell.value = f'=INDEX(F:J,{current_metric_row}-4,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))-INDEX(F:J,{current_metric_row}-2,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))-INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
            elif metric == 'EBITDA Margin %':
                # EBITDA / Revenue
                cell.value = f'=INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))/INDEX(F:J,{current_metric_row}-7,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
                cell.number_format = '0.0%'
            elif metric == 'Depreciation & Amortization':
                # Placeholder for D&A
                cell.value = '=100000'
            elif metric == 'EBIT':
                # EBITDA - D&A
                cell.value = f'=INDEX(F:J,{current_metric_row}-3,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))-INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
            elif metric == 'Interest Expense':
                # Placeholder for interest expense
                cell.value = f'=Assumptions!{year_letter}61*1000000'
            elif metric == 'EBT':
                # EBIT - Interest
                cell.value = f'=INDEX(F:J,{current_metric_row}-2,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))-INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
            elif metric == 'Income Tax':
                # EBT * Tax Rate
                cell.value = f'=INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))*Assumptions!{year_letter}6'
            elif metric == 'Net Income':
                # EBT - Tax
                cell.value = f'=INDEX(F:J,{current_metric_row}-2,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))-INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
            elif metric == 'Net Profit Margin %':
                # Net Income / Revenue
                cell.value = f'=INDEX(F:J,{current_metric_row}-1,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))/INDEX(F:J,{current_metric_row}-14,MATCH({year_letter}${profitability_data_start_row-1},F${profitability_data_start_row-1}:J${profitability_data_start_row-1},0))'
                cell.number_format = '0.0%'
        
            # Format cells
            if '%' in metric:
                cell.number_format = '0.0%'
            else:
                cell.number_format = '#,##0'
        
            # Add borders
            cell.border = thin_border
        
            # Highlight key metrics
            if metric in ['Gross Profit', 'EBITDA', 'Net Income']:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
            elif metric in ['Gross Margin %', 'EBITDA Margin %', 'Net Profit Margin %']:
                cell.font = Font(bold=True, italic=True)
                cell.fill = PatternFill(start_color="E6F1DC", end_color="E6F1DC", fill_type="solid")

    # Add scenario analysis section
    row_idx = profitability_data_start_row + len(profitability_metrics) + 2
    scenario_start_row = row_idx
    worksheet.cell(row=scenario_start_row, column=1).value = 'Scenario Analysis'
    worksheet.cell(row=scenario_start_row, column=1).font = Font(bold=True, size=12)
    worksheet.merge_cells(f'A{scenario_start_row}:J{scenario_start_row}')
    worksheet.cell(row=scenario_start_row, column=1).alignment = Alignment(horizontal='center')

    # Add scenario headers
    row_idx += 1
    scenario_headers = ['Scenario', 'Revenue Growth', 'COGS %', 'SG&A %', 'Year 1 Revenue', 'Year 5 Revenue', 'Year 1 Net Income', 'Year 5 Net Income', 'Year 1 Margin', 'Year 5 Margin']
    for col_idx, header in enumerate(scenario_headers, 1):
        cell = worksheet.cell(row=row_idx, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Add scenario rows
    scenarios = ['Base Case', 'Optimistic', 'Pessimistic']
    growth_rates = ['=Assumptions!C7', '=Assumptions!C7*1.2', '=Assumptions!C7*0.8']
    cogs_pcts = ['=COGS!F29/F24', '=COGS!F29/F24*0.9', '=COGS!F29/F24*1.1']
    sga_pcts = ['=\'SG&A\'!F39/F24', '=\'SG&A\'!F39/F24*0.9', '=\'SG&A\'!F39/F24*1.1']

    # Store the net income row for reference in scenarios
    net_income_row_in_profitability = profitability_data_start_row + profitability_metrics.index('Net Income')

    scenario_data_start_row = row_idx + 1
    for i, scenario in enumerate(scenarios):
        current_scenario_row = scenario_data_start_row + i
    
        # Scenario name
        worksheet.cell(row=current_scenario_row, column=1).value = scenario
        worksheet.cell(row=current_scenario_row, column=1).font = Font(bold=True)
    
        # Growth rate
        worksheet.cell(row=current_scenario_row, column=2).value = growth_rates[i]
        worksheet.cell(row=current_scenario_row, column=2).number_format = '0.0%'
    
        # COGS %
        worksheet.cell(row=current_scenario_row, column=3).value = cogs_pcts[i]
        worksheet.cell(row=current_scenario_row, column=3).number_format = '0.0%'
    
        # SG&A %
        worksheet.cell(row=current_scenario_row, column=4).value = sga_pcts[i]
        worksheet.cell(row=current_scenario_row, column=4).number_format = '0.0%'
    
        # Year 1 Revenue
        if scenario == 'Base Case':
            worksheet.cell(row=current_scenario_row, column=5).value = f'=F{last_row+1}'
        elif scenario == 'Optimistic':
            worksheet.cell(row=current_scenario_row, column=5).value = f'=F{last_row+1}*1.1'
        else:  # Pessimistic
            worksheet.cell(row=current_scenario_row, column=5).value = f'=F{last_row+1}*0.9'
        worksheet.cell(row=current_scenario_row, column=5).number_format = '#,##0'
    
        # Year 5 Revenue
        if scenario == 'Base Case':
            worksheet.cell(row=current_scenario_row, column=6).value = f'=J{last_row+1}'
        elif scenario == 'Optimistic':
            worksheet.cell(row=current_scenario_row, column=6).value = f'=J{last_row+1}*1.2'
        else:  # Pessimistic
            worksheet.cell(row=current_scenario_row, column=6).value = f'=J{last_row+1}*0.8'
        worksheet.cell(row=current_scenario_row, column=6).number_format = '#,##0'
    
        # Year 1 Net Income
        if scenario == 'Base Case':
            worksheet.cell(row=current_scenario_row, column=7).value = f'=F{net_income_row_in_profitability}'
        elif scenario == 'Optimistic':
            worksheet.cell(row=current_scenario_row, column=7).value = f'=F{net_income_row_in_profitability}*1.3'
        else:  # Pessimistic
            worksheet.cell(row=current_scenario_row, column=7).value = f'=F{net_income_row_in_profitability}*0.7'
        worksheet.cell(row=current_scenario_row, column=7).number_format = '#,##0'
    
        # Year 5 Net Income
        if scenario == 'Base Case':
            worksheet.cell(row=current_scenario_row, column=8).value = f'=J{net_income_row_in_profitability}'
        elif scenario == 'Optimistic':
            worksheet.cell(row=current_scenario_row, column=8).value = f'=J{net_income_row_in_profitability}*1.4'
        else:  # Pessimistic
            worksheet.cell(row=current_scenario_row, column=8).value = f'=J{net_income_row_in_profitability}*0.6'
        worksheet.cell(row=current_scenario_row, column=8).number_format = '#,##0'
    
        # Year 1 Margin
        worksheet.cell(row=current_scenario_row, column=9).value = f'=G{current_scenario_row}/E{current_scenario_row}'
        worksheet.cell(row=current_scenario_row, column=9).number_format = '0.0%'
    
        # Year 5 Margin
        worksheet.cell(row=current_scenario_row, column=10).value = f'=H{current_scenario_row}/F{current_scenario_row}'
        worksheet.cell(row=current_scenario_row, column=10).number_format = '0.0%'
    
        # Format all cells in the row
        for col_idx in range(1, 11):
            worksheet.cell(row=current_scenario_row, column=col_idx).border = thin_border
            if col_idx >= 5:  # Format calculation cells
                worksheet.cell(row=current_scenario_row, column=col_idx).fill = calculation_fill

    # Adjust column widths
    worksheet.column_dimensions['A'].width = 20  # Revenue Category
    worksheet.column_dimensions['B'].width = 25  # Revenue Stream
    worksheet.column_dimensions['C'].width = 35  # Revenue Driver
    worksheet.column_dimensions['D'].width = 15  # Year 1 Units
    worksheet.column_dimensions['E'].width = 15  # Year 1 Price
    worksheet.column_dimensions['F'].width = 15  # Year 1 Revenue
    worksheet.column_dimensions['G'].width = 15  # Year 2 Revenue
    worksheet.column_dimensions['H'].width = 15  # Year 3 Revenue
    worksheet.column_dimensions['I'].width = 15  # Year 4 Revenue
    worksheet.column_dimensions['J'].width = 15  # Year 5 Revenue
    worksheet.column_dimensions['K'].width = 40  # Notes


def main(output_path: str | None = None):
    excel_file = resolve_output_path(output_path)
    workbook = load_existing_workbook(excel_file, ['Assumptions', 'COGS', 'Operational Costs', 'SG&A'])
    build_sheet(workbook, str(excel_file))
    save_workbook(workbook, excel_file)
    print(f'Revenue and Profitability sheet created in {excel_file}')
    return excel_file


if __name__ == '__main__':
    args = parse_output_args('Create the revenue sheet for the SkinTwin financial model workbook.')
    main(args.output)
