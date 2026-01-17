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
    print(f"Error: File {excel_file} not found. Please run create_assumptions_sheet.py first.")
    exit(1)

# Create operational costs data
operational_costs_data = {
    'Cost Category': [
        # Facility Costs
        'Facility Costs', 'Facility Costs', 'Facility Costs', 'Facility Costs', 'Facility Costs',
        # Equipment Costs
        'Equipment Costs', 'Equipment Costs', 'Equipment Costs',
        # Supply Chain Costs
        'Supply Chain Costs', 'Supply Chain Costs', 'Supply Chain Costs', 'Supply Chain Costs',
        # Quality Control Costs
        'Quality Control Costs', 'Quality Control Costs', 'Quality Control Costs',
        # Sustainability Costs
        'Sustainability Costs', 'Sustainability Costs', 'Sustainability Costs',
        # Operational Staff Costs
        'Operational Staff Costs', 'Operational Staff Costs', 'Operational Staff Costs', 'Operational Staff Costs',
        # Operational Overhead
        'Operational Overhead', 'Operational Overhead', 'Operational Overhead',
    ],
    'Cost Item': [
        # Facility Costs
        'Rent/Lease', 'Utilities', 'Maintenance & Repairs', 'Property Insurance', 'Property Taxes',
        # Equipment Costs
        'Equipment Maintenance', 'Equipment Calibration', 'Small Equipment Purchases',
        # Supply Chain Costs
        'Supply Chain Software', 'Inventory Management', 'Warehouse Operations', 'Supply Chain Consulting',
        # Quality Control Costs
        'Quality Testing', 'Quality Certifications', 'Quality Staff',
        # Sustainability Costs
        'Sustainability Program', 'Waste Management', 'Energy Efficiency Initiatives',
        # Operational Staff Costs
        'Production Management', 'Warehouse Staff', 'Operations Support', 'Staff Training',
        # Operational Overhead
        'IT Systems (Operations)', 'Communications', 'Miscellaneous Operational Expenses',
    ],
    'Cost Driver': [
        # Facility Costs
        'Square footage × rate with annual increase', 'Usage rates × utility prices with annual increase', 'Percentage of facility value', 'Insurance premium based on asset value', 'Local tax rates on property value',
        # Equipment Costs
        'Percentage of equipment value', 'Fixed cost per calibration × frequency', 'Replacement cycle for minor equipment',
        # Supply Chain Costs
        'Fixed annual cost with increases', 'Percentage of inventory value', 'Per square foot cost × warehouse size', 'Project-based consulting fees',
        # Quality Control Costs
        'Cost per test × number of batches', 'Annual certification fees', 'Headcount × average salary',
        # Sustainability Costs
        'Fixed annual budget with increases', 'Volume of waste × disposal cost', 'Project-based initiatives',
        # Operational Staff Costs
        'Headcount × average salary', 'Headcount × average salary', 'Headcount × average salary', 'Training budget per employee',
        # Operational Overhead
        'Per user license fees × users', 'Fixed monthly costs', 'Percentage of total operational costs',
    ],
    'Monthly Cost Base': [
        # Facility Costs
        '=Assumptions!C19/12', '=Assumptions!C20/12', '=Assumptions!C21/12', '=Assumptions!C22/12', '=Assumptions!C23/12',
        # Equipment Costs
        '=Assumptions!C21/12', '=2000', '=3000',
        # Supply Chain Costs
        '=Assumptions!C24/12', '=Assumptions!C22/12', '=5000', '=2500',
        # Quality Control Costs
        '=4000', '=1500', '=8000',
        # Sustainability Costs
        '=Assumptions!C25/12', '=3000', '=2000',
        # Operational Staff Costs
        '=12000', '=8000', '=6000', '=Assumptions!C47/12',
        # Operational Overhead
        '=2500', '=1000', '=0.02*SUM(D4:D26)',
    ],
    'Units': [
        # Facility Costs
        '12', '12', '12', '12', '12',
        # Equipment Costs
        '12', '4', '2',
        # Supply Chain Costs
        '12', '12', '12', '4',
        # Quality Control Costs
        '24', '2', '5',
        # Sustainability Costs
        '12', '12', '4',
        # Operational Staff Costs
        '8', '15', '6', '29',
        # Operational Overhead
        '20', '12', '1',
    ],
    'Year 1 Total': [
        # Facility Costs
        '', '', '', '', '',
        # Equipment Costs
        '', '', '',
        # Supply Chain Costs
        '', '', '', '',
        # Quality Control Costs
        '', '', '',
        # Sustainability Costs
        '', '', '',
        # Operational Staff Costs
        '', '', '', '',
        # Operational Overhead
        '', '', '',
    ],
    'Year 2 Total': [
        # Facility Costs
        '', '', '', '', '',
        # Equipment Costs
        '', '', '',
        # Supply Chain Costs
        '', '', '', '',
        # Quality Control Costs
        '', '', '',
        # Sustainability Costs
        '', '', '',
        # Operational Staff Costs
        '', '', '', '',
        # Operational Overhead
        '', '', '',
    ],
    'Year 3 Total': [
        # Facility Costs
        '', '', '', '', '',
        # Equipment Costs
        '', '', '',
        # Supply Chain Costs
        '', '', '', '',
        # Quality Control Costs
        '', '', '',
        # Sustainability Costs
        '', '', '',
        # Operational Staff Costs
        '', '', '', '',
        # Operational Overhead
        '', '', '',
    ],
    'Year 4 Total': [
        # Facility Costs
        '', '', '', '', '',
        # Equipment Costs
        '', '', '',
        # Supply Chain Costs
        '', '', '', '',
        # Quality Control Costs
        '', '', '',
        # Sustainability Costs
        '', '', '',
        # Operational Staff Costs
        '', '', '', '',
        # Operational Overhead
        '', '', '',
    ],
    'Year 5 Total': [
        # Facility Costs
        '', '', '', '', '',
        # Equipment Costs
        '', '', '',
        # Supply Chain Costs
        '', '', '', '',
        # Quality Control Costs
        '', '', '',
        # Sustainability Costs
        '', '', '',
        # Operational Staff Costs
        '', '', '', '',
        # Operational Overhead
        '', '', '',
    ],
    'Notes': [
        # Facility Costs
        'Monthly rent/lease payments for manufacturing and warehouse facilities', 'Electricity, water, gas, and other utility costs', 'Ongoing maintenance and repairs for facilities', 'Insurance coverage for facilities and contents', 'Local property taxes on owned facilities',
        # Equipment Costs
        'Regular maintenance of production and packaging equipment', 'Required calibration of precision equipment', 'Replacement of small equipment items not capitalized',
        # Supply Chain Costs
        'Software licenses for supply chain management systems', 'Costs associated with managing inventory levels', 'Labor and equipment for warehouse operations', 'External supply chain optimization consultants',
        # Quality Control Costs
        'Laboratory testing of raw materials and finished products', 'Industry and regulatory certifications', 'Quality assurance and control personnel',
        # Sustainability Costs
        'Corporate sustainability initiatives and programs', 'Disposal and recycling of production waste', 'Projects to improve energy efficiency',
        # Operational Staff Costs
        'Production managers and supervisors', 'Warehouse and logistics personnel', 'Operations administrative support staff', 'Ongoing training for operational staff',
        # Operational Overhead
        'Software and hardware for operations', 'Phone, internet, and other communication costs', 'Miscellaneous expenses related to operations',
    ]
}

# Create a Pandas DataFrame
operational_costs_df = pd.DataFrame(operational_costs_data)

# Check if the sheet already exists and remove it if it does
if 'Operational Costs' in workbook.sheetnames:
    idx = workbook.sheetnames.index('Operational Costs')
    workbook.remove(workbook.worksheets[idx])

# Create a new sheet
worksheet = workbook.create_sheet('Operational Costs')

# Write the header row
for col_idx, column_name in enumerate(operational_costs_df.columns, 1):
    worksheet.cell(row=2, column=col_idx).value = column_name

# Write the data rows
for row_idx, row in enumerate(operational_costs_df.itertuples(index=False), 3):
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
worksheet['A1'] = 'Operational Costs'
worksheet['A1'].font = title_font
worksheet.merge_cells('A1:I1')
worksheet['A1'].alignment = Alignment(horizontal='center')

# Format header row
for col_idx, value in enumerate(operational_costs_df.columns, 1):
    cell = worksheet.cell(row=2, column=col_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thin_border

# Add formulas for Year 1-5 Total columns
for row_idx in range(3, len(operational_costs_data['Cost Item']) + 3):
    # Year 1 Total formula
    worksheet.cell(row=row_idx, column=6).value = f'=D{row_idx}*E{row_idx}'
    
    # Year 2-5 Total formulas with growth factors from Assumptions
    # Year 2
    if 'Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D19)'
    elif 'Utilities' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D20)'
    elif 'Equipment' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D21)'
    elif 'Supply Chain Software' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D24)'
    elif 'Sustainability Program' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D25)'
    else:
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D4)'  # Default to inflation
    
    # Year 3
    if 'Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E19)'
    elif 'Utilities' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E20)'
    elif 'Equipment' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E21)'
    elif 'Supply Chain Software' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E24)'
    elif 'Sustainability Program' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E25)'
    else:
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E4)'  # Default to inflation
    
    # Year 4
    if 'Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F19)'
    elif 'Utilities' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F20)'
    elif 'Equipment' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F21)'
    elif 'Supply Chain Software' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F24)'
    elif 'Sustainability Program' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F25)'
    else:
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F4)'  # Default to inflation
    
    # Year 5
    if 'Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G19)'
    elif 'Utilities' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G20)'
    elif 'Equipment' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G21)'
    elif 'Supply Chain Software' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G24)'
    elif 'Sustainability Program' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G25)'
    else:
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G4)'  # Default to inflation

# Format data rows
last_category = None
for row_idx in range(3, len(operational_costs_data['Cost Item']) + 3):
    category_cell = worksheet.cell(row=row_idx, column=1)
    cost_item_cell = worksheet.cell(row=row_idx, column=2)
    
    # Apply border to all cells in the row
    for col_idx in range(1, 12):  # Columns A through K
        cell = worksheet.cell(row=row_idx, column=col_idx)
        cell.border = thin_border
        
        # Format numbers in calculation columns
        if col_idx >= 4 and col_idx <= 10:  # Columns D through J
            cell.number_format = '#,##0'
            if col_idx == 4:  # Units column
                cell.number_format = '0'
    
    # Format Category column
    if category_cell.value and category_cell.value != last_category:
        last_category = category_cell.value
        category_cell.font = category_font
        category_cell.fill = category_fill
        category_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
    
    # Format Cost Item column
    cost_item_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
    
    # Format calculation cells
    for col_idx in range(4, 11):  # Columns D through J
        cell = worksheet.cell(row=row_idx, column=col_idx)
        cell.fill = calculation_fill
        cell.alignment = Alignment(horizontal='right')

# Add subtotals for each category
last_row = len(operational_costs_data['Cost Item']) + 3
row_idx = last_row + 1

# Add a subtotal row
worksheet.cell(row=row_idx, column=1).value = 'Total Operational Costs'
worksheet.cell(row=row_idx, column=1).font = Font(bold=True)
worksheet.merge_cells(f'A{row_idx}:E{row_idx}')

# Add subtotal formulas
for col_idx in range(6, 11):  # Columns F through J
    cell = worksheet.cell(row=row_idx, column=col_idx)
    col_letter = get_column_letter(col_idx)
    cell.value = f'=SUM({col_letter}3:{col_letter}{last_row})'
    cell.font = Font(bold=True)
    cell.fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
    cell.number_format = '#,##0'

# Add a row for percentage of revenue
row_idx += 2
worksheet.cell(row=row_idx, column=1).value = 'Operational Costs as % of Revenue'
worksheet.cell(row=row_idx, column=1).font = Font(bold=True, italic=True)
worksheet.merge_cells(f'A{row_idx}:E{row_idx}')

# Add percentage formulas (these will be linked to Revenue sheet later)
for col_idx in range(6, 11):  # Columns F through J
    cell = worksheet.cell(row=row_idx, column=col_idx)
    col_letter = get_column_letter(col_idx)
    cell.value = f'={col_letter}{row_idx-2}/100000'  # Placeholder formula, will be updated when Revenue sheet is created
    cell.font = Font(bold=True, italic=True)
    cell.number_format = '0.0%'

# Adjust column widths
worksheet.column_dimensions['A'].width = 20  # Cost Category
worksheet.column_dimensions['B'].width = 25  # Cost Item
worksheet.column_dimensions['C'].width = 35  # Cost Driver
worksheet.column_dimensions['D'].width = 15  # Monthly Cost Base
worksheet.column_dimensions['E'].width = 10  # Units
worksheet.column_dimensions['F'].width = 15  # Year 1 Total
worksheet.column_dimensions['G'].width = 15  # Year 2 Total
worksheet.column_dimensions['H'].width = 15  # Year 3 Total
worksheet.column_dimensions['I'].width = 15  # Year 4 Total
worksheet.column_dimensions['J'].width = 15  # Year 5 Total
worksheet.column_dimensions['K'].width = 40  # Notes

# Save the workbook
workbook.save(excel_file)

print(f"Operational Costs sheet created in {excel_file}")
