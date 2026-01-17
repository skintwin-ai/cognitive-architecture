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

# Create SG&A expenses data
sga_data = {
    'Cost Category': [
        # Sales Expenses
        'Sales Expenses', 'Sales Expenses', 'Sales Expenses', 'Sales Expenses', 'Sales Expenses',
        # Marketing Expenses
        'Marketing Expenses', 'Marketing Expenses', 'Marketing Expenses', 'Marketing Expenses', 'Marketing Expenses', 'Marketing Expenses',
        # General Administrative
        'General Administrative', 'General Administrative', 'General Administrative', 'General Administrative', 'General Administrative',
        # Professional Services
        'Professional Services', 'Professional Services', 'Professional Services', 'Professional Services',
        # IT & Technology
        'IT & Technology', 'IT & Technology', 'IT & Technology', 'IT & Technology',
        # Facilities & Office
        'Facilities & Office', 'Facilities & Office', 'Facilities & Office', 'Facilities & Office',
        # Other SG&A
        'Other SG&A', 'Other SG&A', 'Other SG&A',
    ],
    'Cost Item': [
        # Sales Expenses
        'Sales Team Salaries', 'Sales Commissions', 'Travel & Entertainment', 'Trade Shows & Events', 'Customer Relationship Management',
        # Marketing Expenses
        'Digital Marketing', 'Traditional Advertising', 'Social Media Marketing', 'Influencer Partnerships', 'Public Relations', 'Marketing Materials',
        # General Administrative
        'Executive Salaries', 'Administrative Staff', 'HR Department', 'Finance Department', 'Corporate Insurance',
        # Professional Services
        'Legal Services', 'Accounting & Audit', 'Consulting Services', 'Regulatory Compliance',
        # IT & Technology
        'IT Staff', 'Software Licenses', 'Hardware & Equipment', 'IT Support & Maintenance',
        # Facilities & Office
        'Office Rent', 'Office Utilities', 'Office Supplies', 'Office Maintenance',
        # Other SG&A
        'Training & Development', 'Travel (Non-Sales)', 'Miscellaneous SG&A',
    ],
    'Cost Driver': [
        # Sales Expenses
        'Headcount × average salary', 'Percentage of revenue', 'Per sales employee × travel frequency', 'Fixed cost per event × number of events', 'Per user license fee × number of users',
        # Marketing Expenses
        'Percentage of revenue', 'Percentage of revenue', 'Percentage of revenue', 'Fixed fee per partnership × number of partnerships', 'Monthly retainer × 12', 'Fixed cost per campaign × number of campaigns',
        # General Administrative
        'Headcount × average salary', 'Headcount × average salary', 'Headcount × average salary', 'Headcount × average salary', 'Percentage of revenue',
        # Professional Services
        'Monthly retainer + hourly fees', 'Annual audit fee + monthly accounting', 'Project-based fees', 'Fixed annual cost',
        # IT & Technology
        'Headcount × average salary', 'Per user license fee × number of users', 'Depreciation + new purchases', 'Monthly service fee × 12',
        # Facilities & Office
        'Square footage × rate with annual increase', 'Usage rates × utility prices', 'Per employee × headcount', 'Percentage of facility value',
        # Other SG&A
        'Training budget per employee × headcount', 'Per employee × travel frequency', 'Percentage of total SG&A',
    ],
    'Annual Cost Base': [
        # Sales Expenses
        '=Assumptions!C31*8', '=Assumptions!C32', '=15000', '=25000*4', '=Assumptions!C36*25',
        # Marketing Expenses
        '=Assumptions!C33', '=0.03*1000000', '=0.02*1000000', '=20000*12', '=10000*12', '=15000*4',
        # General Administrative
        '=250000*4', '=Assumptions!C34*10', '=80000*3', '=90000*4', '=0.005*1000000',
        # Professional Services
        '=Assumptions!C37', '=Assumptions!C37*0.8', '=50000*2', '=75000',
        # IT & Technology
        '=120000*3', '=Assumptions!C36*50', '=100000', '=8000*12',
        # Facilities & Office
        '=Assumptions!C35*5000', '=4000*12', '=500*50', '=0.02*1000000',
        # Other SG&A
        '=Assumptions!C47*50', '=3000*20', '=0.01*SUM(D3:D29)',
    ],
    'Year 1 Total': [
        # Formulas will be added later
    ],
    'Year 2 Total': [
        # Formulas will be added later
    ],
    'Year 3 Total': [
        # Formulas will be added later
    ],
    'Year 4 Total': [
        # Formulas will be added later
    ],
    'Year 5 Total': [
        # Formulas will be added later
    ],
    'Notes': [
        # Sales Expenses
        'Base salaries for sales team members', 'Variable compensation based on sales performance', 'Customer meetings, sales trips, and entertainment expenses', 'Industry exhibitions, conferences, and sales events', 'Software for managing customer relationships and sales pipeline',
        # Marketing Expenses
        'Online advertising, SEO, email marketing, and other digital channels', 'Print, TV, radio, and other traditional media advertising', 'Content creation, community management, and paid social campaigns', 'Partnerships with industry influencers and content creators', 'Media relations, press releases, and publicity management', 'Brochures, catalogs, samples, and other physical marketing materials',
        # General Administrative
        'C-suite and senior management compensation', 'Support staff, office management, and administrative personnel', 'Recruitment, benefits administration, and employee relations', 'Accounting, financial planning, and treasury operations', 'General liability, D&O, and other corporate insurance policies',
        # Professional Services
        'Outside counsel, legal advice, and contract review', 'External accounting services and annual financial audits', 'Business strategy, operations, and specialized consulting', 'Costs for ensuring compliance with industry regulations',
        # IT & Technology
        'IT department personnel and support staff', 'Software subscriptions and enterprise applications', 'Computers, servers, networking equipment, and peripherals', 'Ongoing IT maintenance, security, and support services',
        # Facilities & Office
        'Lease or rent payments for office space', 'Electricity, water, heating, and other office utilities', 'Office consumables, stationery, and supplies', 'Cleaning, repairs, and general office maintenance',
        # Other SG&A
        'Employee development programs and educational resources', 'Non-sales related business travel expenses', 'Miscellaneous selling, general, and administrative expenses',
    ]
}

# Create empty columns for Year 1-5 Total
for year in ['Year 1 Total', 'Year 2 Total', 'Year 3 Total', 'Year 4 Total', 'Year 5 Total']:
    sga_data[year] = [''] * len(sga_data['Cost Item'])

# Create a Pandas DataFrame
sga_df = pd.DataFrame(sga_data)

# Check if the sheet already exists and remove it if it does
if 'SG&A' in workbook.sheetnames:
    idx = workbook.sheetnames.index('SG&A')
    workbook.remove(workbook.worksheets[idx])

# Create a new sheet
worksheet = workbook.create_sheet('SG&A')

# Write the header row
for col_idx, column_name in enumerate(sga_df.columns, 1):
    worksheet.cell(row=2, column=col_idx).value = column_name

# Write the data rows
for row_idx, row in enumerate(sga_df.itertuples(index=False), 3):
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
worksheet['A1'] = 'Selling, General & Administrative (SG&A) Expenses'
worksheet['A1'].font = title_font
worksheet.merge_cells('A1:I1')
worksheet['A1'].alignment = Alignment(horizontal='center')

# Format header row
for col_idx, value in enumerate(sga_df.columns, 1):
    cell = worksheet.cell(row=2, column=col_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thin_border

# Add formulas for Year 1-5 Total columns
for row_idx in range(3, len(sga_data['Cost Item']) + 3):
    # Year 1 Total formula - directly use the Annual Cost Base
    worksheet.cell(row=row_idx, column=5).value = f'=D{row_idx}'
    
    # Year 2-5 Total formulas with growth factors from Assumptions
    # Year 2
    if 'Sales Team Salaries' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D31)'
    elif 'Sales Commissions' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D32)'
    elif 'Digital Marketing' in str(worksheet.cell(row=row_idx, column=2).value) or 'Traditional Advertising' in str(worksheet.cell(row=row_idx, column=2).value) or 'Social Media Marketing' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D33)'
    elif 'Administrative Staff' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D34)'
    elif 'Office Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D35)'
    elif 'Software Licenses' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D36)'
    elif 'Legal Services' in str(worksheet.cell(row=row_idx, column=2).value) or 'Accounting & Audit' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D37)'
    else:
        worksheet.cell(row=row_idx, column=6).value = f'=E{row_idx}*(1+Assumptions!D4)'  # Default to inflation
    
    # Year 3
    if 'Sales Team Salaries' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E31)'
    elif 'Sales Commissions' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E32)'
    elif 'Digital Marketing' in str(worksheet.cell(row=row_idx, column=2).value) or 'Traditional Advertising' in str(worksheet.cell(row=row_idx, column=2).value) or 'Social Media Marketing' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E33)'
    elif 'Administrative Staff' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E34)'
    elif 'Office Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E35)'
    elif 'Software Licenses' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E36)'
    elif 'Legal Services' in str(worksheet.cell(row=row_idx, column=2).value) or 'Accounting & Audit' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E37)'
    else:
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!E4)'  # Default to inflation
    
    # Year 4
    if 'Sales Team Salaries' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F31)'
    elif 'Sales Commissions' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F32)'
    elif 'Digital Marketing' in str(worksheet.cell(row=row_idx, column=2).value) or 'Traditional Advertising' in str(worksheet.cell(row=row_idx, column=2).value) or 'Social Media Marketing' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F33)'
    elif 'Administrative Staff' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F34)'
    elif 'Office Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F35)'
    elif 'Software Licenses' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F36)'
    elif 'Legal Services' in str(worksheet.cell(row=row_idx, column=2).value) or 'Accounting & Audit' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F37)'
    else:
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!F4)'  # Default to inflation
    
    # Year 5
    if 'Sales Team Salaries' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G31)'
    elif 'Sales Commissions' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G32)'
    elif 'Digital Marketing' in str(worksheet.cell(row=row_idx, column=2).value) or 'Traditional Advertising' in str(worksheet.cell(row=row_idx, column=2).value) or 'Social Media Marketing' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G33)'
    elif 'Administrative Staff' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G34)'
    elif 'Office Rent' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G35)'
    elif 'Software Licenses' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G36)'
    elif 'Legal Services' in str(worksheet.cell(row=row_idx, column=2).value) or 'Accounting & Audit' in str(worksheet.cell(row=row_idx, column=2).value):
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G37)'
    else:
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!G4)'  # Default to inflation

# Format data rows
last_category = None
for row_idx in range(3, len(sga_data['Cost Item']) + 3):
    category_cell = worksheet.cell(row=row_idx, column=1)
    cost_item_cell = worksheet.cell(row=row_idx, column=2)
    
    # Apply border to all cells in the row
    for col_idx in range(1, 11):  # Columns A through J
        cell = worksheet.cell(row=row_idx, column=col_idx)
        cell.border = thin_border
        
        # Format numbers in calculation columns
        if col_idx >= 4 and col_idx <= 9:  # Columns D through I
            cell.number_format = '#,##0'
    
    # Format Category column
    if category_cell.value and category_cell.value != last_category:
        last_category = category_cell.value
        category_cell.font = category_font
        category_cell.fill = category_fill
        category_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
    
    # Format Cost Item column
    cost_item_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
    
    # Format calculation cells
    for col_idx in range(4, 10):  # Columns D through I
        cell = worksheet.cell(row=row_idx, column=col_idx)
        cell.fill = calculation_fill
        cell.alignment = Alignment(horizontal='right')

# Add subtotals by category
last_row = len(sga_data['Cost Item']) + 3
row_idx = last_row + 1

# Add subtotal rows for each category
categories = ['Sales Expenses', 'Marketing Expenses', 'General Administrative', 'Professional Services', 'IT & Technology', 'Facilities & Office', 'Other SG&A']
category_start_rows = {}
category_end_rows = {}

# Find start and end rows for each category
current_category = None
for row_idx in range(3, last_row + 1):
    category = worksheet.cell(row=row_idx, column=1).value
    if category != current_category:
        if current_category is not None:
            category_end_rows[current_category] = row_idx - 1
        current_category = category
        category_start_rows[current_category] = row_idx
category_end_rows[current_category] = last_row  # Set end row for last category

# Add subtotal row for each category
row_idx = last_row + 1
for category in categories:
    start_row = category_start_rows[category]
    end_row = category_end_rows[category]
    
    # Add category subtotal row
    worksheet.cell(row=row_idx, column=1).value = f'Total {category}'
    worksheet.cell(row=row_idx, column=1).font = Font(bold=True)
    worksheet.merge_cells(f'A{row_idx}:C{row_idx}')
    
    # Add subtotal formulas
    for col_idx in range(5, 10):  # Columns E through I
        cell = worksheet.cell(row=row_idx, column=col_idx)
        col_letter = get_column_letter(col_idx)
        cell.value = f'=SUM({col_letter}{start_row}:{col_letter}{end_row})'
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        cell.number_format = '#,##0'
    
    row_idx += 1

# Add a grand total row
worksheet.cell(row=row_idx, column=1).value = 'Total SG&A Expenses'
worksheet.cell(row=row_idx, column=1).font = Font(bold=True)
worksheet.merge_cells(f'A{row_idx}:C{row_idx}')

# Add grand total formulas
for col_idx in range(5, 10):  # Columns E through I
    cell = worksheet.cell(row=row_idx, column=col_idx)
    col_letter = get_column_letter(col_idx)
    cell.value = f'=SUM({col_letter}3:{col_letter}{last_row})'
    cell.font = Font(bold=True)
    cell.fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
    cell.number_format = '#,##0'

# Add a row for percentage of revenue
row_idx += 2
worksheet.cell(row=row_idx, column=1).value = 'SG&A as % of Revenue'
worksheet.cell(row=row_idx, column=1).font = Font(bold=True, italic=True)
worksheet.merge_cells(f'A{row_idx}:C{row_idx}')

# Add percentage formulas (these will be linked to Revenue sheet later)
for col_idx in range(5, 10):  # Columns E through I
    cell = worksheet.cell(row=row_idx, column=col_idx)
    col_letter = get_column_letter(col_idx)
    cell.value = f'={col_letter}{row_idx-2}/1000000'  # Placeholder formula, will be updated when Revenue sheet is created
    cell.font = Font(bold=True, italic=True)
    cell.number_format = '0.0%'

# Adjust column widths
worksheet.column_dimensions['A'].width = 20  # Cost Category
worksheet.column_dimensions['B'].width = 25  # Cost Item
worksheet.column_dimensions['C'].width = 35  # Cost Driver
worksheet.column_dimensions['D'].width = 20  # Annual Cost Base
worksheet.column_dimensions['E'].width = 15  # Year 1 Total
worksheet.column_dimensions['F'].width = 15  # Year 2 Total
worksheet.column_dimensions['G'].width = 15  # Year 3 Total
worksheet.column_dimensions['H'].width = 15  # Year 4 Total
worksheet.column_dimensions['I'].width = 15  # Year 5 Total
worksheet.column_dimensions['J'].width = 40  # Notes

# Save the workbook
workbook.save(excel_file)

print(f"SG&A Expenses sheet created in {excel_file}")
