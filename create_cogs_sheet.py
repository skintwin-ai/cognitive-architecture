import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Color
from openpyxl.utils import get_column_letter
import os

from financial_model_config import load_existing_workbook, parse_output_args, resolve_output_path, save_workbook


def build_sheet(workbook, excel_file):
    # Create COGS and direct costs data
    cogs_data = {
        'Cost Category': [
            # Raw Materials
            'Raw Materials', 'Raw Materials', 'Raw Materials', 'Raw Materials', 'Raw Materials',
            # Packaging Materials
            'Packaging Materials', 'Packaging Materials', 'Packaging Materials', 'Packaging Materials',
            # Direct Labor
            'Direct Labor', 'Direct Labor', 'Direct Labor',
            # Manufacturing Overhead
            'Manufacturing Overhead', 'Manufacturing Overhead', 'Manufacturing Overhead', 'Manufacturing Overhead',
            # Contract Manufacturing
            'Contract Manufacturing', 'Contract Manufacturing',
            # Quality Control
            'Quality Control', 'Quality Control', 'Quality Control',
            # Freight & Logistics (Inbound)
            'Freight & Logistics (Inbound)', 'Freight & Logistics (Inbound)', 'Freight & Logistics (Inbound)',
        ],
        'Cost Item': [
            # Raw Materials
            'Active Ingredients', 'Base Ingredients', 'Fragrances', 'Preservatives', 'Other Additives',
            # Packaging Materials
            'Primary Packaging', 'Secondary Packaging', 'Labels & Printing', 'Shipping Materials',
            # Direct Labor
            'Production Line Workers', 'Quality Inspectors', 'Material Handlers',
            # Manufacturing Overhead
            'Production Supervision', 'Manufacturing Utilities', 'Equipment Depreciation', 'Production Supplies',
            # Contract Manufacturing
            'Contract Manufacturing Fees', 'Formulation Services',
            # Quality Control
            'Raw Material Testing', 'In-Process Testing', 'Finished Product Testing',
            # Freight & Logistics (Inbound)
            'Raw Material Freight', 'Import Duties & Customs', 'Inbound Logistics Management',
        ],
        'Cost Driver': [
            # Raw Materials
            'Cost per unit × units produced', 'Cost per unit × units produced', 'Cost per unit × units produced', 'Cost per unit × units produced', 'Cost per unit × units produced',
            # Packaging Materials
            'Cost per unit × units produced', 'Cost per unit × units produced', 'Cost per unit × units produced', 'Cost per unit × units produced',
            # Direct Labor
            'Hourly rate × hours per unit × units produced', 'Hourly rate × hours per unit × units produced', 'Hourly rate × hours per unit × units produced',
            # Manufacturing Overhead
            'Salary × headcount', 'Usage rates × utility prices', 'Equipment value ÷ useful life', 'Cost per production run × number of runs',
            # Contract Manufacturing
            'Fee per unit × units outsourced', 'Fixed fee per formulation × number of formulations',
            # Quality Control
            'Cost per test × number of tests', 'Cost per test × number of tests', 'Cost per test × number of tests',
            # Freight & Logistics (Inbound)
            'Freight cost per shipment × number of shipments', 'Percentage of imported material value', 'Fixed monthly cost',
        ],
        'Unit Cost': [
            # Raw Materials
            '2.50', '0.75', '0.50', '0.25', '0.35',
            # Packaging Materials
            '1.20', '0.60', '0.25', '0.15',
            # Direct Labor
            '0.80', '0.30', '0.20',
            # Manufacturing Overhead
            '0.40', '0.25', '0.35', '0.15',
            # Contract Manufacturing
            '1.50', '0.30',
            # Quality Control
            '0.20', '0.15', '0.25',
            # Freight & Logistics (Inbound)
            '0.18', '0.12', '0.10',
        ],
        'Annual Volume': [
            # Raw Materials
            '=Assumptions!C15*100000', '=Assumptions!C15*100000', '=Assumptions!C15*100000', '=Assumptions!C15*100000', '=Assumptions!C15*100000',
            # Packaging Materials
            '=Assumptions!C15*100000', '=Assumptions!C15*100000', '=Assumptions!C15*100000', '=Assumptions!C15*100000',
            # Direct Labor
            '=Assumptions!C15*100000', '=Assumptions!C15*100000', '=Assumptions!C15*100000',
            # Manufacturing Overhead
            '12', '12', '1', '52',
            # Contract Manufacturing
            '=Assumptions!C15*100000*0.2', '10',
            # Quality Control
            '=Assumptions!C15*100000*0.1', '=Assumptions!C15*100000*0.05', '=Assumptions!C15*100000*0.1',
            # Freight & Logistics (Inbound)
            '52', '=Assumptions!C15*100000*0.3', '12',
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
            # Raw Materials
            'Key functional ingredients with specific skin benefits', 'Foundation ingredients that form the product base', 'Scents and aromatic compounds', 'Ingredients that extend shelf life', 'Colorants, thickeners, and other specialty ingredients',
            # Packaging Materials
            'Bottles, jars, tubes, pumps that directly contain product', 'Boxes, cartons, and outer packaging', 'Product information, branding, and regulatory labeling', 'Materials for safe product shipping',
            # Direct Labor
            'Workers directly involved in product manufacturing', 'Staff dedicated to quality inspection during production', 'Workers handling materials and inventory on production floor',
            # Manufacturing Overhead
            'Production managers and supervisors', 'Electricity, water, and other utilities for production', 'Allocation of equipment cost over its useful life', 'Gloves, cleaning supplies, and other production consumables',
            # Contract Manufacturing
            'Fees paid to third-party manufacturers for production', 'Costs for external formulation development services',
            # Quality Control
            'Laboratory testing of incoming raw materials', 'Testing during the production process', 'Final product quality verification testing',
            # Freight & Logistics (Inbound)
            'Transportation costs for raw material delivery', 'Taxes and fees for imported materials', 'Costs for managing inbound logistics operations',
        ]
    }

    # Create empty columns for Year 1-5 Total
    for year in ['Year 1 Total', 'Year 2 Total', 'Year 3 Total', 'Year 4 Total', 'Year 5 Total']:
        cogs_data[year] = [''] * len(cogs_data['Cost Item'])

    # Create a Pandas DataFrame
    cogs_df = pd.DataFrame(cogs_data)

    # Check if the sheet already exists and remove it if it does
    if 'COGS' in workbook.sheetnames:
        idx = workbook.sheetnames.index('COGS')
        workbook.remove(workbook.worksheets[idx])

    # Create a new sheet
    worksheet = workbook.create_sheet('COGS')

    # Write the header row
    for col_idx, column_name in enumerate(cogs_df.columns, 1):
        worksheet.cell(row=2, column=col_idx).value = column_name

    # Write the data rows
    for row_idx, row in enumerate(cogs_df.itertuples(index=False), 3):
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
    worksheet['A1'] = 'Cost of Goods Sold (COGS) and Direct Costs'
    worksheet['A1'].font = title_font
    worksheet.merge_cells('A1:I1')
    worksheet['A1'].alignment = Alignment(horizontal='center')

    # Format header row
    for col_idx, value in enumerate(cogs_df.columns, 1):
        cell = worksheet.cell(row=2, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

    # Add formulas for Year 1-5 Total columns
    for row_idx in range(3, len(cogs_data['Cost Item']) + 3):
        # Year 1 Total formula
        worksheet.cell(row=row_idx, column=6).value = f'=D{row_idx}*E{row_idx}'
    
        # Year 2-5 Total formulas with growth factors from Assumptions
        # Year 2
        worksheet.cell(row=row_idx, column=7).value = f'=F{row_idx}*(1+Assumptions!D14)'
    
        # Year 3
        worksheet.cell(row=row_idx, column=8).value = f'=G{row_idx}*(1+Assumptions!E14)'
    
        # Year 4
        worksheet.cell(row=row_idx, column=9).value = f'=H{row_idx}*(1+Assumptions!F14)'
    
        # Year 5
        worksheet.cell(row=row_idx, column=10).value = f'=I{row_idx}*(1+Assumptions!G14)'

    # Format data rows
    last_category = None
    for row_idx in range(3, len(cogs_data['Cost Item']) + 3):
        category_cell = worksheet.cell(row=row_idx, column=1)
        cost_item_cell = worksheet.cell(row=row_idx, column=2)
    
        # Apply border to all cells in the row
        for col_idx in range(1, 12):  # Columns A through K
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
        
            # Format numbers in calculation columns
            if col_idx >= 4 and col_idx <= 10:  # Columns D through J
                cell.number_format = '#,##0.00'
    
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

    # Add subtotals by category
    last_row = len(cogs_data['Cost Item']) + 3
    row_idx = last_row + 1

    # Add a subtotal row
    worksheet.cell(row=row_idx, column=1).value = 'Total COGS'
    worksheet.cell(row=row_idx, column=1).font = Font(bold=True)
    worksheet.merge_cells(f'A{row_idx}:C{row_idx}')

    # Add subtotal formulas
    for col_idx in range(6, 11):  # Columns F through J
        cell = worksheet.cell(row=row_idx, column=col_idx)
        col_letter = get_column_letter(col_idx)
        cell.value = f'=SUM({col_letter}3:{col_letter}{last_row})'
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
        cell.number_format = '#,##0.00'

    # Add a row for COGS per unit
    row_idx += 2
    worksheet.cell(row=row_idx, column=1).value = 'COGS per Unit'
    worksheet.cell(row=row_idx, column=1).font = Font(bold=True, italic=True)
    worksheet.merge_cells(f'A{row_idx}:C{row_idx}')

    # Add COGS per unit formulas (these will be linked to volume data later)
    for col_idx in range(6, 11):  # Columns F through J
        cell = worksheet.cell(row=row_idx, column=col_idx)
        col_letter = get_column_letter(col_idx)
        cell.value = f'={col_letter}{row_idx-2}/100000'  # Placeholder formula, will be updated when Volume data is available
        cell.font = Font(bold=True, italic=True)
        cell.number_format = '#,##0.00'

    # Add a row for percentage of revenue
    row_idx += 2
    worksheet.cell(row=row_idx, column=1).value = 'COGS as % of Revenue'
    worksheet.cell(row=row_idx, column=1).font = Font(bold=True, italic=True)
    worksheet.merge_cells(f'A{row_idx}:C{row_idx}')

    # Add percentage formulas (these will be linked to Revenue sheet later)
    for col_idx in range(6, 11):  # Columns F through J
        cell = worksheet.cell(row=row_idx, column=col_idx)
        col_letter = get_column_letter(col_idx)
        cell.value = f'={col_letter}{row_idx-4}/100000'  # Placeholder formula, will be updated when Revenue sheet is created
        cell.font = Font(bold=True, italic=True)
        cell.number_format = '0.0%'

    # Adjust column widths
    worksheet.column_dimensions['A'].width = 25  # Cost Category
    worksheet.column_dimensions['B'].width = 25  # Cost Item
    worksheet.column_dimensions['C'].width = 35  # Cost Driver
    worksheet.column_dimensions['D'].width = 15  # Unit Cost
    worksheet.column_dimensions['E'].width = 15  # Annual Volume
    worksheet.column_dimensions['F'].width = 15  # Year 1 Total
    worksheet.column_dimensions['G'].width = 15  # Year 2 Total
    worksheet.column_dimensions['H'].width = 15  # Year 3 Total
    worksheet.column_dimensions['I'].width = 15  # Year 4 Total
    worksheet.column_dimensions['J'].width = 15  # Year 5 Total
    worksheet.column_dimensions['K'].width = 40  # Notes


def main(output_path: str | None = None):
    excel_file = resolve_output_path(output_path)
    workbook = load_existing_workbook(excel_file, ['Assumptions'])
    build_sheet(workbook, str(excel_file))
    save_workbook(workbook, excel_file)
    print(f'COGS sheet created in {excel_file}')
    return excel_file


if __name__ == '__main__':
    args = parse_output_args('Create the COGS sheet for the SkinTwin financial model workbook.')
    main(args.output)
