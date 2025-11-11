#!/usr/bin/env python3
"""
Script to copy net total values from corresponding sheets 
to the estimate column in the total list sheet
"""

import openpyxl
from openpyxl import load_workbook
import sys

def find_net_total(worksheet):
    """Find the net total value in a worksheet"""
    net_total = None
    
    # Based on inspection: NET TOTAL is in column 3 (ITEM), value is in column 6 (AMOUNT)
    # Search from bottom up (totals are usually at the end)
    for row_idx in range(worksheet.max_row, max(1, worksheet.max_row - 30), -1):
        # Check column 3 (ITEM column) for "NET TOTAL"
        item_cell = worksheet.cell(row_idx, 3)
        if item_cell.value:
            item_value = str(item_cell.value).strip().upper()
            if 'NET TOTAL' in item_value:
                # Get value from column 6 (AMOUNT column)
                amount_cell = worksheet.cell(row_idx, 6)
                if amount_cell.value is not None:
                    try:
                        net_total = float(amount_cell.value)
                        break
                    except (ValueError, TypeError):
                        pass
    
    return net_total

def find_estimate_column(worksheet):
    """Find the estimate column index"""
    # Search header row (usually row 1)
    for row_idx in range(1, min(5, worksheet.max_row + 1)):
        for col_idx, cell in enumerate(worksheet[row_idx], 1):
            if cell.value:
                cell_value = str(cell.value).lower().strip()
                if 'estimate' in cell_value:
                    return col_idx
    
    return None

def find_name_column(worksheet):
    """Find the column that contains sheet names/item names"""
    # Search header row
    for row_idx in range(1, min(5, worksheet.max_row + 1)):
        for col_idx, cell in enumerate(worksheet[row_idx], 1):
            if cell.value:
                cell_value = str(cell.value).lower().strip()
                if 'name' in cell_value or 'item' in cell_value or 'sheet' in cell_value or 'description' in cell_value:
                    return col_idx
    
    # Default to first column
    return 1

def update_excel(file_path):
    """Main function to update Excel file"""
    try:
        print(f"Loading workbook: {file_path}")
        workbook = load_workbook(file_path)
        
        # Find total list sheet
        total_list_sheet = None
        sheet_names = workbook.sheetnames
        
        print(f"Available sheets: {', '.join(sheet_names)}")
        
        # Look for "total list" sheet
        for sheet_name in sheet_names:
            if 'total' in sheet_name.lower() and 'list' in sheet_name.lower():
                total_list_sheet = workbook[sheet_name]
                print(f"Found total list sheet: {sheet_name}")
                break
        
        if total_list_sheet is None:
            # Try to find it by other names
            for sheet_name in sheet_names:
                if 'list' in sheet_name.lower():
                    total_list_sheet = workbook[sheet_name]
                    print(f"Using sheet as total list: {sheet_name}")
                    break
        
        if total_list_sheet is None:
            print("ERROR: Could not find total list sheet!")
            return False
        
        # Find estimate column
        estimate_col = find_estimate_column(total_list_sheet)
        if estimate_col is None:
            print("ERROR: Could not find estimate column!")
            return False
        
        print(f"Found estimate column at column {estimate_col}")
        
        # Find name column
        name_col = find_name_column(total_list_sheet)
        print(f"Found name column at column {name_col}")
        
        # Create mapping of sheet names to net totals
        sheet_net_totals = {}
        
        for sheet_name in sheet_names:
            if sheet_name == total_list_sheet.title:
                continue
            
            worksheet = workbook[sheet_name]
            net_total = find_net_total(worksheet)
            
            if net_total is not None:
                sheet_net_totals[sheet_name] = net_total
                print(f"Sheet '{sheet_name}': Net Total = {net_total}")
            else:
                print(f"Sheet '{sheet_name}': Could not find net total")
        
        # Update estimate column in total list sheet
        updated_count = 0
        
        # Start from row 2 (assuming row 1 is header)
        for row_idx in range(2, total_list_sheet.max_row + 1):
            name_cell = total_list_sheet.cell(row_idx, name_col)
            estimate_cell = total_list_sheet.cell(row_idx, estimate_col)
            
            if name_cell.value:
                item_name = str(name_cell.value).strip()
                
                # Try to match with sheet names
                matched = False
                for sheet_name, net_total in sheet_net_totals.items():
                    # Exact match
                    if sheet_name.lower() == item_name.lower():
                        estimate_cell.value = net_total
                        updated_count += 1
                        matched = True
                        print(f"Updated row {row_idx}: {item_name} = {net_total}")
                        break
                    # Partial match (sheet name contains item name or vice versa)
                    elif item_name.lower() in sheet_name.lower() or sheet_name.lower() in item_name.lower():
                        estimate_cell.value = net_total
                        updated_count += 1
                        matched = True
                        print(f"Updated row {row_idx}: {item_name} = {net_total} (matched with {sheet_name})")
                        break
                
                if not matched:
                    # Try matching by base name (before first dot)
                    base_name = item_name.split('.')[0] if '.' in item_name else item_name
                    for sheet_name, net_total in sheet_net_totals.items():
                        sheet_base = sheet_name.split('.')[0] if '.' in sheet_name else sheet_name
                        if base_name.lower() == sheet_base.lower():
                            estimate_cell.value = net_total
                            updated_count += 1
                            print(f"Updated row {row_idx}: {item_name} = {net_total} (matched by base name with {sheet_name})")
                            break
        
        # Save the workbook
        print(f"\nSaving workbook... Updated {updated_count} rows.")
        workbook.save(file_path)
        print("Done!")
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    file_path = "e2.xlsx"
    update_excel(file_path)

