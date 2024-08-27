from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re


def OpenBrowseDialog(prompt):
    Tk().withdraw()  
    file_path = askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx")])
    return file_path

def find_column_index(sheet, column_name):
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(row=1, column=col).value
        if cell_value == column_name:
            return col
    return None

def FindFirstEmptyRow(sheet):
    for row in range(1, sheet.max_row + 1):
        if not any(sheet.cell(row=row, column=col).value for col in range(1, sheet.max_column + 1)):
            return row
    return sheet.max_row + 1  # If no empty row is found, append at the end

def CopyColumns(source_file, destination_file, column_mapping, sourceSheet, destinationSheet):
    source_wb = load_workbook(source_file)
    source_sheet = sourceSheet

    destination_wb = load_workbook(destination_file)
    destination_sheet = destinationSheet

    # Find the columns in the Source and Destination sheets based on the provided mapping
    source_columns = {src_col: find_column_index(source_sheet, src_col) for src_col in column_mapping.keys()}
    missing_source_columns = [col for col, idx in source_columns.items() if idx is None]

    destination_columns = {dest_col: find_column_index(destination_sheet, dest_col) for dest_col in column_mapping.values()}
    missing_destination_columns = [col for col, idx in destination_columns.items() if idx is None]

     # Report missing columns
    if missing_source_columns or missing_destination_columns:
        if missing_source_columns:
            print(f"Source columns not found: {', '.join(missing_source_columns)}")
        if missing_destination_columns:
            print(f"Destination columns not found: {', '.join(missing_destination_columns)}")
        return

    # Find the first empty row in the Destination sheet
    start_row = FindFirstEmptyRow(destination_sheet)

    # Iterate through the rows in the Source sheet and copy data
    for row in range(2, source_sheet.max_row + 1):  # Start from 2 if there's a header row
        data_to_copy = {dest_col: source_sheet.cell(row=row, column=source_columns[src_col]).value
                        for src_col, dest_col in column_mapping.items()}

        # Write the data to the Destination sheet
        for dest_col, value in data_to_copy.items():
            destination_sheet.cell(row=start_row, column=destination_columns[dest_col], value=value)
        start_row += 1  # Move to the next row

    # Save the updated Destination workbook
    destination_wb.save(destination_file)
    print("Data copied successfully!")



if __name__ == "__main__":
    sourceExcel = OpenBrowseDialog("Select the source Excel file")
    destinationExcel = OpenBrowseDialog("Select the destination Excel file")

    sourceSheet = ""
    destinationSheet = ""

    column_mapping = {
    'id': 'idntificatioNo',
    'name': 'Contant-name',
    'address': 'Contact-address'
    }

    CopyColumns(sourceExcel, destinationExcel, column_mapping, sourceSheet, destinationSheet)
