# This file handles the code responsible for converting INPUT SHEET 1 TO 2

from tkinter import filedialog
import warnings
from pandas.errors import SettingWithCopyWarning  # Import the warning class

# Suppress the specific pandas warning
warnings.simplefilter(action='ignore', category=SettingWithCopyWarning)


from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from datetime import datetime


import pandas as pd
import numpy as np
from openpyxl import Workbook
import re
import openpyxl


def generate_input_sheet_two(save_directory, input_sheet_one_path):

    if not save_directory:
        save_directory = filedialog.askdirectory(title="Select Folder to Save Payroll Spreadsheet") or ""

    
    # Creating Workbook
    wb = Workbook()
    wb.active

    year = "TEST"  #re.search(r'FY(\d{4})', input_sheet_1_path).group(1) if re.search(r'FY(\d{4})', input_sheet_1_path) else "TEST"
    final_save_path = os.path.join(save_directory,get_save_name(f"CPP FY{year} Input Data Sheet 2 - ALL INPUT DATA"))
    create_presets(wb,input_sheet_one_path,final_save_path)

    wb.save(final_save_path)
    wb.close()
    print("CLOSING FILE NOW")
    return final_save_path

def create_presets(wb, input_sheet_1_path,input_sheet_2_path):
    # setting tab title
    ws = wb.active
    ws.title = f"USER INPUT"

    # TITLE SETTINGS
    ws.row_dimensions[1].height= 28.50
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"INPUT DATA TABLE 2 - ALL INPUT DATA"
    title_cell.font = Font(size=22, bold=True)

    # TIME STAMP SETTINGS
    time_stamp_cell = ws.cell(row=1, column=6)
    apply_color(time_stamp_cell, "maroon")
    time_stamp_cell.value = "Generated on " + datetime.now().strftime("%b-%d-%Y") + " at " + datetime.now().strftime("%H:%M:%S")
    time_stamp_cell.font = Font(size=14, color="FFFFFF", bold=True)  
    ws.merge_cells(start_row=1, start_column=6, end_row=1, end_column=8) 

    # Set width of the first 200 columns and height of the first 200 rows
    for col in range(1, 201):  # Columns A to GR
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 40
    for row in range(3, 201):  # Rows 1 to 200
        ws.row_dimensions[row].height = 25

    print_headers(wb)
    print_employee_and_salary_data(input_sheet_1_path,wb)
    print_funding_string_data(input_sheet_1_path,wb)

    print("PRESETS COMPLETED")

    return

 

#sotring all the headings and its values in a list of tuples
# Nomenclature = {"CONTENT OF HEADING", "ROW NUMBER", "COLUMN NUMBER"}



def print_funding_string_data(input_sheet_1_path, wb):
     # Initialize an empty list to store DataFrames
    df_list = []

    # Extract table boundaries from the Excel file
    tables = extract_table_boundaries(input_sheet_1_path)

    # Loop through each table, process it, and append to the list
    for table in tables:
        # Read the table into a DataFrame
        df = read_excel_table(input_sheet_1_path, "USER INPUT", table)
        
        # Append the DataFrame to the list
        df_list.append(df)

    # Reverse the order of the list
    df_list = df_list[::-1]
    df_list = process_dfs_based_on_identifier_funding_strings(df_list)

    write_tables_to_excel(
        wb= wb,
        sheet_name="USER INPUT",
        tables_list=  df_list,
        start_row=103,
        buffer=3
    )
    return

def process_dfs_based_on_identifier_funding_strings(dfs):
    # Define the desired column lists for each keyword
    desired_columns_map = {
        "funding strings":["Department", "Color","Fund","Account","Sub-Account","Code","Description"]
    }

    processed_dfs = []

    for df in dfs:
        # Extract the identifier from the first column header
        identifier = df.columns[0]
        raw_identifier = str(identifier).lower()  # Column headers as identifier

        # Match the identifier against the keys in the map
        matched_key = next(
            (key for key in desired_columns_map if key in raw_identifier),
            None
        )

        if matched_key:
            # Get the corresponding desired column list
            desired_columns = desired_columns_map[matched_key]

            # Ensure all desired columns are present in the DataFrame
            for col in desired_columns:
                if col not in df.columns:
                    df.loc[:, col] = np.nan  # Add missing columns as empty

            # Reorder columns to include the identifier column at the start
            processed_columns = [identifier] + [col for col in desired_columns if col in df.columns]
            df = df[processed_columns]

            # Add the processed DataFrame to the list
            processed_dfs.append(df)
        else:
            print(f"Warning: No match found for identifier '{raw_identifier}'.")

    return processed_dfs

def print_employee_and_salary_data(input_sheet_1_path,wb):
    # Initialize an empty list to store DataFrames
    df_list = []

    # Extract table boundaries from the Excel file
    tables = extract_table_boundaries(input_sheet_1_path)

    # Loop through each table, process it, and append to the list
    for table in tables:
        # Read the table into a DataFrame
        df = read_excel_table(input_sheet_1_path, "USER INPUT", table)
        
        # Append the DataFrame to the list
        df_list.append(df)

    # Reverse the order of the list
    df_list = df_list[::-1]
    df_list = process_dfs_based_on_identifier(df_list)

    write_tables_to_excel(
        wb= wb,
        sheet_name="USER INPUT",
        tables_list=  df_list,
        start_row=32,
        buffer=3
    )

    return


def print_headers(wb):
    headings = [
        {"heading": "UPDATES", "row": 3, "col": 1},
        {"heading": "FRINGE TABLE", "row": 15, "col": 1},
        {"heading": "TERM LEAVE TABLE", "row": 22, "col": 1},
        {"heading": "CPP EMPLOYEES & SALARY DATA", "row": 29, "col": 1},
        {"heading": "COLAS (Cost of Living Increases)", "row": 96, "col": 1},
        {"heading": "FUNDING STRINGS & ACCOUNTING -- SALARY", "row": 101, "col": 1},
        {"heading": "FUNDING STRINGS & ACCOUNTING -- NON-SALARY", "row": 150, "col": 1},
    ]

    ws = wb.active  # Use the active sheet

    # Define font and alignment for the headings
    heading_font = Font(bold=True, size=12)
    heading_alignment = Alignment(horizontal="center", vertical="center")

    # Add each heading to the workbook
    for entry in headings:
        heading = entry["heading"]
        row = entry["row"]
        col = entry["col"]

        # Write the heading to the specified cell
        cell = ws.cell(row=row, column=col, value=heading)
        cell.font = heading_font
        cell.alignment = heading_alignment

        # Color the entire row up to column 30
        for col_num in range(1, 31):
            apply_color(ws.cell(row=row, column=col_num), "yellow")

    return wb




def process_dfs_based_on_identifier(dfs):
    # Define the desired column lists for each keyword
    desired_columns_map = {
        "salaried employees": [
            "Employee Number", "List Order", "Name", "%FTE",
            "Actual Fringe Rate (104, 131, 136 accts)", "Salary",
            "Fringe Type", "Start Date (if new)", "End Date (if appl.)"
        ],
        "lump sum": [
            "Employee Number", "Teaching", "Name", "%FTE",
            "Home Dept", "Salary", "Fringe Type",
            "Start Date (if new)", "End Date (if appl.)"
        ],
        "undergraduate": [
            "Employee Number", " ", "Name", "%Appt",
            "  ", "Salary", "Fringe Type",
            "Start Date (if new)", "End Date (if appl.)"
        ],
        "graduate": [
            "Employee Number", " ", "Name", "%Appt",
            "  ", "Salary", "Fringe Type",
            "Start Date (if new)", "End Date (if appl.)"
        ],
        "other department": [
            "Employee Number", " ", "Name", "%Appt",
            "  ", "Salary", "Fringe Type",
            "Start Date (if new)", "End Date (if appl.)"
        ]
    }

    processed_dfs = []

    for df in dfs:
        # Extract the identifier from the first column header
        identifier = df.columns[0]
        raw_identifier = str(identifier).lower()  # Column headers as identifier

        # Match the identifier against the keys in the map
        matched_key = next(
            (key for key in desired_columns_map if key in raw_identifier),
            None
        )

        if matched_key:
            # Get the corresponding desired column list
            desired_columns = desired_columns_map[matched_key]

            # Remove rows where the 'Name' column is NaN
            if "Name" in df.columns:
                df = df.dropna(subset=["Name"])

            # Ensure all desired columns are present in the DataFrame
            for col in desired_columns:
                if col not in df.columns:
                    df.loc[:, col] = np.nan  # Add missing columns as empty

            # Reorder columns to include the identifier column at the start
            processed_columns = [identifier] + [col for col in desired_columns if col in df.columns]
            df = df[processed_columns]

            # Add 3 blank rows after each name for Salaried Employees table
            if matched_key == "salaried employees":
                df = add_blank_rows(df,3)

            # Add the processed DataFrame to the list
            processed_dfs.append(df)
        else:
            print(f"Warning: No match found for identifier '{raw_identifier}'.")

    return processed_dfs


def add_blank_rows(df, blank_row_count=3):
    """
    Add a specified number of blank rows after each entry in the 'Name' column.
    
    Parameters:
    - df: DataFrame to modify.
    - blank_row_count: Number of blank rows to add after each row. Default is 3.
    
    Returns:
    - Modified DataFrame with blank rows added.
    """
    if "Name" not in df.columns:
        return df  # Skip if the 'Name' column doesn't exist

    new_rows = []
    for _, row in df.iterrows():
        new_rows.append(row)  # Add the original row
        for _ in range(blank_row_count):
            # Add blank rows
            new_rows.append(pd.Series([np.nan] * len(df.columns), index=df.columns))


    return pd.DataFrame(new_rows, columns=df.columns)



def write_tables_to_excel(wb, sheet_name, tables_list, start_row=32, buffer=3):
    """
    Write a list of DataFrames to an Excel sheet as tables.

    Args:
        wb (Workbook): Excel workbook object.
        sheet_name (str): Name of the sheet to write to.
        tables_list (list): List of DataFrames to write.
        start_row (int): Starting row for the first table.
        buffer (int): Number of blank rows between tables.
    """
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(title=sheet_name)

    ws = wb[sheet_name]
    current_row = start_row

    for idx, df in enumerate(tables_list):
        # Ensure DataFrame is not empty
        if df.empty:
            print(f"Skipping empty DataFrame at index {idx}")
            continue

        # Ensure DataFrame has unique column headers
        if len(df.columns) != len(set(df.columns)):
            print(f"DataFrame at index {idx} has duplicate column headers. Skipping.")
            continue

        # Use the first column header as the base for the table name
        first_header = str(df.columns[0])
        # Sanitize the table name
        sanitized_header = ''.join(e if e.isalnum() else '_' for e in first_header)
        table_name = f"Table_{sanitized_header}_{idx + 1}"

        # Define the range for the table
        start_col = 1  # Writing starts at column A
        end_col = start_col + len(df.columns) - 1
        end_row = current_row + len(df)

        table_range = f"{get_column_letter(start_col)}{current_row}:{get_column_letter(end_col)}{end_row}"

        # Write the DataFrame to the sheet
        for row in dataframe_to_rows(df, index=False, header=True):
            for col_idx, value in enumerate(row, start=start_col):
                ws.cell(row=current_row, column=col_idx, value=value)
            current_row += 1

        # Add buffer
        current_row += buffer

        # Validate the table range before creating the table
        if len(df) > 0 and len(df.columns) > 0:
            try:
                # Create an Excel table
                table = Table(displayName=table_name, ref=table_range)
                style = TableStyleInfo(
                    name="TableStyleMedium9",  # Choose any predefined Excel style
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False,
                )
                table.tableStyleInfo = style
                ws.add_table(table)
                print(f"Successfully added table: {table_name} with range: {table_range}")
            except ValueError as e:
                print(f"Error creating table {table_name}: {e}")
        else:
            print(f"Skipping DataFrame at index {idx} due to invalid dimensions.")

    print(f"All tables written to sheet '{sheet_name}' successfully.")



# This file contains Utility Functions that are used all over the app

import os
import sys
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
import pandas as pd

COLOR_PALETTE = {
    "maroon": "C00000",
    "yellow":"FFC000",

    #Banner Colors
    "lavender":"D5B8EA",
    "pink": "FF8FFF",
    "blue": "BDD7F1",
    "peach": "F9D2BD",
    "green": "D6EDBD",
    "cyan": "D9FFFF",
    
    }

def resource_path(relative_path):
    """ Get the absolute path to a resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def get_save_name(spreadsheet_name):
    """ This Function forms the new file name using the nomencalture which is {filename_currentDate_cuurentTime)

    Args:
        spreadsheet_name (String): Spreadsheet Base name

    Returns:
        String: Formatted name of file path to be saved
    """
    current_datetime = datetime.now().strftime("%b-%d-%Y_%H-%M-%S")
    return f"{spreadsheet_name}_{current_datetime}.xlsx"


def apply_color(cell, color_name):
    """Applies a predefined color to a cell."""
    color_code = COLOR_PALETTE.get(color_name, "FFFFFF")  # Default to white if not found
    fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")
    cell.fill = fill

def read_excel_table(excel_file, sheet_name, table_range):
    """ This function is used to read the values inside any specified range

    Args:
        excel_file (_type_): Excel spreadsheet path
        sheet_name (_type_): Tab name
        table_range (_type_): Range of the table(can be found using extract ramge function)

    Returns:
        _type_: _description_
    """
    # Load Excel workbook
    wb = load_workbook(excel_file, read_only=True)
    try:
        # Select the worksheet
        ws = wb[sheet_name]

        # Read data from the specified table range into a list of lists
        table_data = []
        for row in ws[table_range]:
            table_data.append([cell.value for cell in row])

        # Convert the list of lists to a DataFrame
        df = pd.DataFrame(table_data[1:], columns=table_data[0])
    finally:
        # Ensure the workbook is closed
        wb.close()

    return df


def extract_table_boundaries(file_path):
    """ Returns table boundaries from a given sheet

    Args:
        file_path (_type_):Path of the Spreadsheet

    Returns:
        _type_: A List of Table Boundaries
    """
    # Load the workbook
    wb = load_workbook(file_path, data_only=True)

    # List of Table Boundaries
    table_ranges= []

    # Loop through all worksheets
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]

        # Check for tables in the sheet
        if hasattr(sheet, 'tables') and sheet.tables:
            for table_name in sheet.tables.keys():
                table = sheet.tables[table_name]  # Retrieve the table object
                if hasattr(table, 'ref'):  # Safely access ref attribute
                    table_range = table.ref
                    table_ranges.append(table_range)

    return table_ranges