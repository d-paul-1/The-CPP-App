# This file contains functions for various processing tasks.
import pandas as pd
import os
import requests

from pdf2docx import Converter
from docx import Document

#funtion to construct the payschedule url from the year
def construct_url(year):
    url = f"https://uwservice.wisconsin.edu/docs/publications/pay-bw-calendar-{year}.pdf"
    
    return url , "url constructed successfully" , "green"


# function to download pdf from url
def download_pdf(url):   
    # Get the current directory of the script (which is the src directory)
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Navigate one level up (to the parent directory of 'src')
    parent_dir = os.path.dirname(script_dir)

    # Define the directory name to save PDFs (in the parent directory)
    save_directory = os.path.join(parent_dir, "pay_schedules")

    # Create the directory if it doesn't exist
    os.makedirs(save_directory, exist_ok=True)

    # Extract the filename from the URL
    filename = os.path.basename(url)
    # Define the full path to save the PDF
    file_path = os.path.join(save_directory, filename)

    response = requests.get(url)
    if response.status_code == 200:
        with open(file_path, 'wb') as f:
            f.write(response.content)
        msg = f"PDF downloaded successfully: {file_path}"
        color="green"
    else:
        msg = f"Failed to download PDF from {url}. Status code: {response.status_code}"
        color = "red"

    return file_path,msg,color  # Return the full path of the downloaded PDF


# function to convert pdf to a Word document
def pdf_to_word(pdf):
    try:
        word_file = pdf.replace('.pdf', '.docx')
        cv = Converter(pdf)
        cv.convert(word_file, start=0, end=None)
        cv.close()
        return word_file, "Word document created successfully.", "green"
    except Exception as e:
        return None, f"Error converting PDF to Word: {str(e)}", "red"


# function to convert a Word document to a pandas DataFrame
def word_to_df(docx_file):
    try:
        doc = Document(docx_file)

        pay_periods_list = []
        pay_period_dates_list = []
        seen = set()

        for table in doc.tables:
            header_row_index = -1
            for i, row in enumerate(table.rows):
                cell_texts = [cell.text for cell in row.cells]
                if 'Pay Period' in cell_texts and 'Pay Period Dates' in cell_texts:
                    header_row_index = i
                    header_cells = cell_texts
                    break

            if header_row_index != -1:
                pay_period_index = header_cells.index('Pay Period')
                pay_period_dates_index = header_cells.index('Pay Period Dates')

                for row in table.rows[header_row_index + 1:]:
                    cells = row.cells
                    pay_period = cells[pay_period_index].text.strip()
                    pay_period_dates = cells[pay_period_dates_index].text.strip()
                    if pay_period and pay_period_dates:
                        if (pay_period, pay_period_dates) not in seen:
                            pay_periods_list.append(pay_period)
                            pay_period_dates_list.append(pay_period_dates)
                            seen.add((pay_period, pay_period_dates))

        df = pd.DataFrame({
            "Pay Period": pay_periods_list,
            "Pay Period Dates": pay_period_dates_list
        })
        return df, "DataFrame created successfully.", "green"

    except Exception as e:
        return None, f"Error extracting data from Word document: {str(e)}", "red"

# function to merge 2 pandas data frames based on fiscal year requiremnts
def merge_dfs(df_previous_year, df_current_year):
    """
    Merges the payschedules of the previous year and current year.
    
    Parameters:
    - df_previous_year: DataFrame containing the payschedule of the previous year.
    - df_current_year: DataFrame containing the payschedule of the current year.
    
    Returns:
    - cleaned_df: A DataFrame that is the merged and cleaned result.
    - error_message: A string indicating if an error occurred.
    - color: A string representing the color for display (e.g., "red" for error, "green" for success).
    """
    # Initialize error message and color
    error_message = ""
    color = "green"  # Default to green if no errors

    # Code which grabs the first half of the payschedule from the previous year
    jul_index_previous_year = df_previous_year[df_previous_year["Pay Period"].str.contains("JUL A")].index
    jul_index_previous_year = jul_index_previous_year[0] if len(jul_index_previous_year) > 0 else None

    if jul_index_previous_year is not None:
        df_previous_year_sliced = df_previous_year.iloc[jul_index_previous_year:]
    else:
        error_message += "No JUL A pay period found in the previous year's DataFrame.\n"
        color = "red"  # Set color to red for error
        df_previous_year_sliced = pd.DataFrame()  # Return an empty DataFrame to avoid further processing

    # Code which grabs the second half of the payschedule from the current year
    jul_index_current_year = df_current_year[df_current_year["Pay Period"].str.contains("JUL A")].index
    jul_index_current_year = jul_index_current_year[0] if len(jul_index_current_year) > 0 else None

    if jul_index_current_year is not None:
        df_current_year_sliced = df_current_year.iloc[1:jul_index_current_year]
    else:
        error_message += "No JUL A pay period found in the current year's DataFrame.\n"
        color = "red"  # Set color to red for error
        df_current_year_sliced = pd.DataFrame()  # Return an empty DataFrame to avoid further processing
    
    # If both slices are empty, we need to indicate a severe error
    if df_previous_year_sliced.empty and df_current_year_sliced.empty:
        error_message += "Both previous and current year DataFrames are empty after slicing.\n"
        color = "red"
    
    # Code to merge and clean the payschedules
    merged_df = pd.concat([df_previous_year_sliced, df_current_year_sliced], ignore_index=True)
    cleaned_df = merged_df.drop_duplicates(subset="Pay Period Dates", keep="first")
    
    return cleaned_df, error_message.strip(), color


# function to calculate number of pay periods from merged pandas data 
def get_payperiod(pay_schedule):
    return 0

# Function to return df for the Automatic choice in Option Menu
def process_automatic(year):
    url = construct_url(year)
    pdf = download_pdf(url)
    word = pdf_to_word(pdf)
    df = word_to_df(word)
    return df
