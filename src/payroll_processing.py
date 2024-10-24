# This file contains functions for various processing tasks.
import pandas as pd
import os
import requests

#funtion to construct the payschedule url from the year
def construct_url(year):
    url = f"https://uwservice.wisconsin.edu/docs/publications/pay-bw-calendar-{year}.pdf"
    return url


# function to download pdf from url
def download_pdf(url):
    from payroll_spreadsheet import PayrollSpreadsheet 

    # payroll_instance 
    ps = PayrollSpreadsheet()   
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
        # self.update_status(f"PDF downloaded successfully: {file_path}")
    else:
        ps.display_msg(f"Failed to download PDF from {url}. Status code: {response.status_code}")

    return file_path  # Return the full path of the downloaded PDF


# function to convert pdf to a pandas data frame
def pdf_to_word(pdf):
    return 0

def word_to_df(docx_file):
    return 0 

# function to merge 2 pandas data frames based on fiscal year requiremnts

def merge_dfs(year1,year2):
    return 0

# function to calculate number of pay periods from merged pandas data 
def get_payperiod(pay_schedule):
    return 0

# Function to return df for the Automatic choice in Option Menu
def process_automatic(year):
    url = construct_url(year)
    pdf = download_pdf(url)
    df = pdf_to_df(pdf)
    return df
