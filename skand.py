import argparse
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from openpyxl.styles import NamedStyle

def convert_datetime(column):
    #Converts datetime strings to objects
    return pd.to_datetime(column, errors='coerce', dayfirst=True) # ensures the format is dd/mm/yyyy before converting

def process_csv(file_path, threshold_date=None):
    # Read CSV file
    df = pd.read_csv(file_path)
    # print(df.head())  # Check the original data
    
    # Convert the first column (assumed to contain datetime strings) to datetime format
    # -c arg to specifiy what column it uses, or it should auto find the column using the heading value?
    df.iloc[:, 0] = convert_datetime(df.iloc[:, 0])
    # print(df.head())  # Check after conversion
    
    # Filter DataFrame if time threshold argument is provided
    if threshold_date:
        threshold_datetime = pd.to_datetime(threshold_date, format='%Y-%m-%d')
        df['Time'] = pd.to_datetime(df['Time'], dayfirst=True)
        # Filter the DataFrame to include rows on or after the threshold date.
        df = df[df['Time'] >= threshold_datetime]
     # Convert the DataFrame to an Excel file for formatting purposes (need to work on CSV conversion)
        base_name = os.path.splitext(file_path)[0]
        output_xlsx_file = f"{base_name}_CLEAN.xlsx" # saves with appendix (maybe overwrite?)
        wb = Workbook()
        ws = wb.active

    # Create a named style for datetime cells
    date_time_style = NamedStyle(name='datetime', number_format='YYYY-MM-DD HH:MM:SS')

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            # If the cell contains a datetime, apply the named style
            if isinstance(value, pd.Timestamp):
                cell.style = date_time_style

    wb.save(output_xlsx_file)
    Ascii()
    print(f"Excel file saved successfully. Output saved to {output_xlsx_file}")

def Ascii(): 
    Ascii = print(r"""
   /$$$$$$  /$$   /$$  /$$$$$$  /$$   /$$ /$$$$$$$ 
 /$$__  $$| $$  /$$/ /$$__  $$| $$$ | $$| $$__  $$
| $$  \__/| $$ /$$/ | $$  \ $$| $$$$| $$| $$  \ $$
|  $$$$$$ | $$$$$/  | $$$$$$$$| $$ $$ $$| $$  | $$
 \____  $$| $$  $$  | $$__  $$| $$  $$$$| $$  | $$
 /$$  \ $$| $$\  $$ | $$  | $$| $$\  $$$| $$  | $$
|  $$$$$$/| $$ \  $$| $$  | $$| $$ \  $$| $$$$$$$/
 \______/ |__/  \__/|__/  |__/|__/  \__/|_______/                             
        """)


def main():
    parser = argparse.ArgumentParser(description='A simple DFIR tool to format reports generated from disk images')
    parser.add_argument('-f', '--file', help='Specify the CSV file to process')
    parser.add_argument('-t', '--threshold', help='Specify the threshold date for filtering records (format: YYYY-MM-DD)')
    args = parser.parse_args()

    if args.file:
        process_csv(args.file, args.threshold)
    else:
        print(Ascii())
        parser.print_help()

if __name__ == "__main__":
    main()


# Future improvements: 
# Detect day / time or -c arg for specific column
# time threshold range e.g. 2024-01-01 - 2024-01-30
