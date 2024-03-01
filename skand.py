import argparse
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import NamedStyle
import os

def convert_datetime(column):
    # Converts datetime strings to objects
    return pd.to_datetime(column, errors='coerce', dayfirst=True)  # ensures the format is dd/mm/yyyy before converting

def process_csv(file_path, threshold_date=None):
    try:
        # Read the CSV file
        df = pd.read_csv(file_path)
        
        # Convert the first column to datetime format
        df.iloc[:, 0] = convert_datetime(df.iloc[:, 0])
    
        # Filter DataFrame if time threshold argument is provided
        if threshold_date:
            threshold_datetime = pd.to_datetime(threshold_date, infer_datetime_format=True)
            # Filter the DataFrame to include rows on or after the threshold date using the first column
            df = df[df.iloc[:, 0] >= threshold_datetime]
        
        # Convert the DataFrame to an Excel file for formatting purposes
        base_name = os.path.splitext(file_path)[0]
        output_xlsx_file = f"{base_name}_CLEAN.xlsx"
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
        print(f"Excel file saved successfully. Output saved to {output_xlsx_file}")

    except Exception as e:
        print(f"Error processing file: {e}")
        Ascii()

def Ascii(): 
    print(r"""
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
        Ascii()
        parser.print_help()

if __name__ == "__main__":
    main()


    
# Future improvements: 
# Detect day / time or -c arg for specific column
# time threshold range e.g. 2024-01-01 - 2024-01-30
# include directory instead of single file (-d)