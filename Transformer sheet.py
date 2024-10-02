import pandas as pd
import re
from pathlib import Path
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def transform_excel(input_file, output_file, cost_file, date_file_path, gallons_file):
    try:
        input_file_path = Path(input_file)
        if input_file_path.suffix == '.xls':
            # Convert .xls to .xlsx
            df = pd.read_excel(input_file, engine='xlrd')
            converted_file_path = input_file_path.with_suffix('.xlsx')
            df.to_excel(converted_file_path, index=False)
            input_file = converted_file_path

        # Read the converted .xlsx file using openpyxl engine
        df = pd.read_excel(input_file, engine='openpyxl')
        logging.info(f"Shape of the DataFrame after loading: {df.shape}")

        # Read the date from the text file
        with open(date_file_path, 'r') as file:
            date_text = file.readline().strip()

        logging.info(f"Date read from file: {date_text}")

        # Assuming date_text is '20240711', keep it as 'YYYYMMDD'
        try:
            # Validate date format
            datetime.strptime(date_text, '%Y%m%d')
            formatted_date = date_text  # Keep in YYYYMMDD format
        except ValueError as e:
            logging.error(f"Error: {e}")
            logging.error("Check if the date in the file is in the expected format 'YYYYMMDD'.")
            return

        # Modify the output file name to include the formatted date
        output_file = Path(output_file)
        output_file = output_file.with_name(f"{output_file.stem}_{formatted_date}{output_file.suffix}")

        # Read the cost from the text file
        with open(cost_file, 'r') as file:
            cost_text = file.readline().strip()
        
        # Read the gallons from the text file
        with open(gallons_file, 'r') as file:
            gallons_text = file.readline().strip()

        # Function to extract only the numeric parts from a string
        def extract_numbers(s):
            return ''.join(filter(str.isdigit, str(s)))

        # Updated function to split the value in column K and correctly extract parts
        def split_k_value(value):
            # Find the first sequence of digits (before "GRDY") and the second sequence (after "/")
            match = re.match(r"(\d+).*GRDY/(\d+)", str(value))
            if match:
                num1 = int(match.group(1))  # Part before "GRDY"
                num2 = int(match.group(2))  # Part after "/"
            else:
                num1, num2 = None, None
            return num1, num2

        # Apply the updated split function to column K (index 10)
        df['K_num1'], df['K_num2'] = zip(*df.iloc[:, 10].apply(split_k_value))
        
        # Assign the lower value to F and the higher value to G
        df['F'] = df[['K_num1', 'K_num2']].min(axis=1)
        df['G'] = df[['K_num1', 'K_num2']].max(axis=1)

        # Create new DataFrame with the required transformations
        new_df = pd.DataFrame()
        new_df['A'] = df.iloc[:, 6]    # Column G to Column A
        new_df['B'] = ''               # Column B is blank
        new_df['C'] = 8                # Column C is 8 in every row
        new_df['D'] = df.iloc[:, 9]    # Column J to Column D
        new_df['E'] = df.iloc[:, 3]    # Column D to Column E
        new_df['F'] = df['F']          # Lower value of the split result to Column F
        new_df['G'] = df['G']          # Higher value of the split result to Column G
        new_df['H'] = df.iloc[:, 11].str.split('/').str[1]  # Low value of Column L to Column H
        new_df['I'] = df.iloc[:, 11].str.split('/').str[0]  # High value of Column L to Column I
        new_df['J'] = formatted_date  # Keep date in 'YYYYMMDD' format
        new_df['K'] = df.iloc[:, 1].apply(extract_numbers)  # Column B to Column K, extracting only numbers
        new_df['L'] = str(cost_text)   # Column L is filled with a value read from a text file, formatted as string
        new_df['M'] = ''               # Column M is blank
        new_df['N'] = ''               # Column N is blank
        new_df['O'] = df.iloc[:, 14]   # Column O to Column O
        new_df['P'] = df.iloc[:, 13]   # Column N to Column P
        new_df['Q'] = ''               # Column Q is blank
        new_df['R'] = 'BR'             # Column R is "BR" in every row
        new_df['S'] = df.iloc[:, 7].apply(lambda x: 'OH' if x == 'OV' else 'URD' if x == 'PD' else x)  # Column H with transformation to Column S
        new_df['T'] = ''               # Column T is blank
        new_df['U'] = ''               # Column U is blank
        new_df['V'] = ''               # Column V is blank
        new_df['W'] = df.iloc[:, 8]    # Column I to Column W
        new_df['X'] = ''               # Column X is blank
        new_df['Y'] = ''               # Column Y is blank
        new_df['Z'] = ''               # Column Z is blank
        new_df['AA'] = ''              # Column AA is blank
        new_df['AB'] = ''              # Column AB is blank
        new_df['AC'] = ''              # Column AC is blank
        new_df['AD'] = ''              # Column AD is blank
        new_df['AE'] = ''              # Column T to Column AE
        new_df['AF'] = ''              # Column AF is blank
        new_df['AG'] = str(gallons_text) # Column AG is filled with a value read from a text file, formatted as string
        new_df['AH'] = ''              # Column AH is blank
        new_df['AI'] = df.iloc[:, 22]  # Column W to Column AI
        new_df['AJ'] = ''              # Column AJ is blank
        new_df['AK'] = 'N'             # Column AK is "N" in every row

        # Reorder columns according to the required order
        columns_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK']
        new_df = new_df[columns_order]
        
        logging.info(f"Shape of the new DataFrame: {new_df.shape}")

        # Export to CSV without header
        new_df.to_csv(output_file, index=False, header=False)

        logging.info(f"Number of rows in the input file (excluding header): {len(df)}")
        logging.info(f"Number of rows in the output file: {len(new_df)}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")

# Example usage
input_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\ERMCO-TRANSFORMER-DATA082124.xls"  # Path to the input Excel file
output_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\NewErmcoData.csv"  # Base path for the output CSV file
cost_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\cost.txt"
date_file_path = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\date.txt"
gallons_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\gallons.txt"

transform_excel(input_file, output_file, cost_file, date_file_path, gallons_file)
