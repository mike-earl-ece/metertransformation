import pandas as pd
import re
from pathlib import Path

def transform_excel(input_file, output_file, cost_file, date_file_path):
    # Convert .xls to .xlsx if necessary
    input_file_path = Path(input_file)
    if input_file_path.suffix == '.xls':
        # Convert .xls to .xlsx
        df = pd.read_excel(input_file, engine='xlrd')
        converted_file_path = input_file_path.with_suffix('.xlsx')
        df.to_excel(converted_file_path, index=False)
        input_file = converted_file_path

    # Read the converted .xlsx file using openpyxl engine
    df = pd.read_excel(input_file, engine='openpyxl')
    
    # Remove the header
    df = df.iloc[1:]

    # Read the date from the text file
    with open(date_file_path, 'r') as file:
        date_text = file.readline().strip()

    # Read the cost from the text file
    with open(cost_file, 'r') as file:
        cost_text = file.readline().strip()

    # Function to extract only the numeric parts from a string
    def extract_numbers(s):
        return ''.join(filter(str.isdigit, str(s)))

    # Function to split the value in column K
    def split_k_value(value):
        parts = re.split(r'GRDY|/', str(value))
        if len(parts) < 2:
            return None, None
        num1 = parts[0]
        num2 = parts[1]
        return num1, num2

    # Create new DataFrame with the required transformations
    new_df = pd.DataFrame()
    new_df['A'] = df.iloc[:, 6]    # Column G to Column A
    new_df['B'] = ''               # Column B is blank
    new_df['C'] = 8                # Column C is 8 in every row
    new_df['D'] = df.iloc[:, 9]    # Column J to Column D
    new_df['E'] = df.iloc[:, 3]    # Column D to Column E
    
    # Split column K and assign to F and G
    df['K_num1'], df['K_num2'] = zip(*df.iloc[:, 10].apply(split_k_value))
    new_df['F'] = df['K_num2']  # Number after "/" to Column F
    new_df['G'] = df['K_num1']  # Number before "GRDY" to Column G
    
    new_df['H'] = df.iloc[:, 11].str.split('/').str[1]  # Low value of Column L to Column H
    new_df['I'] = df.iloc[:, 11].str.split('/').str[0]  # High value of Column L to Column I
    new_df['J'] = str(date_text)        # Column J is filled with a value read from a text file
    new_df['K'] = df.iloc[:, 1].apply(extract_numbers)  # Column B to Column K, extracting only numbers
    new_df['L'] = cost_text        # Column L is filled with a value read from a text file
    new_df['M'] = ''               # Column M is blank
    new_df['N'] = ''               # Column N is blank
    new_df['O'] = df.iloc[:, 14]   # Column O to Column O
    new_df['P'] = df.iloc[:, 13]   # Column N to Column P
    new_df['Q'] = ''               # Column Q is blank
    new_df['R'] = 'BR'             # Column R is "BR" in every row
    new_df['S'] = df.iloc[:, 7].apply(lambda x: 'OH' if x == 'OV' else 'URD' if x == 'PD' else x)  # Column H with transformation to Column S
    new_df['T'] = ''               # Column T is blank
    new_df['U'] = ''               # Column U is blank
    new_df['V'] = df.iloc[:, 12]   # Column M to Column V
    new_df['W'] = df.iloc[:, 8]    # Column I to Column W
    new_df['X'] = ''               # Column X is blank
    new_df['Y'] = ''               # Column Y is blank
    new_df['Z'] = ''               # Column Z is blank
    new_df['AA'] = ''              # Column AA is blank
    new_df['AB'] = ''              # Column AB is blank
    new_df['AC'] = ''              # Column AC is blank
    new_df['AD'] = ''              # Column AD is blank
    new_df['AE'] = ''              # Column AE is blank
    new_df['AF'] = ''              # Column AF is blank
    new_df['AG'] = ''              # Column AG is blank
    new_df['AH'] = ''              # Column AH is blank
    new_df['AI'] = df.iloc[:, 22]  # Column W to Column AI
    new_df['AJ'] = ''              # Column AJ is blank
    new_df['AK'] = 'N'             # Column AK is "N" in every row

    # Reorder columns according to the required order
    columns_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK']
    new_df = new_df[columns_order]
    
    # Export to CSV without header
    new_df.to_csv(output_file, index=False, header=False)

    # Verify the number of rows
    print(f"Number of rows in the input file (excluding header): {len(df)}")
    print(f"Number of rows in the output file: {len(new_df)}")



# Example usage
input_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\ERMCO-TRANSFORMER-DATA071124.xls"  # Path to the input Excel file
output_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\NewErmcoData.csv"  # Path to the output CSV file
cost_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\cost.txt"
date_file_path = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Microsoft Forms\Meter Data Helper\Question\date.txt"



import pandas as pd
import re
from pathlib import Path

def transform_excel(input_file, output_file, cost_file, date_file_path):
    # Convert .xls to .xlsx if necessary
    input_file_path = Path(input_file)
    if input_file_path.suffix == '.xls':
        # Convert .xls to .xlsx
        df = pd.read_excel(input_file, engine='xlrd')
        converted_file_path = input_file_path.with_suffix('.xlsx')
        df.to_excel(converted_file_path, index=False)
        input_file = converted_file_path

    # Read the converted .xlsx file using openpyxl engine
    df = pd.read_excel(input_file, engine='openpyxl')
    # Log the shape of the DataFrame after loading
    print(f"Shape of the DataFrame after loading: {df.shape}")

    # Remove the header
    df = df.iloc[1:]

    # Log the shape of the DataFrame after removing the header
    print(f"Shape of the DataFrame after removing the header: {df.shape}")

    # Read the date from the text file
    with open(date_file_path, 'r') as file:
        date_text = file.readline().strip()

    # Read the cost from the text file
    with open(cost_file, 'r') as file:
        cost_text = file.readline().strip()

    # Function to extract only the numeric parts from a string
    def extract_numbers(s):
        return ''.join(filter(str.isdigit, str(s)))

    # Function to split the value in column K
    def split_k_value(value):
        parts = re.split(r'GRDY|/', str(value))
        if len(parts) < 2:
            return None, None
        num1 = parts[0]
        num2 = parts[1]
        return num1, num2

    # Create new DataFrame with the required transformations
    new_df = pd.DataFrame()
    new_df['A'] = df.iloc[:, 6]    # Column G to Column A
    new_df['B'] = ''               # Column B is blank
    new_df['C'] = 8                # Column C is 8 in every row
    new_df['D'] = df.iloc[:, 9]    # Column J to Column D
    new_df['E'] = df.iloc[:, 3]    # Column D to Column E
    
    # Split column K and assign to F and G
    df['K_num1'], df['K_num2'] = zip(*df.iloc[:, 10].apply(split_k_value))
    new_df['F'] = df['K_num2']  # Number after "/" to Column F
    new_df['G'] = df['K_num1']  # Number before "GRDY" to Column G
    
    new_df['H'] = df.iloc[:, 11].str.split('/').str[1]  # Low value of Column L to Column H
    new_df['I'] = df.iloc[:, 11].str.split('/').str[0]  # High value of Column L to Column I
    new_df['J'] = str(date_text)   # Column J is filled with a value read from a text file, formatted as string
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
    new_df['V'] = df.iloc[:, 12]   # Column M to Column V
    new_df['W'] = df.iloc[:, 8]    # Column I to Column W
    new_df['X'] = ''               # Column X is blank
    new_df['Y'] = ''               # Column Y is blank
    new_df['Z'] = ''               # Column Z is blank
    new_df['AA'] = ''              # Column AA is blank
    new_df['AB'] = ''              # Column AB is blank
    new_df['AC'] = ''              # Column AC is blank
    new_df['AD'] = ''              # Column AD is blank
    new_df['AE'] = ''              # Column AE is blank
    new_df['AF'] = ''              # Column AF is blank
    new_df['AG'] = ''              # Column AG is blank
    new_df['AH'] = ''              # Column AH is blank
    new_df['AI'] = df.iloc[:, 22]  # Column W to Column AI
    new_df['AJ'] = ''              # Column AJ is blank
    new_df['AK'] = 'N'             # Column AK is "N" in every row

    # Reorder columns according to the required order
    columns_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK']
    new_df = new_df[columns_order]
    
    # Log the shape of the new DataFrame
    print(f"Shape of the new DataFrame: {new_df.shape}")

    # Export to CSV without header
    new_df.to_csv(output_file, index=False, header=False)

    # Verify the number of rows
    print(f"Number of rows in the input file (excluding header): {len(df)}")
    print(f"Number of rows in the output file: {len(new_df)}")


# Example usage
input_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\ERMCO-TRANSFORMER-DATA071124.xls"  # Path to the input Excel file
output_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\NewErmcoData.csv"  # Path to the output CSV file
cost_file = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Data Transformations via Automation\cost.txt"
date_file_path = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Microsoft Forms\Meter Data Helper\Question\date.txt"

transform_excel(input_file, output_file, cost_file, date_file_path)






import pandas as pd

def transform_excel_to_fixed_width(input_excel_path, output_txt_path, date_file_path):
    # Read the Excel file
    df = pd.read_excel(input_excel_path)
    
    # Remove the header
    df = df.iloc[1:]
    
    # Extract the data from columns B and D
    column_b_data = df.iloc[:, 1].astype(str).str[:10].str.pad(10, fillchar=' ')
    column_d_data = df.iloc[:, 3].astype(str).str[:7].str.pad(7, fillchar=' ')
    
    # Read the date from the text file
    with open(date_file_path, 'r') as file:
        date_text = file.readline().strip()
    
    # Create a new DataFrame with the specified columns
    new_df = pd.DataFrame({
        'A': ['M'] * len(df),
        'B': column_b_data,
        'C': ['00:00'] * len(df),
        'D': [date_text] * len(df),
        'E': [' ' * 10] * len(df),
        'F': column_d_data
    })
    
    # Concatenate the columns to create the fixed width format
    fixed_width_lines = new_df.apply(lambda row: ''.join(row), axis=1)
    
    # Write to the output text file
    with open(output_txt_path, 'w') as file:
        for line in fixed_width_lines:
            file.write(line + '\n')



# Usage
input_excel_path = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Microsoft Forms\Meter Data Helper\Question\Cy 2 Meter Report.csv"
output_txt_path = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Microsoft Forms\Meter Data Helper\Question\newmeterdata.txt"
date_file_path = r"C:\Users\mikee\OneDrive - East Central Energy\Apps\Microsoft Forms\Meter Data Helper\Question\date.txt"

transform_excel_to_fixed_width(input_excel_path, output_txt_path, date_file_path)