import pandas as pd
from datetime import datetime
import os
import json
import azure.functions as func

def validate_paths(input_path, output_txt_path):
    output_dir = os.path.dirname(output_txt_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Info: The output directory '{output_dir}' was created.")
    return True

def read_file(input_path):
    try:
        if input_path.endswith('.xlsx'):
            return pd.read_excel(input_path, sheet_name='Sheet1')
        elif input_path.endswith('.csv'):
            return pd.read_csv(input_path)
        else:
            print("Error: Unsupported file format.")
            return None
    except Exception as e:
        print(f"Error reading the input file: {e}")
        return None

def process_columns(df):
    try:
        column_b_data = df.iloc[:, 0].astype(str).str[:10].str.pad(10, fillchar=' ')
        column_c_data = df.iloc[:, 1].astype(str).str[:7].str.pad(7, fillchar=' ')

        column_e_data = pd.to_datetime(df.iloc[:, 2], errors='coerce').dt.strftime('%y/%m/%d').fillna(' ' * 8)

        return column_b_data, column_c_data, column_e_data
    except Exception as e:
        print(f"Error processing columns: {e}")
        return None, None, None

def write_fixed_width_file(output_txt_path, new_df):
    try:
        formatted_lines = new_df.apply(lambda row: ','.join(row.astype(str)), axis=1)
        with open(output_txt_path, 'w') as file:
            for line in formatted_lines:
                file.write(line + '\n')

        print(f"File written successfully to {output_txt_path}")
    except Exception as e:
        print(f"Error writing to the output file: {e}")

def transform_to_fixed_width(input_path, output_txt_path):
    df = read_file(input_path)
    if df is None:
        return

    column_b_data, column_c_data, column_e_data = process_columns(df)
    if column_b_data is None or column_c_data is None or column_e_data is None:
        return

    new_df = pd.DataFrame({
        'A': ['M'] * len(df),
        'B': column_b_data,
        'C': column_c_data,
        'D': ['00:00'] * len(df),
        'E': column_e_data,
        'F': [' ' * 10] * len(df)
    })

    write_fixed_width_file(output_txt_path, new_df)

def main(req: func.HttpRequest) -> func.HttpResponse:
    try:
        req_body = req.get_json()
        input_path = req_body.get('input_path')
        output_txt_path = req_body.get('output_txt_path')

        transform_to_fixed_width(input_path, output_txt_path)
        
        return func.HttpResponse(
            json.dumps({'status': 'Success', 'message': 'File processed successfully'}),
            mimetype='application/json'
        )
    except Exception as e:
        return func.HttpResponse(
            json.dumps({'status': 'Error', 'message': str(e)}),
            mimetype='application/json',
            status_code=500
        )
