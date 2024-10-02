import logging
import pandas as pd
from datetime import datetime
import azure.functions as func
from io import StringIO

def main(req: func.HttpRequest) -> func.HttpResponse:
    try:
        file = req.files['file']
        if file:
            df = pd.read_csv(file)
            column_a_data = df.iloc[:, 0].astype(str).str[:10].str.pad(10, fillchar=' ')
            column_b_data = df.iloc[:, 1].astype(str).str[:7].str.pad(7, fillchar=' ')
            column_c_data = pd.to_datetime(df.iloc[:, 2], format='%m/%d/%Y', errors='coerce').dt.strftime('%y/%m/%d').fillna(' ')

            new_df = pd.DataFrame({
                'A': ['M'] * len(df),
                'B': column_a_data,
                'C': column_b_data,
                'D': ['00:00'] * len(df),
                'E': column_c_data,
                'F': [' ' * 10] * len(df)
            })

            output = StringIO()
            new_df.to_csv(output, index=False, header=False)
            output.seek(0)
            return func.HttpResponse(output.getvalue(), mimetype="text/plain")
        else:
            return func.HttpResponse("No file uploaded", status_code=400)
    except Exception as e:
        logging.error(f"Error processing file: {e}")
        return func.HttpResponse(f"Error processing file: {e}", status_code=500)
