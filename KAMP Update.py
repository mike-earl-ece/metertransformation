


import pyodbc
import pyautogui
import time

from datetime import datetime, timedelta

# Get the current date
now = datetime.now()

# Calculate the date for the previous month
prev_month = now - timedelta(days=30)

# Format the date as YYYYMM
prev_month_text = prev_month.strftime('%Y%m')

# Set up the connection string
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=R:\IT\Data Analyst\Databases\KAMP\KAMP_UPDATE.mdb;'
    )

# Connect to the database
conn = pyodbc.connect(conn_str)

# Create a cursor
cursor = conn.cursor()

# Execute a macro
cursor.execute("SELECT * FROM M_RUN QUERIES")

# Wait for the text box to appear
time.sleep(2)

# Enter the text in the box using pyautogui
pyautogui.typewrite(prev_month_text)

# Press enter to submit the text
pyautogui.press('enter')

# Commit the changes
conn.commit()

# Close the connection
conn.close()



