import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# Data is stored in an Excel file named 'Total Knees Only NR - CY 2022.xlsx', and the data is on the "Master" tab
# Read it into a pandas DataFrame
df = pd.read_excel('Total Knees Only NR - CY 2022.xlsx', sheet_name='Master')

# Convert DATE column to datetime and remove time component
df['DATE'] = pd.to_datetime(df['DATE']).dt.date

# Convert COST_EXT to US Dollars
df['COST_EXT'] = df['COST_EXT'].apply(lambda x: f"${x:,.2f}")

# Explode the DESCRIPTION, CAT#, CONSOLIDATED MFG, and COST_EXT columns individually
df = df.explode('DESCRIPTION').explode('CAT#').explode('CONSOLIDATED MFG').explode('COST_EXT').reset_index(drop=True)

# Group the data by LOG_ID and aggregate the other columns
grouped = df.groupby('LOG_ID').agg({
    'DATE': 'first',
    'PRI_SURG_NAME': 'first',
    'PROCEDURE': 'first',
    'DESCRIPTION': lambda x: '\n'.join(map(str, x)),
    'CAT#': lambda x: '\n'.join(map(str, x)),
    'CONSOLIDATED MFG': lambda x: '\n'.join(map(str, x)),
    'COST_EXT': lambda x: '\n'.join(map(str, x))
}).reset_index()

# Calculate the total COST_EXT for each LOG_ID
total_cost_ext = df.groupby('LOG_ID')['COST_EXT'].apply(lambda x: x.str.replace('$', '').str.replace(',', '').astype(float).sum()).reset_index()
total_cost_ext['DATE'] = 'TOTAL'
total_cost_ext['PROCEDURE'] = ''

# Retrieve PRI_SURG_NAME for the TOTAL row
total_cost_ext['PRI_SURG_NAME'] = total_cost_ext.apply(lambda row: df[df['LOG_ID'] == row['LOG_ID']]['PRI_SURG_NAME'].iloc[0], axis=1)

# Append the total COST_EXT row to the grouped DataFrame
grouped = pd.concat([grouped, total_cost_ext], ignore_index=True)

# Create a new Excel workbook and worksheet
wb = Workbook()
ws = wb.active

# Convert the Pandas DataFrame to rows and write to the worksheet
for r_idx, row in enumerate(dataframe_to_rows(grouped, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=value)
        # Adjust row heights to fit the content
        if isinstance(value, str):
            cell.alignment = Alignment(wrapText=True)

# Save the workbook
wb.save('output.xlsx')

print("Output saved to 'output.xlsx'")
