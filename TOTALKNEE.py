import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

data_frame = pd.read_excel('Total Knees Only NR - CY 2022.xlsx', sheet_name='Master')
data_frame['DATE'] = pd.to_datetime(data_frame['DATE']).dt.date
data_frame['COST_EXT'] = data_frame['COST_EXT'].replace(r'[\$,]', '', regex=True).astype(float)

aggregation_dictionary = {
    'COST_EXT': 'sum',
    'PRI_SURG_NAME': 'first',
    'DATE': lambda x: 'TOTAL'
}

total_cost_ext = data_frame.groupby('LOG_ID').agg(aggregation_dictionary).reset_index()

data_frame = pd.concat([data_frame, total_cost_ext], ignore_index=True)
data_frame = data_frame.sort_values(by=['LOG_ID', 'DATE']).reset_index(drop=True)
data_frame['COST_EXT'] = data_frame['COST_EXT'].apply(lambda x: f"${x:,.2f}")

wb = Workbook()
ws = wb.active

for row_index, row in enumerate(dataframe_to_rows(data_frame, index=False, header=True), 1):
    for column_index, value in enumerate(row, 1):
        cell = ws.cell(row=row_index, column=column_index, value=value)
        if isinstance(value, str):
            cell.alignment = Alignment(wrapText=True)

wb.save('output.xlsx')

print("Output saved to 'output.xlsx'")
