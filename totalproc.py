#!/usr/bin/env python3

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import sys
import os

def process_procedures(input_file):

    base_name = os.path.splitext(os.path.basename(input_file))[0]

    data_frame = pd.read_excel(input_file, sheet_name='Master')
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

    output_file = f"{base_name}_processed.xlsx"
    wb.save(output_file)

    print(f"Output saved to '{output_file}'")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: totalproc <input_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    process_procedures(input_file)

