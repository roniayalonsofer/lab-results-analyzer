import sys; sys.path.insert(0, '.')
import pandas as pd, re

vsl_path = r'thresholds/soil_vsl_v7_full.xlsx'
xl = pd.ExcelFile(vsl_path)
print('Sheets:', xl.sheet_names)

# Search all sheets for C10-C40 / diesel / petroleum
for sname in xl.sheet_names:
    df = xl.parse(sname, header=None, dtype=str).fillna('')
    for i, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values).lower()
        if 'c10' in row_str or 'diesel' in row_str or 'c8-c40' in row_str:
            vals = list(row.values)[:10]
            print(f'Sheet={repr(sname)} row={i}: {vals}')
