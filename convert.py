import pandas as pd

xlsx_file = 'TBM Organization Parameters.xlsx'

# Load the Excel file
excel = pd.ExcelFile(xlsx_file)

# Loop through each sheet and convert to .csv
for sheet_name in excel.sheet_names:
    df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
    # Save each sheet as a separate .csv file
    csv_file = f'diff/{sheet_name}.csv'
    df.to_csv(csv_file, index=False)

print("Conversion completed successfully.")