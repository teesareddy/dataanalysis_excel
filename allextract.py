import openpyxl
import pandas as pd
import re
import os

# Load the Excel file for extracting strings
wb = openpyxl.load_workbook('ref.xlsx')
sheet = wb.active

# Get the header row
header_row = [cell.value for cell in sheet[1]]

# Extract all strings from the sheet, excluding the header row
strings = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    for cell in row:
        if isinstance(cell, str):
            strings.append(cell)

# Use the extracted strings as keywords
keywords = [keyword for keyword in strings if keyword.strip()]

# Specify the file names to filter
file_names = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx', 'file4.xlsx', 'file5.xlsx', 'file6.xlsx', 'file7.xlsx', 'file8.xlsx', 'file9.xlsx', 'file10.xlsx', 'file11.xlsx', 'file12.xlsx']

for file_name in file_names:
    if os.path.isfile(file_name):
        # Load the Excel file for filtering
        df = pd.read_excel(file_name)

        if 'Goods Description' in df.columns:
            # Create a new column to store matched keywords
            df['Matched Keywords'] = ''

            # Filter rows based on keywords
            for keyword in keywords:
                pattern = fr'\b({keyword})\b'
                matches = df['Goods Description'].str.extract(pattern, flags=re.IGNORECASE)
                df.loc[matches.notnull().any(axis=1), 'Matched Keywords'] += matches.iloc[:, 0] + ', '

            # Remove the trailing comma and space from the 'Matched Keywords' column
            df['Matched Keywords'] = df['Matched Keywords'].str.rstrip(', ')

            # Create a new column 'Matched Keyword' with the first matched keyword
            df['Matched Keyword'] = df['Matched Keywords'].str.split(',').str[0]

            # Filter the data and drop the 'Matched Keywords' column
            filtered_data = df[df['Matched Keywords'] != ''].drop('Matched Keywords', axis=1)

            # Write filtered data to a new Excel sheet
            output_file = f'output_{os.path.splitext(file_name)[0]}.xlsx'
            filtered_data.to_excel(output_file, index=False)

            print(f"Filtered data for {file_name} has been saved to '{output_file}'")
        else:
            print(f"Column 'Goods Description' not found in '{file_name}'. Skipping...")
    else:
        print(f"File '{file_name}' not found. Skipping...")