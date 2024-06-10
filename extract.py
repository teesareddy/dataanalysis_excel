import openpyxl
import pandas as pd
import re

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

# Load the Excel file for filtering
df = pd.read_excel('file8.xlsx')

# Use the extracted strings as keywords
keywords = strings

# Specify the column in which to search for keywords
search_column = 'Goods Description'

# Create a new column to store matched keywords
df['Matched Keywords'] = ''

# Filter rows based on keywords
for keyword in keywords:
    pattern = fr'\b({keyword})\b'  # Use a regex pattern to match whole words with a capture group 
    matches = df[search_column].str.extract(pattern, flags=re.IGNORECASE)
    df.loc[matches.notnull().any(axis=1), 'Matched Keywords'] += matches.iloc[:, 0] + ', '

# Remove the trailing comma and space from the 'Matched Keywords' column
df['Matched Keywords'] = df['Matched Keywords'].str.rstrip(', ')

# Create a new column 'Matched Keyword' with the first matched keyword
df['Matched Keyword'] = df['Matched Keywords'].str.split(',').str[0]

# Filter the data and drop the 'Matched Keywords' column
filtered_data = df[df['Matched Keywords'] != ''].drop('Matched Keywords', axis=1)

# Write filtered data to a new Excel sheet
filtered_data.to_excel('output_file.xlsx', index=False)

print("Filtered data has been saved to 'output_file.xlsx'")