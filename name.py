import pandas as pd

# List to store DataFrames of all Excel files
dfs = []

# Load and store the first row of all 12 Excel files
for i in range(1, 13):
    df = pd.read_excel(f'file{i}.xlsx', header=None, nrows=1)
    dfs.append(df)

# Dictionary to store the count of similar comparisons for each file
similar_counts = {i: 0 for i in range(1, 13)}

# Compare the first row of each file with the first rows of all other files
for i, df1 in enumerate(dfs, start=1):
    for j, df2 in enumerate(dfs, start=1):
        if i != j:  # Skip self-comparison
            if df1.equals(df2):
               similar_counts[i] += 1

# Find the file names with the most similar comparisons
max_count = max(similar_counts.values())
similar_files = [i for i, count in similar_counts.items() if count == max_count]

# Output the file names with the most similar comparisons
if similar_files:
    print("Files with the most similar comparisons:")
    for file_num in similar_files:
        print(f"file{file_num}.xlsx")
else:
    print("No files have similar comparisons.")