import pandas as pd

# Read the Excel file into a DataFrame
excel_file = 'authors_organisations.xlsx'
df = pd.read_excel(excel_file)

# Group rows by 'author_id' and 'org_id'
grouped = df.groupby(['author_id', 'org_id'])

# Create a list to store the grouped data
grouped_data = []

# Iterate through the groups and store the rows with the same author_id and org_id
for key, group in grouped:
    if len(group) > 1:
        grouped_data.append(group)

# Create a new DataFrame with the grouped data
result_df = pd.concat(grouped_data)

# Save the result to a new Excel file
result_excel_file = 'result_excel_file2.xlsx'
result_df.to_excel(result_excel_file, index=False)
