import pandas as pd
import re
from apply_colours_to_excel import apply_fill_colors

def format_name(name):
    match = re.match(r"(([А-Яа-я]+) ([А-Яа-я]+) ([А-Яа-я]+))", name)
    if match:
        first_letter = match.group(3)[0]
        second_letter = match.group(4)[0]
        return f"{match.group(2)} {first_letter}.{second_letter}."
    else:
        return name

def find_similar_fullnames(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'author_fullname' in df.columns:
            author_fullnames = df['author_fullname'].tolist()
            formatted_data = [format_name(name) for name in author_fullnames]
            df['formatted_author_name'] = formatted_data
            grouped = df.groupby('formatted_author_name')
            # Create a list to store DataFrames for each group
            group_dfs = []

            # Iterate through the groups
            for name, group in grouped:
                group_dfs.append(group)

            # Concatenate all group DataFrames into one DataFrame
            result_df = pd.concat(group_dfs)

            # Save the combined DataFrame to a single Excel file
            result_df.to_excel(file_path, index=False)
            apply_fill_colors(file_path)
    except Exception as e:
        print(f"Error processing the Excel file: {str(e)}")

if __name__ == "__main__":
    find_similar_fullnames('author_filtered_data.xlsx')