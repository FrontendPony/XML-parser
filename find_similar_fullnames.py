import pandas as pd
import re
from apply_colours_to_excel import apply_fill_colors
from find_ru_en_equivalents import process_excel_file

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
            df.to_excel('author_filtered_data.xlsx')
            process_excel_file('author_filtered_data.xlsx')
            df = pd.read_excel('author_filtered_data.xlsx')
            grouped = df.groupby('formatted_author_name')
            group_dfs = []

            for name, group in grouped:
                group_dfs.append(group)

            result_df = pd.concat(group_dfs)
            result_df = result_df.loc[:, ~result_df.columns.str.contains('^Unnamed')]
            result_df.to_excel(file_path, index=False)
            apply_fill_colors(file_path)
    except Exception as e:
        print(f"Error processing the Excel file: {str(e)}")

if __name__ == "__main__":
    find_similar_fullnames('author_filtered_data.xlsx')