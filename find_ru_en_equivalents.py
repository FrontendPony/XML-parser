import pandas as pd
import re
from transliterate import translit
import jellyfish

def process_excel_file(input_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    english_name_pattern = re.compile(r'^[A-Za-z]')

    english_names_df = df[df['author_fullname'].apply(lambda x: bool(english_name_pattern.match(x)))]

    english_names_df['russian_equivalent'] = english_names_df['author_fullname'].apply(lambda x: translit(x, 'ru'))

    df['has_en_pair'] = df.apply(
        lambda row: any(jellyfish.jaro_winkler_similarity(row['author_fullname'], x) > 0.9 for x in english_names_df['russian_equivalent']),
        axis=1
    )

    df.to_excel(input_file)

# Example usage:
# process_excel_file('author_filtered_data.xlsx')
