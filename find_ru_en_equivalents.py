import pandas as pd
import re
from transliterate import translit
import jellyfish
from apply_colours_to_excel import apply_fill_colors

def process_excel_file(input_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Define a regular expression pattern for English names
    english_name_pattern = re.compile(r'^[A-Za-z]')

    # Use the pattern to filter rows with English full names
    english_names_df = df[df['author_fullname'].apply(lambda x: bool(english_name_pattern.match(x)))]


    # Transliterate English names to Russian
    english_names_df['russian_equivalent'] = english_names_df['author_fullname'].apply(lambda x: translit(x, 'ru'))
    # Add a new column 'has_en_pair' and mark it as True or False based on Jaro-Winkler similarity
    df['has_en_pair'] = df.apply(
        lambda row: any(jellyfish.jaro_winkler_similarity(row['author_fullname'], x) > 0.9 for x in english_names_df['russian_equivalent']),
        axis=1
    )

    # Create empty lists to store matched English names and corresponding row author names
    matched_english_names = []
    matched_row_author_names = []

    # Iterate through each row and check for matching criteria
    for index, row in df.iterrows():
        # Check Jaro-Winkler similarity for each English name
        matches = [x for x in english_names_df['russian_equivalent'] if
                   jellyfish.jaro_winkler_similarity(row['author_fullname'], x) > 0.9]

        # If there's a match, store the English name and row author name
        if matches:
            for match in matches:
                matched_english_names.append(
                    english_names_df.loc[english_names_df['russian_equivalent'] == match, 'author_fullname'].iloc[0])
                matched_row_author_names.append(row['author_fullname'])

    combined_array = [[x, y] for x, y in zip(matched_english_names, matched_row_author_names)]
    print(combined_array)
    replace_dict = {pair[0]: pair[1] for pair in combined_array}
    df = pd.read_excel('author_filtered_data.xlsx')
    df['formatted_author_name'] = df['formatted_author_name'].apply(lambda x: replace_dict.get(x, x))
    df.to_excel('author_filtered_data.xlsx')

