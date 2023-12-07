import pandas as pd


def update_additional_author_id(author_id, additional_author_id, upper_row_counter, lower_row_counter):
    data = {'author_id': [author_id], 'additional_author_id': [additional_author_id]}
    additional_ids_df = pd.DataFrame(data)

    new_file_path = 'additional_ids.xlsx'
    additional_ids_df.to_excel(new_file_path, index=False)

    existing_file_path = 'merged_ao.xlsx'
    ao_df = pd.read_excel(existing_file_path)

    ao_df = ao_df[ao_df['counter'] != lower_row_counter]
    ao_df.to_excel(existing_file_path, index=False)

    another_file_path = 'merged_link.xlsx'
    another_df = pd.read_excel(another_file_path)

    mask = another_df['counter'] == lower_row_counter
    another_df.loc[mask, 'counter'] = upper_row_counter

    another_df.to_excel(another_file_path, index=False)

