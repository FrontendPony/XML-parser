import pandas as pd

def update_author_id_on_choice(file_path, another_file_path, output_file):
    df = pd.read_excel(file_path)
    another_df = pd.read_excel(another_file_path)

    new_df = df[(df['author_id'] != df['author_id_choice']) & (df['author_id_choice'].notna())][
        ['author_id', 'author_id_choice']]
    new_df = new_df.rename(columns={'author_id': 'additional_author_id', 'author_id_choice': 'author_id'})
    new_df.to_excel('alternative_ids.xlsx')

    for index, row in df.iterrows():
        author_id = row['author_id']
        author_id_choice = row['author_id_choice']

        if not pd.isnull(author_id_choice):
            mask = another_df['author_id'] == int(author_id)
            another_df.loc[mask, 'author_id'] = int(author_id_choice)

    another_df.to_excel(output_file, index=False)


