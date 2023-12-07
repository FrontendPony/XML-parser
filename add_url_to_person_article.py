import pandas as pd

def add_url_to_person_article(data_arrays_file, merged_link_file, merged_file):
    file1 = pd.read_excel(data_arrays_file)
    file2 = pd.read_excel(merged_link_file)
    file3 = pd.read_excel(merged_file)

    file1['original_order'] = range(len(file1))

    merged_files = pd.merge(file1, file2[['counter', 'item_id']], on='counter')
    merged_files = merged_files.drop_duplicates(subset='counter')
    merged_files = pd.merge(merged_files, file3[['linkurl', 'item_id']], on='item_id')
    merged_files = merged_files.drop_duplicates(subset='counter')
    if 'item_id' in merged_files.columns:
        merged_files = merged_files.drop('item_id', axis=1)

    merged_files = merged_files.sort_values(by='original_order')

    if 'original_order' in merged_files.columns:
        merged_files = merged_files.drop('original_order', axis=1)

    merged_files.to_excel('possible_duplicate_people.xlsx', index=False)


