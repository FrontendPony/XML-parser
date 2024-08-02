import psycopg2
import pandas as pd
from dbsettings import database_parametres as database_parameters
from sqlalchemy import create_engine
def connect_and_fetch_data(database_params, query):
    try:
        connection_str = f"postgresql://{database_params['user']}:{database_params['password']}@{database_params['host']}:{database_params['port']}/{database_params['dbname']}"
        engine = create_engine(connection_str)
        df = pd.read_sql(query, engine)
        return df
    except :
        return pd.DataFrame()

def add_url_to_person_article(data_arrays_file, merged_link_file, merged_file,  journal, conferences):
    file1 = pd.read_excel(data_arrays_file)
    file2 = pd.read_excel(merged_link_file)
    file3 = pd.read_excel(merged_file)

    file1['original_order'] = range(len(file1))

    merged_files = pd.merge(file1, file2[['counter', 'item_id']], on='counter', how='left')
    merged_files = merged_files.drop_duplicates(subset='counter')
    merged_files = pd.merge(merged_files, file3[['linkurl', 'item_id']], on='item_id', how='left')
    merged_files = merged_files.drop_duplicates(subset='counter')
    if journal:
        data = connect_and_fetch_data(database_parameters, "SELECT item_id, linkurl FROM conference")
    elif conferences:
        data = connect_and_fetch_data(database_parameters, "SELECT item_id, linkurl FROM article")
    merged_files = pd.merge(merged_files, data[['linkurl', 'item_id']], on='item_id', how='left')
    merged_files = merged_files.drop_duplicates(subset='counter')
    if 'item_id' in merged_files.columns:
        merged_files = merged_files.drop('item_id', axis=1)

    merged_files = merged_files.sort_values(by='original_order')

    if 'original_order' in merged_files.columns:
        merged_files = merged_files.drop('original_order', axis=1)

    merged_files.to_excel('possible_duplicate_people.xlsx', index=False)

