import pandas as pd
from sqlalchemy import create_engine
from dbsettings import database_parametres
import random
import os
import time
import psutil


def generate_unique_id(org_id_list):
    while True:
        new_id = random.randint(1000000000, 9999999999)
        if new_id not in org_id_list:
            org_id_list.append(new_id)
            return new_id

def update_org_id(excel_file_path):
    try:
        database_url = f"postgresql://{database_parametres['user']}:{database_parametres['password']}@{database_parametres['host']}:{database_parametres['port']}/{database_parametres['dbname']}"
        engine = create_engine(database_url)
        query = """
                    SELECT DISTINCT org_id, org_name
                    FROM authors_organisations
                    """
        df_database = pd.read_sql_query(query, engine)
        org_id_list = df_database['org_id'].tolist()
        df = pd.read_excel(excel_file_path, index_col=0)
        df_null = pd.read_excel('org_filtered_data.xlsx', index_col=0)
        df_null['generated_ids'] = df_null.apply(lambda row: generate_unique_id(org_id_list), axis=1)
        df['org_id'] = df['org_id'].astype(str)

        unique_org_names = df_null['org_name']

        filtered_rows = df[(df['org_name'].isin(unique_org_names)) & (df['org_id'] != ' ')]
        filtered_rows_from_db = df_database[(df_database['org_name'].isin(unique_org_names))]

        if len(filtered_rows) > 0:
            filtered_rows = filtered_rows.drop_duplicates(subset=['org_name'], keep='last')
            org_name_id_dict = dict(zip(filtered_rows['org_name'], filtered_rows['org_id']))

            def compute_possible_id(row):
                if row['org_name'] in org_name_id_dict:
                    return org_name_id_dict[row['org_name']]
                else:
                    return ' '
            df_null['possible_id_from_xml'] = df_null.apply(compute_possible_id, axis=1)
        if len(filtered_rows_from_db) > 0:

                filtered_rows_from_db = filtered_rows_from_db.drop_duplicates(subset=['org_name'], keep='last')
                org_name_id_dict = dict(zip(filtered_rows_from_db['org_name'], filtered_rows_from_db['org_id']))

                def compute_possible_id(row):
                    if row['org_name'] in org_name_id_dict:
                        return org_name_id_dict[row['org_name']]
                    else:
                        return ' '

                df_null['possible_id_from_db'] = df_null.apply(compute_possible_id, axis=1)
    except Exception as e:
        print("An error occurred:", str(e))
    df_null['enter_id'] = ''
    df_null.to_excel('org_filtered_data.xlsx')
    def check_author_id(file_path):
        df = pd.read_excel(file_path)
        return all(df['enter_id'] != ' ')
    while True:
        os.system(f'start excel org_filtered_data.xlsx')
        while True:
            time.sleep(1)
            excel_running = False
            for process in psutil.process_iter(attrs=['pid', 'name']):
                if "EXCEL.EXE" in process.info['name']:
                    excel_running = True
                    break
            if not excel_running:
                break
        if check_author_id('org_filtered_data.xlsx'):
            break
        else:
            print("Some rows have empty 'org_id'. Rerunning the code.")

    print("Excel file has been closed. Now, running additional code.")
    df_null = pd.read_excel('org_filtered_data.xlsx', index_col=0)
    def update_org_id(row):
        if row['org_id'] == ' ' and row['org_name'] in df_null['org_name'].values:
            matching_row = df_null[df_null['org_name'] == row['org_name']]
            return matching_row['enter_id'].values[0]
        else:
            return row['org_id']

    df['org_id'] = df.apply(update_org_id, axis=1)
    df.to_excel('authors_organisations.xlsx')

if __name__ == "__main__":
    update_org_id('authors_organisations.xlsx')