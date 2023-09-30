import pandas as pd
from sqlalchemy import create_engine
from dbsettings import database_parametres
import random
import os
import time
import psutil

def generate_unique_id(author_id_list):
    while True:
        new_id = random.randint(1000000000, 9999999999)
        if new_id not in author_id_list:
            author_id_list.append(new_id)
            return new_id

def update_author_id(excel_file_path):
    try:
        database_url = f"postgresql://{database_parametres['user']}:{database_parametres['password']}@{database_parametres['host']}:{database_parametres['port']}/{database_parametres['dbname']}"
        engine = create_engine(database_url)
        query = """
                    SELECT DISTINCT author_id, author_name, author_initials
                    FROM authors_organisations
                    """
        df_database = pd.read_sql_query(query, engine)
        author_id_list = df_database['author_id'].tolist()
        df = pd.read_excel(excel_file_path, index_col=0)
        df_null = pd.read_excel('author_filtered_data.xlsx', index_col=0)
        df_null['generated_ids'] = df_null.apply(lambda row: generate_unique_id(author_id_list), axis=1)
        df['author_id'] = df['author_id'].astype(str)
        unique_author_names = (df_null['author_fullname'])
        filtered_rows = df[(df['author_fullname'].isin(unique_author_names)) & (df['author_id'] != ' ')]
        df_database['author_fullname'] = df_database['author_name'] + " " + df_database['author_initials']
        filtered_rows_from_db = df_database[(df_database['author_fullname'].isin(unique_author_names))]

        if len(filtered_rows) > 0:
            filtered_rows = filtered_rows.drop_duplicates(subset=['author_fullname'], keep='last')
            author_name_id_dict = dict(zip(filtered_rows['author_fullname'], filtered_rows['author_id']))

            def compute_possible_id(row):
                if row['author_fullname'] in author_name_id_dict:
                    return author_name_id_dict[row['author_fullname']]
                else:
                    return ' '

            df_null['possible_id_from_xml'] = df_null.apply(compute_possible_id, axis=1)
        if len(filtered_rows_from_db) > 0:
                filtered_rows_from_db = filtered_rows_from_db.drop_duplicates(subset=['author_fullname'], keep='last')


                author_name_id_dict = dict(zip(filtered_rows_from_db['author_fullname'], filtered_rows_from_db['author_id']))

                def compute_possible_id(row):
                    if row['author_fullname'] in author_name_id_dict:
                        return author_name_id_dict[row['author_fullname']]
                    else:
                        return ' '

                df_null['possible_id_from_db'] = df_null.apply(compute_possible_id, axis=1)
    except Exception as e:
        print("An error occurred:", str(e))
    df_null['enter_id'] = ''
    df_null.to_excel('author_filtered_data.xlsx')
    def check_author_id(file_path):
        df = pd.read_excel(file_path)
        return all(df['enter_id'] != ' ')
    while True:
        os.system(f'start excel author_filtered_data.xlsx')
        while True:
            time.sleep(1)
            excel_running = False
            for process in psutil.process_iter(attrs=['pid', 'name']):
                if "EXCEL.EXE" in process.info['name']:
                    excel_running = True
                    break
            if not excel_running:
                break
        if check_author_id('author_filtered_data.xlsx'):
            break
        else:
            print("Some rows have empty 'org_id'. Rerunning the code.")

    print("Excel file has been closed. Now, running additional code.")

    def update_author_id(row):
        if row['author_id'] == ' ' and row['author_fullname'] in df_null2['author_fullname'].values:
            matching_row = df_null2[df_null2['author_fullname'] == row['author_fullname']]
            return matching_row['enter_id'].values[0]
        else:
            return row['author_id']

    df_null2 = pd.read_excel('author_filtered_data.xlsx', index_col=0)
    df['author_id'] = df.apply(update_author_id, axis=1)
    df.to_excel('authors_organisations.xlsx')

if __name__ == "__main__":
    update_author_id('authors_organisations.xlsx')

