import jellyfish
from dbsettings import database_parametres
from sqlalchemy import create_engine
import pandas as pd
from find_rearranged_ids import filter_arrays

def extract_initials(name):
    words = name.split()
    initials = ".".join(word[0] for word in words) + "."

    return initials
def update_elibrary_id(merged_data_link):
    try:
        merged_data = pd.read_excel(merged_data_link, index_col=False)
        matched_records = []
        matched_ids = []

        # Iterate through each row in filtered_author_info
        for index, row in merged_data.iterrows():
            author_id_filter = int(row['author_id'])
            author_name_filter = row['author_name']
            author_name_initials = row['author_initials']
            author_counter = row['counter']
            for _, db_row in merged_data.iterrows():
                author_counter_db = db_row['counter']
                author_id_db = db_row['author_id']
                author_name_db = db_row['author_name']
                author_initials_db = db_row['author_initials']
                similarity_ratio = jellyfish.jaro_winkler_similarity(author_name_db,  author_name_filter)
                if similarity_ratio >= 0.85 and author_id_filter != author_id_db:
                    if '.' in author_name_initials and  '.' in author_initials_db and author_name_initials == author_initials_db:
                        if [author_id_db, author_id_filter] not in matched_ids:
                            matched_records.append([author_counter_db, author_id_db, author_name_db, author_initials_db,
                                           author_counter, author_id_filter, author_name_filter, author_name_initials])
                            matched_ids.append([author_id_db, author_id_filter])
                    elif '.' in author_name_initials and  '.' not in author_initials_db:
                        author_initials_db = extract_initials(author_initials_db)
                        if author_initials_db == author_name_initials:
                            if [author_id_db, author_id_filter] not in matched_ids:
                                matched_records.append([author_counter_db, author_id_db, author_name_db, author_initials_db,
                                           author_counter, author_id_filter, author_name_filter, author_name_initials])
                                matched_ids.append([author_id_db, author_id_filter])
                    elif '.' not in author_name_initials and  '.'  in author_initials_db:
                        author_name_initials = extract_initials(author_name_initials)
                        if author_initials_db == author_name_initials:
                            if [author_id_db, author_id_filter] not in matched_ids:
                                matched_records.append([author_counter_db, author_id_db, author_name_db, author_initials_db,
                                           author_counter, author_id_filter, author_name_filter, author_name_initials])
                                matched_ids.append([author_id_db, author_id_filter])
                    elif '.' not in author_name_initials and  '.'  not  in author_initials_db:
                        initials_similarity_ratio = jellyfish.jaro_winkler_similarity(author_name_initials, author_initials_db)
                        if initials_similarity_ratio >= 0.85:
                            if [author_id_db, author_id_filter] not in matched_ids:
                                matched_records.append([author_counter_db, author_id_db, author_name_db, author_initials_db,
                                           author_counter, author_id_filter, author_name_filter, author_name_initials])
                                matched_ids.append([author_id_db, author_id_filter])
        matched_records = filter_arrays(matched_records)
        print(matched_records)
        return matched_records


    except Exception as e:
        print(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    update_elibrary_id('merged_ao.xlsx')