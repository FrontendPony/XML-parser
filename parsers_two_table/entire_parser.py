from bs4 import BeautifulSoup
import pandas as pd

from sqlalchemy import create_engine
from dbsettings import database_parametres
import psycopg2
import re
from find_similar_fullnames import find_similar_fullnames
def extract_numbers_from_string(input_string):
    pattern = r'\d+'
    number = re.findall(pattern, input_string)
    number = [int(num) for num in number]
    return number
def parse_articles_to_excel(xml_filename):
    connection_str = f"postgresql://{database_parametres['user']}:{database_parametres['password']}@{database_parametres['host']}:{database_parametres['port']}/{database_parametres['dbname']}"
    engine = create_engine(connection_str)
    existing_data_query = f"SELECT * FROM authors_organisations"
    existing_data = pd.read_sql(existing_data_query, engine)
    counter_dict = {}
    unique_pairs = set()

    for _, row in existing_data.iterrows():
        key = (str(row['author_id']), str(row['org_id']))
        value = row['counter']
        counter_dict[key] = value

    for _, row in existing_data.iterrows():
        pair = (str(row['author_id']), str(row['org_id']))
        unique_pairs.add(pair)
    query = """
                             SELECT MAX(counter) FROM authors_organisations
                            """

    conn = psycopg2.connect(
        dbname=database_parametres['dbname'],
        user=database_parametres['user'],
        password=database_parametres['password'],
        host=database_parametres['host'],
        port=database_parametres['port']
    )
    cur = conn.cursor()
    cur.execute(query)
    fetcheData = cur.fetchone()
    number = extract_numbers_from_string(str(fetcheData))
    fields = {"item_id": [], 'linkurl': [], 'genre': [], 'type': [], "journal_title": [], "issn": [], "eissn": [],
              "publisher": [], "vak": [], "rcsi": [], "wos": [], "scopus": [], "quartile": [], "year": [], "number": [],
              'contnumber': [], "volume": [], "page_begin": [], "page_end": [], "language": [], "title_article": [],
              "doi": [], "edn": [], 'grnti': [], 'risc': [], 'corerisc': [], 'counter': []}

    fd = open(xml_filename, 'r', encoding='utf-8')
    xml_file = fd.read()
    soup = BeautifulSoup(xml_file, 'lxml')

    author_organisation = []
    array_of_dicts = []

    if len(number) == 0:
        counter = 0
        counter_all = 0
    else:
        counter = number[0]
        counter_all = number[0]

    unique_combinations = set()
    counter_dict_fornull_author = {}
    counter_dict_fornull_org = {}
    counter_dict_fornull_author_and_org = {}
    author_count = []
    for tag in soup.findAll("item"):
        # item
        fields['item_id'].append(tag['id'])
        fields['linkurl'].append(tag.find('linkurl').text if tag.find('linkurl') is not None else "")
        linkurl = tag.find('linkurl').text if tag.find('linkurl') is not None else ""
        fields['genre'].append(tag.find('genre').text if tag.find('genre') is not None else "")
        fields['type'].append(tag.find('type').text if tag.find('type') is not None else "")

        # journal
        fields['journal_title'].append(tag.find('journal').find('title').text if tag.find('journal').find('title') is not None else "")
        fields['issn'].append(tag.find('journal').find('issn').text if tag.find('journal').find('issn') is not None else "")
        fields['eissn'].append(tag.find('journal').find('eissn').text if tag.find('journal').find('eissn') is not None else "")
        fields['publisher'].append(tag.find('journal').find('publisher').text if tag.find('journal').find('publisher') is not None else "")
        fields['vak'].append(tag.find('journal').find('vak').text if tag.find('journal').find('vak') is not None else "")
        fields['rcsi'].append(tag.find('journal').find('rcsi').text if tag.find('journal').find('rcsi') is not None else "")
        fields['wos'].append(tag.find('journal').find('wos').text if tag.find('journal').find('wos') is not None else "")
        fields['scopus'].append(tag.find('journal').find('scopus').text if tag.find('journal').find('scopus') is not None else "")
        fields['quartile'].append("")

        # issue
        fields['year'].append(tag.find('issue').find('year').text if tag.find('issue').find('year') is not None else "")
        fields['number'].append(tag.find('issue').find('number').text if tag.find('issue').find('number') is not None else "")
        fields['contnumber'].append(tag.find('issue').find('contnumber').text if tag.find('issue').find('contnumber') is not None else "")
        fields['volume'].append(tag.find('issue').find('volume').text if tag.find('issue').find('volume') is not None else "")

        # item
        list_pages = tag.find('pages').text.split("-") if tag.find('pages') is not None else [" "]
        if len(list_pages) == 2:
            fields["page_begin"].append(list_pages[0])
            fields["page_end"].append(list_pages[1])
        else:
            fields["page_begin"].append(list_pages[0])
            fields["page_end"].append(list_pages[0])
        fields['language'].append(tag.find('language').text if tag.find('language') is not None else "")

        # titles
        fields['title_article'].append(tag.find('titles').find('title').text if tag.find('titles').find('title') is not None else "")
        # item
        fields['doi'].append(tag.find('doi').text if tag.find('doi') is not None else "")
        fields['edn'].append(tag.find('edn').text if tag.find('edn') is not None else "")
        fields['grnti'].append(tag.find('grnti').text if tag.find('grnti') is not None else "")
        fields['risc'].append(tag.find('risc').text if tag.find('risc') is not None else "")
        fields['corerisc'].append(tag.find('corerisc').text if tag.find('corerisc') is not None else "")

        count_author_org = []
        langArray = []
        author = soup.find('author')
        if author == tag.find('authors').findAll('author')[-1]:
            num_value = author.find('num').text if author.find('num') is not None else " "
            author_count.append(num_value)
        for author in tag.find('authors'):
            lang = author.get('lang')
            num = author.get('num')
            langArray.append([lang, num])

        for author in tag.find('authors').findAll('author'):
            second_loop_executed = False
            if len(langArray) != 1:
                for i in range(len(langArray) - 1):
                    if second_loop_executed:
                        break
                    else:
                        current_item = langArray[i]
                        next_item = langArray[i + 1]
                        langArray = langArray[1:]
                        if current_item[0] != next_item[0] and current_item[1] == next_item[1]:
                            break
                        else:
                            second_loop_executed = True
                            author_id = author.find('authorid').text if author.find('authorid') is not None else " "
                            author_name = author.find('lastname').text if author.find('lastname') is not None else ""
                            author_initials = author.find('initials').text if author.find('initials') is not None else ""
                            try:
                                for aff in author.find('affiliations'):
                                    org_id = aff.find('orgid').text if aff.find('orgid') is not None else " "
                                    org_name = aff.find('orgname').text if aff.find('orgname') is not None else " "
                                    pair = (author_id, org_id)
                                    if author_id != " " and org_id != " ":
                                        if pair not in unique_pairs:
                                            unique_pairs.add(pair)
                                            counter_all += 1
                                            count_author_org.append(counter_all)
                                            counter += 1
                                            counter_dict[pair] = counter
                                            author_organisation.append([counter, author_id, author_name,author_initials, org_id, org_name])
                                        else:
                                            count_author_org.append(counter_dict[pair])
                                    elif author_id == " " and org_id != " ":
                                        key = (author_name + ' ' + author_initials, org_id)
                                        if key not in unique_combinations:
                                         counter_all += 1
                                         counter += 1
                                         data_dict = {
                                             "counter": counter,
                                             "author_id": author_id,
                                             "author_name": author_name,
                                             "author_initials": author_initials,
                                             "org_id": org_id,
                                             "org_name": org_name,
                                             "linkurl": linkurl
                                         }
                                         # Append the dictionary to the list
                                         array_of_dicts.append(data_dict)
                                         count_author_org.append(counter_all)
                                         counter_dict_fornull_author[key] = counter
                                         unique_combinations.add(key)
                                         author_organisation.append([counter, author_id, author_name, author_initials, org_id, org_name])
                                        else:
                                            count_author_org.append(counter_dict_fornull_author[key])
                                    elif author_id != " " and org_id == " ":
                                        key = (author_id, org_name)
                                        if key not in unique_combinations:
                                         counter_all += 1
                                         counter += 1
                                         data_dict = {
                                             "counter": counter,
                                             "author_id": author_id,
                                             "author_name": author_name,
                                             "author_initials": author_initials,
                                             "org_id": org_id,
                                             "org_name": org_name,
                                             "linkurl": linkurl
                                         }
                                         # Append the dictionary to the list
                                         array_of_dicts.append(data_dict)
                                         count_author_org.append(counter_all)
                                         counter_dict_fornull_org[key] = counter
                                         unique_combinations.add(key)
                                         author_organisation.append([counter, author_id, author_name, author_initials, org_id, org_name])
                                        else:
                                            count_author_org.append(counter_dict_fornull_org[key])
                                    elif author_id == " " and org_id == " ":
                                        key = (author_name + ' ' + author_initials, org_id)
                                        # Check if the key is not in the set of unique combinations
                                        if key not in unique_combinations:
                                         counter_all += 1
                                         counter += 1
                                         data_dict = {
                                             "counter": counter,
                                             "author_id": author_id,
                                             "author_name": author_name,
                                             "author_initials": author_initials,
                                             "org_id": org_id,
                                             "org_name": org_name,
                                             "linkurl": linkurl
                                         }
                                         # Append the dictionary to the list
                                         array_of_dicts.append(data_dict)
                                         count_author_org.append(counter_all)
                                         counter_dict_fornull_author_and_org[key] = counter
                                         unique_combinations.add(key)
                                         author_organisation.append([counter, author_id, author_name, author_initials, org_id, org_name])
                                        else:
                                            count_author_org.append(counter_dict_fornull_author_and_org[key])
                            except TypeError:
                                continue
            else:
                author_id = author.find('authorid').text if author.find('authorid') is not None else " "
                author_name = author.find('lastname').text if author.find('lastname') is not None else ""
                author_initials = author.find('initials').text if author.find('initials') is not None else ""
                try:
                    for aff in author.find('affiliations'):
                        org_id = aff.find('orgid').text if aff.find('orgid') is not None else " "
                        org_name = aff.find('orgname').text if aff.find('orgname') is not None else " "
                        pair = (author_id, org_id)
                        if author_id != " " and org_id != " ":
                            if pair not in unique_pairs:
                                unique_pairs.add(pair)
                                counter_all += 1
                                count_author_org.append(counter_all)
                                counter += 1
                                counter_dict[pair] = counter
                                author_organisation.append(
                                    [counter, author_id, author_name, author_initials, org_id, org_name])
                            else:
                                count_author_org.append(counter_dict[pair])
                        elif author_id == " " and org_id != " ":
                            key = (author_name + ' ' + author_initials, org_id)
                            # Check if the key is not in the set of unique combinations
                            if key not in unique_combinations:
                                counter_all += 1
                                counter += 1
                                data_dict = {
                                    "counter": counter,
                                    "author_id": author_id,
                                    "author_name": author_name,
                                    "author_initials": author_initials,
                                    "org_id": org_id,
                                    "org_name": org_name,
                                    "linkurl": linkurl
                                }
                                # Append the dictionary to the list
                                array_of_dicts.append(data_dict)
                                count_author_org.append(counter_all)
                                counter_dict_fornull_author[key] = counter
                                unique_combinations.add(key)
                                author_organisation.append(
                                    [counter, author_id, author_name, author_initials, org_id, org_name])
                            else:
                                count_author_org.append(counter_dict_fornull_author[key])
                        elif author_id != " " and org_id == " ":
                            key = (author_id, org_name)
                            # Check if the key is not in the set of unique combinations
                            if key not in unique_combinations:
                                counter_all += 1
                                counter += 1
                                data_dict = {
                                    "counter": counter,
                                    "author_id": author_id,
                                    "author_name": author_name,
                                    "author_initials": author_initials,
                                    "org_id": org_id,
                                    "org_name": org_name,
                                    "linkurl": linkurl
                                }
                                # Append the dictionary to the list
                                array_of_dicts.append(data_dict)
                                count_author_org.append(counter_all)
                                counter_dict_fornull_org[key] = counter
                                unique_combinations.add(key)
                                author_organisation.append(
                                    [counter, author_id, author_name, author_initials, org_id, org_name])
                            else:
                                count_author_org.append(counter_dict_fornull_org[key])
                        elif author_id == " " and org_id == " ":
                            key = (author_name + ' ' + author_initials, org_id)
                            # Check if the key is not in the set of unique combinations
                            if key not in unique_combinations:
                                counter_all += 1
                                counter += 1
                                data_dict = {
                                    "counter": counter,
                                    "author_id": author_id,
                                    "author_name": author_name,
                                    "author_initials": author_initials,
                                    "org_id": org_id,
                                    "org_name": org_name,
                                    "linkurl": linkurl
                                }
                                # Append the dictionary to the list
                                array_of_dicts.append(data_dict)
                                count_author_org.append(counter_all)
                                counter_dict_fornull_author_and_org[key] = counter
                                unique_combinations.add(key)
                                author_organisation.append(
                                    [counter, author_id, author_name, author_initials, org_id, org_name])
                            else:
                                count_author_org.append(counter_dict_fornull_author_and_org[key])
                except TypeError:
                    continue
        fields['counter'].append(count_author_org)
    article = pd.DataFrame(data=fields)
    article = article.explode('counter')

    authors_organisations = pd.DataFrame(author_organisation,
                                         columns=['counter', 'author_id', 'author_name', 'author_initials', 'org_id', 'org_name'])
    authors_organisations['author_fullname'] = authors_organisations['author_name'] + ' ' + authors_organisations['author_initials']
    authors_organisations.to_excel('authors_organisations.xlsx', index=False)
    article.to_excel("article.xlsx", index=False)
    fd.close()

    unique_author_pairs = set()  # To keep track of unique author_name + author_initials pairs
    unique_org_names = set()  # To keep track of unique org_name values
    author_filtered_data = []
    org_filtered_data = []

    for data_dict in array_of_dicts:
        author_id = data_dict.get("author_id")
        author_name = data_dict.get("author_name")
        author_initials = data_dict.get("author_initials")
        org_id = data_dict.get("org_id")
        org_name = data_dict.get("org_name")

        # Check if author_id is None and the author_name + author_initials pair is unique
        if author_id == " " and (author_name, author_initials) not in unique_author_pairs:
            unique_author_pairs.add((author_name, author_initials))
            author_filtered_data.append({
                "author_name": data_dict["author_name"],
                "author_initials": data_dict["author_initials"],
                "author_fullname": data_dict["author_name"] + ' ' + data_dict["author_initials"],
                "linkurl": data_dict["linkurl"]
            })

        # Check if org_id is None and the org_name is unique
        if org_id == " " and org_name not in unique_org_names:
            unique_org_names.add(org_name)
            org_filtered_data.append({"org_id": data_dict["org_id"],
                "org_name": data_dict["org_name"]})


    # Create dataframes from the filtered data
    author_df = pd.DataFrame(author_filtered_data)
    org_df = pd.DataFrame(org_filtered_data)

    # Define Excel writer for author_filtered_data
    author_writer = pd.ExcelWriter('author_filtered_data.xlsx', engine='xlsxwriter')
    author_df.to_excel(author_writer, sheet_name='Author Filtered Data', index=False)



    # Save the Excel file for author_filtered_data
    author_writer._save()

    # Define Excel writer for org_filtered_data
    org_writer = pd.ExcelWriter('org_filtered_data.xlsx', engine='xlsxwriter')
    org_df.to_excel(org_writer, sheet_name='Org Filtered Data', index=False)

    # Save the Excel file for org_filtered_data 
    org_writer._save()


if __name__ == "__main__":
    parse_articles_to_excel('org_items_570_149195.xml')


