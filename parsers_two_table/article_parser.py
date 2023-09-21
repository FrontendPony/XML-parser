from bs4 import BeautifulSoup
import pandas as pd
from dbsettings import database_parametres
import psycopg2
import re
def extract_numbers_from_string(input_string):
    pattern = r'\d+'
    number = re.findall(pattern, input_string)
    number = [int(num) for num in number]
    return number

def parse_articles_to_excel(xml_filename):
    query = """
                         SELECT MAX(counter) FROM article
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

    if len(number) == 0:
        counter_all = 0
    else:
        counter_all = number[0]
    for tag in soup.findAll("item"):
        # item
        fields['item_id'].append(tag['id'])
        fields['linkurl'].append(tag.find('linkurl').text if tag.find('linkurl') is not None else "")
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
        # count of organisations
        for author in tag.find('authors').findAll('author'):
            for aff in author.descendants:
                if aff.name == 'orgname':
                    counter_all += 1
                    count_author_org.append(counter_all)
        fields['counter'].append(count_author_org)

    article = pd.DataFrame(data=fields)
    article = article.explode('counter')
    article.to_excel("article.xlsx", index=False)

    fd.close()


if __name__ == "__main__":
    parse_articles_to_excel('../xml_parser/article.xml')


