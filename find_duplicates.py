import pandas as pd
from delete_counter_in_article import delete_rows_in_excel
def deduplicate_excel(excel_file):
    try:
        df = pd.read_excel(excel_file)
        df = df.drop("Unnamed: 0", axis=1)
        print(df.columns)
        grouped = df.groupby(['author_id', 'org_id'])
        rows_to_delete = []

        for _, group in grouped:
            if len(group) > 1:
                print(group)
                rows_to_delete.extend(group.loc[group['counter'] != group['counter'].min(), 'counter'])
        print(rows_to_delete)
        df = df[~df['counter'].isin(rows_to_delete)]
        if excel_file == 'merged_ao.xlsx':
            print(228)
            df = df.drop_duplicates(subset='counter', keep='first')
        df.to_excel(excel_file, index=False)
        delete_rows_in_excel('article_authors_linkage.xlsx', rows_to_delete)

    except Exception as e:
        print(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    deduplicate_excel('merged_ao.xlsx')

