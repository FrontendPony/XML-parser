import pandas as pd
from delete_counter_in_article import delete_rows_in_excel
def deduplicate_excel(excel_file):
    df = pd.read_excel(excel_file)
    grouped = df.groupby(['author_id', 'org_id'])
    deleted_counters = []
    rows_to_delete = []

    for _, group in grouped:
        if len(group) > 1:
            deleted_counters.extend(group['counter'].iloc[1:].tolist())
            rows_to_delete.extend(group.index[1:])
    print(rows_to_delete)
    result_df = df.drop(rows_to_delete)
    result_df.to_excel(excel_file, index=False)
    delete_rows_in_excel('article.xlsx', rows_to_delete)


if __name__ == "__main__":
    deduplicate_excel('authors_organisations.xlsx')

