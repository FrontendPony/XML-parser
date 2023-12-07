import pandas as pd
from delete_counter_in_article import update_counter_column
def deduplicate_excel(excel_file):
    try:
        df = pd.read_excel(excel_file)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        grouped = df.groupby(['author_id', 'org_id'])
        rows_to_delete = []
        values_to_delete = []
        for _, group in grouped:
            if len(group) > 1:
                min_counter = group['counter'].min()
                rows_to_delete.extend(group.loc[group['counter'] != min_counter, 'counter'])
                for value in rows_to_delete:
                    values_to_delete.append([min_counter, value])
        df = df[~df['counter'].isin(rows_to_delete)]
        if excel_file == 'merged_ao.xlsx':
            df = df.drop_duplicates(subset='counter', keep='first')
        df.to_excel(excel_file, index=False)
        update_counter_column('merged_link.xlsx', values_to_delete)

    except Exception as e:
        print(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    deduplicate_excel('merged_ao.xlsx')

