import pandas as pd

def merge_authors_by_enter_id(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
        df_old = pd.read_excel(output_file)
        grouped = df.groupby('enter_id')['author_fullname'].apply(lambda x: ', '.join(x)).reset_index()
        grouped = grouped.rename(columns={'author_fullname': 'new_column_name'})
        grouped = grouped[grouped['new_column_name'].str.count(',') >= 1]
        merged_df = pd.concat([grouped, df_old], ignore_index=True)
        merged_df['new_column_name'] = merged_df.groupby('enter_id')['new_column_name'].transform(
            lambda x: ', '.join(x))
        merged_df['new_column_name'] = merged_df['new_column_name'].apply(
            lambda x: ', '.join(sorted(set(x.split(', ')))))
        merged_df.drop_duplicates(subset=['enter_id'], inplace=True)
        merged_df.reset_index(drop=True, inplace=True)
        print(merged_df)
        merged_df.to_excel(output_file, index=False)
        return True
    except Exception as e:
        return str(e)


if __name__ == "__main__": merge_authors_by_enter_id('author_filtered_data.xlsx','alternative_names.xlsx')