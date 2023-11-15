import pandas as pd

def update_rinc_ids(authors_organisations_file, authors_ref_file, sheet_name='РИНЦ ID'):
    df1 = pd.read_excel(authors_organisations_file)
    df2 = pd.read_excel(authors_ref_file, sheet_name=sheet_name)

    for index, row in df2.iterrows():
        author_fullname = row['Автор публикации']
        ids = row['РИНЦ ID']

        if author_fullname in df1['author_fullname'].values and ids == '-':
            matching_row = df1[df1['author_fullname'] == author_fullname]
            df2.at[index, 'РИНЦ ID'] = matching_row['author_id'].values[0]
        else:
            author_fullname_concat = row['фамилия'] + ' ' + row['имя'] + ' ' + row['отчество']

            if author_fullname_concat in df1['author_fullname'].values and ids == '-':
                matching_row = df1[df1['author_fullname'] == author_fullname_concat]
                df2.at[index, 'РИНЦ ID'] = matching_row['author_id'].values[0]

    df2.to_excel(authors_ref_file, sheet_name=sheet_name)



