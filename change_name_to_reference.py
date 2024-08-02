import pandas as pd

def update_df1_with_df2(df1_file, df2_file):
    df2 = pd.read_excel(df2_file, sheet_name='РИНЦ ID')
    df1 = pd.read_excel(df1_file, index_col=0)

    for index, row in df1.iterrows():
        if row['author_id'] in df2['РИНЦ ID'].tolist():
            matching_row = df2[df2['РИНЦ ID'] == row['author_id']].iloc[0]
            df1.at[index, 'author_name'] = matching_row['фамилия']
            df1.at[index, 'author_initials'] = matching_row['имя'] + ' ' + matching_row['отчество']

    df1.to_excel(df1_file)



