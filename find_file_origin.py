import pandas as pd
def update_excel_file(input_file_path):
    df = pd.read_excel(input_file_path, index_col= 0)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df['data_origin'] = 'excel'
    if len(df) == 2:
        df.loc[df.index[0], 'data_origin'] = 'excel'
        df.loc[df.index[1], 'data_origin'] = 'sql'
        df.to_excel(input_file_path)
    else:
        for i in range(1, len(df)):
            if df.index[i - 1] > df.index[i]:
                df.iat[i, 26] = 'sql'
        df.to_excel(input_file_path)
        df = pd.read_excel(input_file_path)
        if 'sql' in df['data_origin'].values:
            sql_index = df[df['data_origin'] == 'sql'].index[0]
            print(sql_index)
            for index in range(sql_index + 1, len(df)):
                df.at[index, 'data_origin'] = 'sql'
        df.to_excel(input_file_path)

if __name__ == "__main__":
    update_excel_file('merged.xlsx')



