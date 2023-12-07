import pandas as pd

def update_counter_column(file_path, new_values):
    df = pd.read_excel(file_path)

    for idx, val in enumerate(new_values):
        df.loc[df['counter'] == val[1], 'counter'] = val[0]

    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    df.to_excel(file_path, index=False)


