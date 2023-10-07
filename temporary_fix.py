import pandas as pd

def read_and_clean_excel(file_path):
    data = pd.read_excel(file_path, index_col=0)
    df = pd.DataFrame(data)
    df = df.drop_duplicates()
    return df

