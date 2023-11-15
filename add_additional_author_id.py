import pandas as pd

def update_additional_author_id(df_link, data_list_to_delete):
    df = pd.read_excel(df_link, index_col=0)
    if "Unnamed: 0" in df.columns:
        df = df.drop("Unnamed: 0", axis=1)
    for target_counter_value, new_value in data_list_to_delete:
        target_counter_value = int(target_counter_value)
        new_value = int(new_value)
        df.loc[df['author_id'] == target_counter_value, 'additional_author_id'] = new_value

    df.to_excel(df_link)

