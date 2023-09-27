import pandas as pd

def delete_rows_in_excel(excel_file_path, rows_to_delete):
    try:
        df = pd.read_excel(excel_file_path)
        df = df[~df['counter'].isin(rows_to_delete)]
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        print("Rows with specified counter values deleted successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")


