import os
import time
import psutil

def run_excel():
    while True:
        os.system(f'start excel possible_duplicate_people.xlsx')
        while True:
            time.sleep(1)
            excel_running = False
            for process in psutil.process_iter(attrs=['pid', 'name']):
                if "EXCEL.EXE" in process.info['name']:
                    excel_running = True
                    break
            if not excel_running:
                break
        break
    print("Excel file has been closed. Now, running additional code.")
