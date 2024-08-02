import subprocess
import os
from PyQt6.QtWidgets import QFileDialog, QApplication
from PyQt6.QtWidgets import QMessageBox
from clear_database import delete_data
def restore_database(self, file_path):
    previous_directory = os.getcwd()
    os.chdir('C:/Program Files/PostgreSQL/15/bin')
    os.environ['PGPASSWORD'] = 'sword9999'
    restore_command = [
        'psql',
        '-U',
        'postgres',
        '-d',
        'xml-parser',
        '-f',
        file_path
    ]
    delete_data()
    try:
        subprocess.run(restore_command, shell=True)
        os.chdir(previous_directory)
        QMessageBox.information(self, "Удачное восстановление", "База данных была восстановлена")
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        os.chdir(previous_directory)
        QMessageBox.information(self, "Неудачное восстановление", f"Произошла ошибка {e} !")


