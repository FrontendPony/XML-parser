import subprocess
import os
from datetime import datetime

def create_postgres_backup():
    # Change directory to where you want the command to run
    previous_directory = os.getcwd()
    os.chdir('C:/Program Files/PostgreSQL/15/bin')
    os.environ['PGPASSWORD'] = 'sword9999'

    # Generate the current date as a string to use in the file name
    current_datetime = datetime.now().strftime("%Y-%d-%m_%H-%M")

    # Construct the file path with the date in the file name
    file_path = f'C:/Program Files/PostgreSQL/15/data/backup_files/db_{current_datetime}.sql'

    # Command to execute
    command = [
        'pg_dump',
        '--username=postgres',
        '--dbname=xml-parser',
        '--host=localhost',
        '--username=postgres',
        '-f',
        file_path
    ]

    # Run the command
    try:
        subprocess.run(command, check=True, shell=True)
        os.chdir(previous_directory)
        return f"Backup created successfully at {file_path}"
    except subprocess.CalledProcessError as e:
        os.chdir(previous_directory)
        return f"Error: {e}"

