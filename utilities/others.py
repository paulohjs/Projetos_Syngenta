import os
import pandas as pd
import json
from time import sleep

def cleanup_single_folder(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)

def convert_xl_to_dataframe(xl_file_path):
    xl_sheet = pd.read_excel(xl_file_path)
    df_excel = pd.DataFrame(xl_sheet)
    return df_excel

def get_email_password(index):
    """
    Retrieves the email and password for a user from a JSON file based on the given index.

    Args:
        index (int): The index of the user in the JSON file.

    Returns:
        tuple: A tuple containing the email and password of the user.
    """

    from pathlib import Path
    folder_path = Path.cwd()

    with open(f'{folder_path}\\user.json', 'r') as arquivo_user:
        login_data = json.load(arquivo_user)
    email = login_data['Users'][index]['email']
    password = login_data['Users'][index]['password']
    return email, password

def cleanup_folders(destination_path_file, origin_path_file):
    """
    Cleans up the folders by removing existing files in the destination and origin paths.

    Args:
        destination_path_file (str): The file path in the destination folder.
        origin_path_file (str): The file path in the origin folder (usually the Downloads folder).
    """

    if os.path.exists(origin_path_file):
        os.remove(origin_path_file)
    if os.path.exists(destination_path_file):
        os.remove(destination_path_file)
    print('\033[32mFolders cleaned!\033[m')

def move_file_dowloaded(origin, destination):
    """
    Moves the downloaded file from the origin path to the destination path.

    Args:
        origin (str): The original file path where the file was downloaded.
        destination (str): The destination file path where the file should be moved.
    """

    while True:
        sleep(2)
        try:
            os.rename(origin, destination)
            print('\033[32mFile downloaded and moved!\033[m')
            sleep(2)
            break
        except:
            None
    sleep(2)

def get_smartsheet_key():

    from pathlib import Path
    folder_path = Path.cwd()

    with open(f'{folder_path}\\smartsheet_key.json', 'r') as json_key:
        login_data = json.load(json_key)
    key = login_data['key']
    return key