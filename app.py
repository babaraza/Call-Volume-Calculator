from openpyxl import load_workbook
from dotenv import load_dotenv
from datetime import datetime
from pathlib import Path
import pandas as pd
import datetime
import zipfile
import json
import os

load_dotenv()

# Path to the dictionary
main_dir = os.getenv('DIR_PATH')


# Look through Call Directory for wav/zip files
def generate_data():
    # Create empty dictionary
    years_array = {}

    calls_dir = Path(main_dir).glob('*[0-9]')
    # Returns the entire path of folder
    for years in calls_dir:
        # Getting last 4 digits of the year folder (ex: 2018)
        year_number = str(years)[-4:]
        years_array[year_number] = {}
        months = years.glob('*')
        for month in months:
            # Getting last 2 digits of the month folder (ex: 02 for Feb)
            month_number = str(month)[-2:]
            years_array[year_number][month_number] = []
            filenames = month.glob('*.*')
            for filename in filenames:
                # Open zipfile to get file count
                if str(filename).endswith(".zip"):
                    zip_file = zipfile.ZipFile(filename, 'r')
                    for name in zip_file.namelist():
                        years_array[year_number][month_number].append(str(name))
                    zip_file.close()
                elif str(filename).endswith(".wav"):
                    years_array[year_number][month_number].append(str(filename))
                # Any file extension other than .wav or .zip
                else:
                    print(f"File extension {str(filename)[-4:]} is not supported")
    return years_array


# Save File to Excel
def save_file(filename, data):
    # Creating Data Frame with Data from Dictionary
    final_df = pd.concat({k: pd.DataFrame(v).transpose() for k, v in data.items()}, sort=True, axis=1)

    save_filename = main_dir + filename + '.xlsx'
    print(f'Working Directory: {main_dir}')
    check_path = Path(save_filename)

    # Checking if file already exists
    if check_path.exists():
        print(f'{filename}.xlsx already exists, creating new sheet')
        book = load_workbook(save_filename)
        writer = pd.ExcelWriter(save_filename, engine='openpyxl')
        writer.book = book
        # Putting data into the Excel Sheet
        final_df.to_excel(writer, sheet_name=datetime.today().strftime('%m-%d-%y'))
        writer.save()
        writer.close()
    else:
        print(f'{filename}.xlsx doesnt exist, creating new file')
        # Putting data into the Excel Sheet
        final_df.to_excel(save_filename, sheet_name=datetime.today().strftime('%m-%d-%y'))


def create_json():
    # Creating a json file with complete call data:
    with open('data.json', 'w') as file:
        json.dump(generate_data(), file)


def create_excel():
    # Creating an excel file with json data
    with open('data.json', 'r') as file:
        data = json.load(file)

    save_file('Call Logs', data)


# TODO: parse data

def parse_time(time_string: str):
    """
    125716 remove last 2 digits 1257 + 200 = 1457 (- 1200) = 2:57 pm
    074433 remove last 2 digits 0744 + 200 = 0944 = 9:44 am
    """

    time_string = int(time_string[:-2])
    time_string += 200

    return datetime.datetime.strptime(str(time_string), '%H%M').strftime('%I:%M %p')
