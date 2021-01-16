from utils import save_file, parse_month, parse_date, parse_time
from dotenv import load_dotenv
from pathlib import Path
import zipfile
import json
import os

load_dotenv()

# Path to the call logs
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


def create_json():
    # Creating a json file with complete call data:
    with open('data.json', 'w') as file:
        json.dump(generate_data(), file)


def load_json():
    with open('data-with-time.json', 'r') as file:
        data = json.load(file)
    return data


def create_excel(data):
    # Creating an excel file with json data
    save_file(main_dir, 'Call Logs', data)


# TODO: parse data
final_data = []

for year, call_log in load_json().items():
    print(year)
    for month, calls in call_log.items():
        print(parse_month(month=month, abbreviated=True))
        for call in calls[0:1]:
            call_data = call.split("-")
            call_date = parse_date(call_data[0].split("_")[0])
            call_time = parse_time(call_data[0].split("_")[1])
            call_number = call_data[1]
            call_direction = call_data[2].split(".")[0]
            print(f'{call_direction} - On {call_date} at {call_time} from {call_number}')
