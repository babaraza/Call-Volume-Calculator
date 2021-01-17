from utils import create_excel, parse_month, parse_date, parse_time, parse_call_number
from dotenv import load_dotenv
from pathlib import Path
import zipfile
import json
import os

load_dotenv()

# Path to the call logs
main_dir = os.getenv('DIR_PATH')
# Path to the directory where to save file
save_dir = os.getenv('SAVE_PATH')


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


def create_json(data, filename):
    # Creating a json file with complete call data:
    with open(filename, 'w') as file:
        json.dump(data, file)


def load_json(filename):
    with open(filename, 'r') as file:
        data = json.load(file)
    return data


def parse_data(data):
    final_data = {}

    for year, call_log in data.items():
        print(year)
        # UNCOMMENT THIS FOR CREATING A YEAR KEY FOR EACH YEAR (1 OF 3)
        # final_data[year] = {}
        for month, calls in call_log.items():
            # UNCOMMENT THIS FOR CREATING A MONTH KEY FOR EACH MONTH (2 OF 3)
            # month = parse_month(month=month, abbreviated=False)
            # final_data[year][month] = {}
            for call in calls[0:5]:
                call_data = call.split("-")
                call_date = parse_date(call_data[0].split("_")[0])
                call_time = parse_time(call_data[0].split("_")[1])
                call_number = parse_call_number(call_data[1])
                call_direction = call_data[2].split(".")[0]
                # UNCOMMENT THIS FOR CREATING A MONTH KEY FOR EACH MONTH (3 OF 3)
                # final_data[year][month].setdefault(call_date, {})
                # final_data[year][month][call_date].setdefault(call_direction, [])
                # final_data[year][month][call_date][call_direction].append(f'{call_time} - {call_number}')
                final_data.setdefault(call_date, [])
                final_data[call_date].append(f'{call_direction} - {call_time} - {call_number}')

    return final_data


raw_data = load_json('data-with-time.json')
parsed_data = parse_data(raw_data)
# print(json.dumps(parsed_data, indent=2))
create_excel(save_dir, 'Call Logs', parsed_data)
