from dotenv import load_dotenv
from pathlib import Path
import zipfile
import utils
import os

load_dotenv()

# Directory that has all the calls/files
main_dir = os.getenv('DIR_PATH')
# Directory where to save file
save_dir = os.getenv('SAVE_PATH')
# Filename for JSON file
json_file = os.getenv('JSON_FILE')


# Look through Call Directory for wav/zip files
def generate_data():
    # Create empty dictionary
    raw_dictionary = {}

    calls_dir = Path(main_dir).glob('*[0-9]')
    # Returns the entire path of folder
    for years in calls_dir:
        # Getting last 4 digits of the year folder (ex: 2018)
        year_number = str(years)[-4:]
        raw_dictionary[year_number] = {}
        months = years.glob('*')
        for month in months:
            # Getting last 2 digits of the month folder (ex: 02 for Feb)
            month_number = str(month)[-2:]
            raw_dictionary[year_number][month_number] = []
            filenames = month.glob('*.*')
            for filename in filenames:
                # Open zipfile to get file count
                if str(filename).endswith(".zip"):
                    zip_file = zipfile.ZipFile(filename, 'r')
                    for name in zip_file.namelist():
                        raw_dictionary[year_number][month_number].append(str(name))
                    zip_file.close()
                elif str(filename).endswith(".wav"):
                    raw_dictionary[year_number][month_number].append(str(filename))
                # Any file extension other than .wav or .zip
                else:
                    print(f"File extension {str(filename)[-4:]} is not supported")
    return raw_dictionary


def parse_data(data):
    final_data = {}

    for year, call_log in data.items():
        print(f'Getting call data for {year}')
        # UNCOMMENT THIS FOR CREATING A YEAR KEY FOR EACH YEAR (1 OF 3)
        # final_data[year] = {}
        for month, calls in call_log.items():
            # UNCOMMENT THIS FOR CREATING A MONTH KEY FOR EACH MONTH (2 OF 3)
            # month = parse_month(month=month, abbreviated=False)
            # final_data[year][month] = {}
            for call in calls:
                data_time, call_id, voip, number, direction = call.split("-")
                date, time = data_time.split("_")
                call_date = utils.parse_date(date)
                call_time = utils.parse_time(time)
                call_number = utils.parse_call_number(number)
                call_direction = direction.split(".")[0]
                # UNCOMMENT THIS FOR CREATING A MONTH KEY FOR EACH MONTH (3 OF 3)
                # final_data[year][month].setdefault(call_date, {})
                # final_data[year][month][call_date].setdefault(call_direction, [])
                # final_data[year][month][call_date][call_direction].append(f'{call_time} - {call_number}')
                final_data.setdefault(call_date, [])
                final_data[call_date].append(f'{call_direction} - {call_time} - {call_number}')

    return final_data


utils.create_json(generate_data(), json_file)
raw_data = utils.load_json(json_file)
parsed_data = parse_data(raw_data)
# print(json.dumps(parsed_data, indent=2))
utils.create_excel(save_dir, 'Call Logs', parsed_data)
