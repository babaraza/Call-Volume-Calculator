from openpyxl import load_workbook
from dotenv import load_dotenv
from datetime import datetime
from pathlib import Path
import pandas as pd
import zipfile
import os

load_dotenv()

# Create empty dictionary
years_array = {}

# Path to the dictionary
main_dir = os.getenv('DIR_PATH')


# Look through Call Directory for wav/zip files
def generate_data():
    calls_dir = Path(main_dir).glob('*[0-9]')
    for years in calls_dir:  # Returns the entire path of folder
        year_number = str(years)[-4:]  # Getting last 4 digits of the year folder (ex: 2018)
        years_array[year_number] = {}
        months = years.glob('*')
        for month in months:
            in_count = 0
            out_count = 0
            month_number = str(month)[-2:]  # Getting last 2 digits of the month folder (ex: 02 for Feb)
            years_array[year_number][month_number] = {}
            filenames = month.glob('*.*')
            for filename in filenames:
                if str(filename).endswith(".zip"):  # Open zipfile to get file count
                    zip_file = zipfile.ZipFile(filename, 'r')
                    for name in zip_file.namelist():
                        if str(name).endswith("IN.wav"):  # IN.wav are incoming calls
                            in_count += 1
                        else:
                            out_count += 1  # OUT.wav would be outgoing calls
                    zip_file.close()
                elif str(filename).endswith(".wav"):
                    if str(filename).endswith("IN.wav"):  # IN.wav are incoming calls
                        in_count += 1
                    else:
                        out_count += 1  # OUT.wav would be outgoing calls
                else:  # Any file extension other than .wav or .zip
                    print(f"File extension {str(filename)[-4:]} is not supported")
            years_array[year_number][month_number]['Incoming'] = in_count
            years_array[year_number][month_number]['Outgoing'] = out_count


# Save File to Excel
def save_file(filename):
    # Creating Data Frame with Data from Dictionary
    final_df = pd.concat({k: pd.DataFrame(v).transpose() for k, v in years_array.items()}, sort=True, axis=1)

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


generate_data()
save_file('Call Logs')
