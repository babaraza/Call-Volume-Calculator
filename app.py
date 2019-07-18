import zipfile
from pathlib import Path
import openpyxl as xl
from datetime import datetime

# Create empty dictionary
years_array = {}

# Look through Call Directory for wav/zip files
def generate_filenames():
    calls_dir = Path("C:/Users/a058943/Desktop/CALLS/").glob('*[0-9]')
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
                else:
                    print(f"File extension {str(filename)[-4:]} is not supported")  # Any file extension other than .wav or .zip
            years_array[year_number][month_number]['Incoming'] = in_count
            years_array[year_number][month_number]['Outgoing'] = out_count


# Save an Excel file with results
def save_file(filename):
    save_filename = 'C:/Users/a058943/Desktop/CALLS/' + filename + '.xlsx'
    check_path = Path(save_filename)

# Checking if file already exists
    if check_path.exists():
        print(f'{filename} already exists, creating new sheet')
        wb = xl.load_workbook(save_filename)
        ws = wb.create_sheet()
        ws.title = datetime.today().strftime('%m-%d-%y')
    else:
        print(f'{filename} doesnt exist, creating new file')
        wb = xl.Workbook()
        ws = wb.active
        ws.title = datetime.today().strftime('%m-%d-%y')

# Putting data into the Excel Sheet
    for years_in_array, months_in_array in years_array.items():
        ws.append([years_in_array])
        for subMonths, total_calls in months_in_array.items():
            ws.append([subMonths])
            for call_labels, call_counts in total_calls.items():
                ws.append([call_labels, call_counts])
            ws.append([""])

    wb.save(save_filename)  # Save Excel file

generate_filenames()
save_file('Call Logs')
