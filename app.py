import zipfile
from pathlib import Path
import openpyxl as xl

def generate_filenames():
    years_array = {}
    calls_dir = Path("C:/Users/a058943/Desktop/CALLS/").glob('*[0-9]')
    for years in calls_dir:
        year_number = str(years)[-4:]
        years_array[year_number] = {}
        months = years.glob('*')
        for month in months:
            in_count = 0
            out_count = 0
            month_number = str(month)[-2:]
            years_array[year_number][month_number] = {}
            filenames = month.glob('*.*')
            for filename in filenames:
                if str(filename).endswith(".zip"):
                    zip_file = zipfile.ZipFile(filename, 'r')
                    for name in zip_file.namelist():
                        if str(name).endswith("IN.wav"):
                            in_count += 1
                        else:
                            out_count += 1
                    zip_file.close()
                elif str(filename).endswith(".wav"):
                    if str(filename).endswith("IN.wav"):
                        in_count += 1
                    else:
                        out_count += 1
            years_array[year_number][month_number]['Incoming'] = in_count
            years_array[year_number][month_number]['Outgoing'] = out_count
    print(years_array)

generate_filenames()

# TODO:
#   Create Excel File
#   Create Sheet for each year or table
#   Put results for each month