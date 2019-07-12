import zipfile
from pathlib import Path
import openpyxl as xl

test_array = {'2017': {'06': {'Incoming': 0, 'Outgoing': 1}}, '2018': {'01': {'Incoming': 355, 'Outgoing': 323}, '02': {'Incoming': 496, 'Outgoing': 483}, '03': {'Incoming': 663, 'Outgoing': 469}, '04': {'Incoming': 760, 'Outgoing': 865}, '05': {'Incoming': 762, 'Outgoing': 844}, '06': {'Incoming': 602, 'Outgoing': 886}, '07': {'Incoming': 460, 'Outgoing': 427}, '08': {'Incoming': 734, 'Outgoing': 776}, '09': {'Incoming': 619, 'Outgoing': 462}, '10': {'Incoming': 804, 'Outgoing': 642}, '11': {'Incoming': 593, 'Outgoing': 497}, '12': {'Incoming': 454, 'Outgoing': 382}}, '2019': {'01': {'Incoming': 672, 'Outgoing': 570}, '02': {'Incoming': 472, 'Outgoing': 359}, '03': {'Incoming': 553, 'Outgoing': 466}, '04': {'Incoming': 638, 'Outgoing': 368}, '05': {'Incoming': 447, 'Outgoing': 323}, '06': {'Incoming': 642, 'Outgoing': 422}}}
years_array = {}


def generate_filenames():
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
                else:
                    print(f"File extension {str(filename)[-4:]} is not supported")
            years_array[year_number][month_number]['Incoming'] = in_count
            years_array[year_number][month_number]['Outgoing'] = out_count
            # for row in range(4, sheet.max_row + 1):

# print(years_array)
# next(iter(years_array.keys()))
# list(years_array.keys())[0]
# for y in years_array: print(y)

# cell = sheet['a1'] or sheet.cell(1,1)
# for row in range(2, sheet.max_row + 1):
#   cell = sheet.cell(row, 3)
#   print(cell.value)

generate_filenames()

wb = xl.Workbook()
ws = wb.active
ws.title = 'Logs'

for years_in_array, months_in_array in test_array.items():
    ws.append([years_in_array])
    for subMonths, total_calls in months_in_array.items():
        ws.append([subMonths])
        for call_labels, call_counts in total_calls.items():
            ws.append([call_labels, call_counts])
        ws.append([""])

wb.save('C:/Users/a058943/Desktop/CALLS/Log Test.xlsx')



# TODO:
#   Create Excel File
#   Create Sheet for each year or table
#   Put results for each month