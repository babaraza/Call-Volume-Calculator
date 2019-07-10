import datetime
import zipfile
from pathlib import Path

in_count = 0
out_count = 0
call_month = ""


def do_search(compare_question):
    in_count = 0
    out_count = 0
    input_year = input("Enter year > ")
    input_month = input("Enter month or all > ")

    if input_month == 'all':
        call_dir = f"C:/Users/a058943/Desktop/CALLS/{input_year}"
        call_month = input_year
    else:
        x = datetime.datetime(int(input_year), int(input_month), 1)
        call_dir = f"C:/Users/a058943/Desktop/CALLS/{input_year}/{input_month}"
        call_month = x.strftime('%b')

    for filename in Path(call_dir).glob('**/*.wav'):
        if str(filename).endswith("IN.wav"):
            in_count += 1
        else:
            out_count += 1

    for filename_zip in Path(call_dir).glob('**/*.zip'):
        zip_file = zipfile.ZipFile(filename_zip, 'r')
        for name in zip_file.namelist():
            if str(name).endswith("IN.wav"):
                in_count += 1
            else:
                out_count += 1
        zip_file.close()
    show_results(call_month, input_year, in_count, out_count, compare_question)


def show_results(orig_month, orig_year, orig_incount, orig_outcount, comp_question):
    print(f"\nFor {orig_year} {orig_month}\n"
          f"Total incoming calls: {orig_incount:,}\n"
          f"Total outgoing calls: {orig_outcount:,}\n"
          f"Total calls for {orig_month}: {orig_incount + orig_outcount:,}")
    if comp_question == "y":
        print("")
        do_search("n")
    else:
        print("\nFinished")


compare_results = input("Compare results (y/n) > ")
if compare_results == "y":
    do_search("y")
else:
    do_search("n")

# TODO:
#   Create Excel File
#   Create Sheet for each year or table
#   Put results for each month