[![Python 3.8](https://img.shields.io/badge/Python-3.8-blue.svg)](https://www.python.org/downloads/release/python-38/)

# Call Volume Calculator
Calculate Call Volume by scanning Recorded `CALLS` in user specified folder

> Script supports scrapping through folder for `.wav` files and `.zip` files

> The .wav file names should end in `-OUT` for outgoing calls and `-IN` for incoming calls

### Usage

Specify directory to scan in `app.py`
Pass filename as argument when calling `save_file()`
Run `app.py`

**Script will follow these steps:**

- Look through `CALLS` Directory for year folders (ex: 2017, 2018, 2019...)
- Find all the month folders (ex: 01, 02, 03...)
- Scan all files inside each month folder using `glob()`
- If its a `.zip` file:
    - Open the zip file(s) and scan all `.wav` files inside
    - Go to `generate_data`
- If its a `.wav` file:
    - Go to `generate_data`
- **generate_data**
    - Create a `raw_dictionary` 
    - Create dictionary key with year and month 
    - Append the filename to the correct key
    - Return `raw_dictionary`
- Save the data into a JSON file
- Load the JSON file above and then parse the data using custom parser
- Flatten the JSON into a list and create a `dataframe` using `Pandas`
- Check if excel `filename` exists
    - If true, create new **worksheet** (with current date as name) into the existing excel file
    - If false, create new excel **file**