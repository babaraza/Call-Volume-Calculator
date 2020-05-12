[![Python 3.8](https://img.shields.io/badge/Python-3.8-blue.svg)](https://www.python.org/downloads/release/python-38/)

# Call Volume Calculator
Calculate Call Volume by scanning Recorded `CALLS` Folder

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
    - Skip to below
- If its a `.wav` file:
    - Check if the name of the file has IN or OUT
        - **IN** = incoming calls
        - **OUT** = outgoing calls
- Track count of incoming and outgoing calls using a dictionary
- Create a dataframe of the dictionary above using `Pandas`
- Check if excel `filename` exists
    - If true, create new **worksheet**
    - If false, create new **file**