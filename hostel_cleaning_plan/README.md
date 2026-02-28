# Hostel Cleaning Plan Generator

This repository contains two Python scripts for generating hostel cleaning plans in PDF format based on Excel data. The scripts create clear tables indicating the status of each bed and room. The goal of the project is to simplify cleaning planning and visually display the current status of rooms and beds.

## Repository Structure
`hostel_cleaning_plan/` — root folder of the repository.

Inside the `scripts` folder:

- `hostel_cleaning_plan_symbol_version.py` — generates a PDF with symbols  
- `hostel_cleaning_plan_color_version.py` — uses colored blocks for different statuses  

The `example(xlsx)` folder contains a sample Excel source file:

- `template(house_keeping).xlsx`

The `output(examples)` folder is intended for generated PDFs, such as:

- `plan_symbol_version.pdf`
- `plan_color_version.pdf`

`README.md` is located in the root and contains all instructions and explanations.

## Installation and Running

To run the scripts, **Python 3.8 or higher** is required, along with the following libraries:

- `pandas`
- `reportlab`

Install dependencies:

```bash
pip install pandas reportlab
```

After that, the script can be executed from the command line.

The Excel file with room data can be placed in the **Downloads** folder, from which the script will automatically select the most recently downloaded file.

## Using a Specific Excel File

By default, the script searches for the latest Excel file in the **Downloads** folder.

If you need to use a specific file, open the script in an editor and find this block:

```python
downloads = os.path.join(os.path.expanduser("~"), "Downloads")
excel_files = glob.glob(os.path.join(downloads, "*.xlsx"))
if not excel_files:
    raise FileNotFoundError("No Excel (.xlsx) files found in your Downloads folder.")
latest_excel = max(excel_files, key=os.path.getmtime)
excel_file = latest_excel
```
Remove it and replace with the full path to your own Excel file:
```
excel_file = r"C:\Users\User\Documents\my_rooms.xlsx"
```
## Tips and Recommendations

The Excel file must be properly formatted with room names and statuses in a format recognized by the scripts.

It is not recommended to modify the script logic if the user is not familiar with Python. It is sufficient to replace the file selection block when using a specific Excel file.

To generate different PDF versions, you can use the same Excel file — the results will differ only in the visualization method (symbols or colors).

For new users, it is recommended to first test the script with the provided example:
- `template(house_keeping).xlsx`
## Display Logic and Visual Structure of the PDF

Each PDF is built as five vertical sections (by room ranges), which makes it convenient to print and use in housekeeping work.

**Symbol Version**

- The first cell of the left column contains the room number.

- The right column displays the status of each bed or the entire room.

- For private rooms, a single row is displayed.

**Color Version**

- The left column contains the room number (only for private rooms).

- For dorm rooms, the room number is located in the first long cell at the top spanning the width of both columns.

- For private rooms, a single row is displayed.
## Meaning of Visual Elements
**Symbol Version**

- `+` — Stayover

- Black block — Turnover / Check-out

**Color Version**

- Blue block — Stayover

- Red block — Turnover / Check-out

If a cell is empty, it means:

- The bed is free

- There is no cleaning status
## Practical Use of Left Empty Cells

Empty left cells can be used by staff as working space:

- For manual markings with a marker after printing

- For setting cleaning priorities

- For marking completed work

- For internal labeling (for example: deep cleaning, inspection, maintenance)

- For temporary notes that do not affect Excel data

Thus, the PDF serves not only as a report but also as a working tool during a shift.
## General Project Architecture

The project is built according to the principle:
`Excel → status processing → internal distribution logic → PDF generation`
#### Key Features

- Automatic detection of the most recently downloaded Excel file

- Separation of statuses into:

  - Room-level status

  - Bed-level status

- Support for different room types:

  - 3, 4, 6, 10 beds

  - Private double

  - Private twin

- Dynamic rendering of elements inside cells (via a custom Flowable)

- Automatic addition of document generation time

The scripts do not depend on external APIs and run entirely locally.
## Supported Statuses

The scripts process the following statuses:

- Stayover

- Turnover

- Check-out

The **Not Reserved** status is intentionally ignored, as it does not require housekeeping action.

If necessary, new statuses can be added to the processing logic.

## Scalability

The project can be adapted:

- To a different room numbering format

- To a different number of beds

- To a specific brand color scheme

- To integration into internal operational processes

The room distribution logic is defined by lists at the beginning of the script, which makes it easy to modify the structure without rewriting the entire program.
## Project Purpose

This project was created as a practical tool for automating hostel operational work.

Its goals are:

- To reduce the time required to prepare cleaning plans

- To minimize errors in manual list preparation

- To visualize work priorities

- To standardize internal processes

The project is oriented toward real working conditions of housekeeping staff and administrators.

## Contact and Feedback

If you have questions or suggestions for improving the scripts, you can create an issue in the repository or submit a pull request.

This repository is intended to simplify the cleaning planning process and can be adapted for other hostels or hotel properties.
