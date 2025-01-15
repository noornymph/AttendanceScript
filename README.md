# Meeting Attendance Compiler Script

## Overview

This script processes attendance data from meeting reports contained in a ZIP file. It compiles and analyzes the data, generates summaries of attendance percentages, and considers employee leave data from an Excel file. The script outputs the compiled data in Excel format.

## Features

- **ZIP File Extraction**: Extracts meeting reports from a ZIP file.
- **Attendance Data Compilation**: Reads attendee information from individual Excel files and compiles the data.
- **Leave Data Integration**: Considers employees on leave (provided in an Excel file) while calculating attendance.
- **Attendance Summary**: Outputs overall and individual attendance percentages.
- **Error Handling**: Robust error handling for file processing.

## Requirements

- Python 3.8 or above
- Libraries:
  - `openpyxl`
  - `argparse`
  - `datetime`
  - `zipfile`
  - `os`
  - `tempfile`

Install required libraries using:
```bash
pip install openpyxl
```
