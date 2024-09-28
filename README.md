# Excel Column Cleaner Application

## Overview

The **Excel Column Cleaner Application** is a GUI-based tool designed to streamline the cleaning and formatting of Excel files. Users can upload Excel files, select specific columns to retain, and apply custom color formatting to specific content. The application also provides options for adjusting column widths and saving the cleaned files in a new format. This app is built with Python, pandas, CustomTkinter, and openpyxl, providing a user-friendly experience for handling Excel files.

## Features

- **Upload Excel Files**: Select any `.xlsx` file to process.
- **Column Selection**: Choose the specific columns you want to keep from the uploaded file.
- **Color Formatting**: Automatically apply alternating color formatting to rows based on changes in the "Custom Message" column.
- **Column Width Adjustment**: Automatically adjust column widths based on content length for better readability.
- **Persistent State**: Track unchecked columns across sessions and save this data for future use.
- **File Explorer Integration**: Open the directory containing the processed file after saving it.
- **CustomTkinter GUI**: User-friendly interface with dark and light themes, custom icons, and logo integration.

## How It Works

1. **Upload File**: The user selects an Excel file to be processed. 
2. **Select Columns**: All columns in the Excel file are displayed in a scrollable list, and users can select which columns to keep.
3. **Process File**: After selecting the columns, the app creates a new cleaned Excel file with only the selected columns.
4. **Color Formatting**: If the Excel file contains a column named "Custom Message" and the file name includes "Nurture", alternating colors are applied to rows where the message changes.
5. **Save and Open Directory**: The cleaned file is saved with a "- QC_CLEAN.xlsx" suffix, and the directory is opened for easy access to the saved file.

## Installation

### Prerequisites

Ensure you have the following installed on your machine:
- Python 3.x
- `pip` package manager

### Required Python Packages

You can install the required dependencies using the following command:

```bash
pip install pandas openpyxl customtkinter pillow
