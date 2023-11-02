
# DBC-Excel Converter

## Overview

The DBC-Excel Converter is a Python application that allows you to convert DBC files to Excel and Excel files to DBC format. This can be especially useful for working with CAN databases and signal definitions.

## Features

- **Convert DBC to Excel**: Import DBC files and export the data to Excel format.
- **Convert Excel to DBC**: Import Excel files and generate DBC files with mapping.
- **User-Defined Mappings**: Easily map DBC parameters to column indexes.

## Installation

1. Clone the repository: `git clone https://github.com/pardhu_prd/dbc-excel-converter.git`
2. Install the required dependencies: `pip install -r requirements.txt`

## Usage

1. Run the application: `python main.py`
2. Click on the "DBC to Excel" or "Excel to DBC" button to select the input file.
3. Map the DBC parameters to column indexes in the Excel file.
4. Click the "Convert" button to perform the conversion.
5. The output file will be displayed in the "Output" section.

## Requirements

- Python 3
- PyQt5
- pandas
- cantools


