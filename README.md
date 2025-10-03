# Table Fusion Utility

A Python utility to merge multiple XLSX tables into a single consolidated table.

Created for Roman Horobets.

## Overview

This utility provides a streamlined algorithm for merging Excel tables with improved accuracy and automatic header detection.

## Features

- **Automatic header detection** - intelligently finds header rows in Excel files
- **Accurate data extraction** - preserves all meaningful data without loss
- **Complete column mapping** - creates unified structure with all unique headers
- **Source tracking** - adds `source_file` column for data traceability
- **Improved algorithm** - handles complex Excel structures more reliably
- **Recursive search** - finds XLSX files in nested folder structures

## Installation

See also [guide.uk](guide.uk.md).

Install required packages:

```bash
pip install -r requirements.txt
```

Or install dependencies directly:

```bash
pip install pandas openpyxl
```

## Usage

Run the table fusion utility:

```bash
python table_fusion.py
```

The utility will:

- Read all XLSX files from the `data/` directory and all subfolders
- Process and merge them into a single table
- Save the result in the `result/` directory with timestamp format: `YYYY-MM-DD_HH-MM-SS.xlsx`

## Algorithm

1. **Structure Analysis**: Automatically detects header rows in each Excel file
   - Analyzes first 10 rows of each file
   - Looks for rows with minimum 5 non-empty values
   - Checks for typical headers (Title, Composer, Artist, Album)

2. **Header Extraction**: Creates unified set of all unique headers from all files

3. **Data Consolidation**: Collects data from all tables with proper column mapping
   - Reads data starting from header row
   - Removes empty rows
   - Preserves all meaningful data without loss

4. **Source Attribution**: Adds `source_file` column with source filename for each row

## Advantages

1. **Accuracy**: Correctly identifies headers and reads only meaningful data
2. **Completeness**: Preserves all data without loss
3. **Transparency**: Adds source information for each row
4. **Simplicity**: Automatically processes files without manual configuration

## Project Structure

```text
table-fusion/
├── data/                           # Source XLSX files
│   ├── SRC_1.xlsx
│   ├── folder1/
│   │   ├── SRC_2.xlsx
│   │   └── subfolder/
│   │       └── SRC_3.xlsx
│   └── folder2/
│       └── SRC_4.xlsx
├── result/                         # Output directory
│   └── YYYY-MM-DD_HH-MM-SS.xlsx
├── table_fusion.py                 # Main utility script
├── requirements.txt                # Python dependencies
└── README.md                      # This file
```

## Requirements

- Python 3.6+
- pandas
- openpyxl
