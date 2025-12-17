# Monthly Media Monitoring Report Compiler

This is a Python script to automatically compile monthly media monitoring reports from multiple Word files into one neat Excel file with multiple sheet categories.

## Features

- Automatically search for all `.docx` files inside folders and sub-folders.
- Extracts news data from tables inside Word files.
- Combines multiple news categories into one main category.
- Generates a single Excel file with separate sheets for each news category.
- Format headers (yellow, center-aligned) and cells (wrap text).
- Set column widths automatically.

## How to use

1.  Make sure all required libraries are installed: `pip install pandas python-docx xlsxwriter`
2.  Open the file `monthly_compiler.py`.
3.  Change the path in the `root_folder_path` variable to point to the main folder of the month you want to process.
4.  Run the script: `python monthly_compiler.py`

## Acknowledgements

This script was developed by myself with help and guidance from Google's AI assistant, Gemini.

