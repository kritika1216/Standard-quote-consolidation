Standard Quote Consolidation Script
Project Description (300 characters)
A Python script that consolidates and standardizes multiple manufacturer or portfolio Excel files into a single unified dataset. It dynamically detects headers, normalizes columns, cleans numeric data (e.g., GST, cost), and outputs a formatted Excel for easy analysis.

Overview
This script automates the consolidation and normalization of price quote data from diverse manufacturer or portfolio Excel files. It ensures consistent data structure and formatting across files sourced from different teams or vendors, enabling smoother downstream reporting and analytics.

Features
Batch reads all Excel files in a specified input folder

Dynamically locates header rows in each sheet for flexible input formats

Normalizes and standardizes column names based on a predefined output schema

Cleans text and numeric fields, rounds decimals, handles missing or malformed GST values

Consolidates all processed data into a unified output Excel file

Applies Excel color formatting for headers to enhance visibility and review

Logs errors and warnings for missing data or mapping issues

Getting Started
Prerequisites
Python 3.x

pandas

numpy

openpyxl

Install required Python packages using pip:

bash
pip install pandas numpy openpyxl
Installation
Clone this repository or download the script directly.

Usage
Place all manufacturer or portfolio Excel files to consolidate within a single input folder.

Modify the script to set:

inputfolder: Path to the input directory containing Excel files.

outputfile: Desired output Excel file path and name.

outputheaders: List of desired standardized output columns.

Run the script:

bash
python standard_quote-Consolidation.py
The consolidated and cleaned Excel file will be generated at the specified output location.

File Structure
text
/inputfolder/          # Contains input Excel files
standard_quote-Consolidation.py   # Main consolidation script
/Outputs/              # Output folder for consolidated Excel files
README.md              # Project documentation
How it Works
The script iterates through each Excel file and sheet.

Detects the row containing the header based on expected columns.

Maps varied input column names to a consistent set of output headers.

Normalizes data types and cleans numeric fields, correcting GST and rounding decimals.

Aggregates all cleaned data into a single pandas DataFrame.

Writes the final unified DataFrame to an Excel file with header color formatting.

Contributing
Contributions are welcome! Please fork the repository and submit pull requests for enhancements or bug fixes.

Support
For issues or questions, please open an issue in this GitHub repository or contact the maintainer.

License
This project is licensed under the MIT License.
