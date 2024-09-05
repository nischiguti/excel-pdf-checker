# Equipment PDF and Excel Processing Script

## Overview

This script processes folders containing equipment data by extracting and categorizing information from Excel and PDF files. It performs the following tasks:
1. **Extracts serial numbers** from Excel files.
2. **Checks PDFs** for specific equipment and serial numbers.
3. **Logs results** into categorized text files based on findings.

## Features

- **PDF Text Extraction**: Extracts text from PDF files to check for equipment and serial numbers.
- **Excel Data Extraction**: Retrieves serial numbers from Excel files based on a specific label.
- **Categorized Logging**: Saves results into different text files based on the presence of equipment and serial numbers.

## Prerequisites

Ensure you have the following Python libraries installed:
- `openpyxl` (for reading Excel files)
- `PyPDF2` (for reading PDF files)

Install these libraries using pip:

```bash
pip install openpyxl PyPDF2
```
## Setup

1. **Directory Structure**: Organize your directories as follows:
- A root directory containing subfolders named by equipment numbers.
- Each equipment folder should contain an Excel file with 'Ficha' in its name and a subfolder named 'calibração' containing PDF files.

2. **Script Configuration**: Update the script to reflect your actual directory paths.
- `root_dir`: Path to the root directory containing all equipment folders.
- `output_files`: Dictionary mapping result categories to output file paths.

## Usage

1. **Configure the Script**: Edit the script to set the correct paths for `root_dir` and `output_files`.
2. **Run the Script**: Execute the script using Python. For example:
```bash
python your_script_name.py
```
3. **Review Results**: After the script completes, check the specified output files for categorized results:

- **Contains both equipment number and serial number.**
- **Contains equipment number but not serial number.**
- **Contains serial number but not equipment number.**
- **Does not contain equipment number or serial number.**
- **Error: Could not find serial number in the Excel file.**

## Script Breakdown

1. `extract_text_from_pdf(pdf_path)`: Extracts all text from a PDF file.
2. `check_pdf_for_numbers(pdf_path, equipment_number, serial_number)`: Checks if the equipment number and serial number are present in the PDF text.
3. `find_value_next_to_label(sheet, label): Retrieves the value next to a specified label in an Excel sheet.`
4. **main()**: Main function that orchestrates the processing of folders, files, and logging.

## Example

Given the following directory structure:

```bash
/home/freguez/Documents/equipamentos/
    1234/
        Ficha1234.xlsx
        calibração/
            document1.pdf
            document2.pdf
    5678/
        Ficha5678.xlsx
        calibração/
            document3.pdf
```

Running the script will check each PDF in the 'calibração' subfolder for the equipment number and serial number extracted from the corresponding Excel file. Results will be logged in the specified output files based on their content.

## Troubleshooting

- Ensure paths in the script are correct.
- Verify that your directory structure matches the expected format.
- Check permissions for reading files and writing output.

## License

This script is provided under the MIT License. See the LICENSE file for more details.

## Contact

For questions or issues, please contact [Nishchiguti] at [nischiguti@pm.me].
