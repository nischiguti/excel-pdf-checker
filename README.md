Equipment PDF and Excel Processing Script
Overview
This script processes folders containing equipment data by extracting and categorizing information from Excel and PDF files. It performs the following tasks:

Extracts serial numbers from Excel files.
Checks if PDFs contain specific equipment numbers and serial numbers.
Logs results into categorized text files based on the findings.
Features
PDF Text Extraction: Extracts text from PDF files to check for the presence of equipment and serial numbers.
Excel Data Extraction: Retrieves serial numbers from Excel files based on a specific label.
Categorized Logging: Saves results into different text files based on the presence of equipment and serial numbers.
Prerequisites
Ensure you have the following Python libraries installed:

openpyxl (for reading Excel files)
PyPDF2 (for reading PDF files)
You can install these libraries using pip if you don't have them already:

bash
Copy code
pip install openpyxl PyPDF2
Setup
Directory Structure: Organize your directories as follows:

A root directory containing subfolders named by equipment numbers.
Each equipment folder should contain an Excel file with 'Ficha' in its name and a subfolder named 'calibração' containing PDF files.
Script Configuration: Update the script to reflect your actual directory paths.

root_dir: Path to the root directory containing all equipment folders.

python
Copy code
root_dir = '/path/to/your/root_directory'
output_files: Dictionary mapping result categories to output file paths.

python
Copy code
output_files = {
    "Contains both equipment number and serial number.": '/path/to/contains_both.txt',
    "Contains equipment number but not serial number.": '/path/to/contains_equipment_only.txt',
    "Contains serial number but not equipment number.": '/path/to/contains_serial_only.txt',
    "Does not contain equipment number or serial number.": '/path/to/contains_none.txt',
    "Error: Could not find serial number in the Excel file.": '/path/to/errors.txt'
}
Usage
Configure the Script: Edit the script to set the correct paths for root_dir and output_files.

Run the Script: Execute the script using Python. For example:

bash
Copy code
python your_script_name.py
Review Results: After the script completes, check the specified output files for categorized results:

Contains both equipment number and serial number.
Contains equipment number but not serial number.
Contains serial number but not equipment number.
Does not contain equipment number or serial number.
Error: Could not find serial number in the Excel file.
Script Breakdown
extract_text_from_pdf(pdf_path): Extracts all text from a PDF file.
check_pdf_for_numbers(pdf_path, equipment_number, serial_number): Checks if the equipment number and serial number are present in the PDF text.
find_value_next_to_label(sheet, label): Retrieves the value next to a specified label in an Excel sheet.
main(): Main function that orchestrates the processing of folders, files, and logging.
Example
Given the following directory structure:

bash
Copy code
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
Running the script will check each PDF in the 'calibração' subfolder for the equipment number and serial number extracted from the corresponding Excel file. Results will be logged in the specified output files based on their content.

Troubleshooting
Ensure paths in the script are correct.
Verify that your directory structure matches the expected format.
Check permissions for reading files and writing output.
License
This script is provided under the MIT License. See the LICENSE file for more details.

Contact
For questions or issues, please contact [Your Name] at [your.email@example.com].

Key Points for Copying Commands
Code Blocks: Use triple backticks (```) to create a block of code that is easy to copy.
Inline Code: Use single backticks (`) for inline code snippets to highlight commands or file paths within text.
Copy-Friendly Paths: Provide example paths in the script configuration sections that users can easily copy and paste into their scripts.
Feel free to modify paths, email addresses, or any specific details to fit your actual requirements!
