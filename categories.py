def filter_by_category(input_file, output_files):
    """
    Filters blocks in the input file into categories and writes them to corresponding output files.
    Each block's category is determined by its last line.
    """
    # Open output files
    output_handlers = {category: open(file, 'w') for category, file in output_files.items()}
    
    with open(input_file, 'r') as infile:
        # Initialize variables
        current_block = []
        
        for line in infile:
            if line.strip() == "":
                # End of current block
                if current_block:
                    # Determine the category of the current block
                    block_text = "".join(current_block).strip()
                    last_line = block_text.splitlines()[-1].strip()  # Use splitlines and strip
                    
                    # Debug: Print the last line to verify category matching
                    print(f"Processing block, last line: '{last_line}'")
                    
                    # Find the right category for the block
                    found = False
                    for category in output_files.keys():
                        if category in last_line:
                            print(f"Match found for category: '{category}'")  # Debugging output
                            output_handlers[category].write(block_text + "\n\n")
                            found = True
                            break
                    
                    # If no match, categorize as an error
                    if not found:
                        output_handlers["Error: Could not find equipment number or serial number in the Excel file."].write(block_text + "\n\n")
                
                # Reset for the next block
                current_block = []
            else:
                # Collect lines for the current block
                current_block.append(line)
        
        # Handle the last block if the file doesn't end with a blank line
        if current_block:
            block_text = "".join(current_block).strip()
            last_line = block_text.splitlines()[-1].strip()
            print(f"Processing last block, last line: '{last_line}'")  # Debugging output
            
            found = False
            for category in output_files.keys():
                if category in last_line:
                    print(f"Match found for category: '{category}'")  # Debugging output
                    output_handlers[category].write(block_text + "\n\n")
                    found = True
                    break
            
            if not found:
                output_handlers["Error: Could not find equipment number or serial number in the Excel file."].write(block_text + "\n\n")
    
    # Close all output files
    for handler in output_handlers.values():
        handler.close()

# Define file paths
input_file = '/home/freguez/Documents/equipamentos/filter_filter.txt'  # The file containing the results

# Define output files for each category
output_files = {
    "Contains both equipment number and serial number.": '/home/freguez/Documents/equipamentos/contains_both.txt',
    "Contains equipment number but not serial number.": '/home/freguez/Documents/equipamentos/contains_equipment_only.txt',
    "Contains serial number but not equipment number.": '/home/freguez/Documents/equipamentos/contains_serial_only.txt',
    "Does not contain equipment number or serial number.": '/home/freguez/Documents/equipamentos/contains_none.txt',
    "Error: Could not find equipment number or serial number in the Excel file.": '/home/freguez/Documents/equipamentos/errors.txt'
}

# Run the filtering
filter_by_category(input_file, output_files)

print("Filtering complete. Results have been saved to the respective files.")

