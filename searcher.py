import os
import logging
from openpyxl import load_workbook

# Suppress openpyxl warnings
logging.getLogger('openpyxl').setLevel(logging.ERROR)

def find_string_in_workbook(directory, search_string):
    found_info = []
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            workbook = load_workbook(filename=file_path)
            for sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
                for row in worksheet.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and search_string.lower() in str(cell.value).lower():
                            found_info.append((filename, sheet, cell.coordinate))
    return found_info

# Use dot `.` to represent the current directory
directory = '.'
# Replace 'search_string' with the string you want to search for
search_string = 'polder'

found_cells = find_string_in_workbook(directory, search_string)

if found_cells:
    print(f"The search string '{search_string}' was found in the following locations:")
    for cell_info in found_cells:
        print(f"In workbook '{cell_info[0]}' in sheet '{cell_info[1]}' at cell '{cell_info[2]}'")
else:
    print(f"The search string '{search_string}' was not found in any of the files.")
