'''
This script checks sheet names for all .xlsx files in selected folder.
If the sheet name contains Work List with additional characters,
the script will rename the sheet to unify the data.

WARNING!
This script wont recalculate the formulas so it may cause some problems 
when reading the files with pandas.
'''

import os
import openpyxl
import tkinter as tk
from tkinter import filedialog

def update_sheet_name(workbook, old_name, new_name): 
    workbook[old_name].title = new_name
    workbook.save(file_path)
    print(f'Sheet name "{old_name}" has been changed to "{new_name}"')


def process_excel_file(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        
        for sheet_name in workbook.sheetnames:
            if sheet_name == target_phrase:
                print(f'Sheet name "{sheet_name}" already matches the target phrase.')
                workbook.close()
                return  # No need to update, so skip the rest of the loop
            
            if target_phrase in sheet_name:
                update_sheet_name(workbook, sheet_name, target_phrase)
                break
        
        workbook.close()
    except Exception as e:
        print(f"An error occurred while processing '{file_path}': {e}")


# Define the target phrase
target_phrase = 'Work List'

# Create a tkinter root window (this won't be displayed)
root = tk.Tk()
root.withdraw()  # Hide the root window

# Prompt the user to select a directory using a file dialog
selected_directory = filedialog.askdirectory(title="Select the root directory")

# Traverse through all files in the directory and its subdirectories
try:
    if not selected_directory:
        print("No directory selected. Exiting...")
    else:
        for folder_path, _, file_names in os.walk(selected_directory):
            for file_name in file_names:
                if file_name.endswith('.xlsx'):
                    file_path = os.path.join(folder_path, file_name)
                    print(f'\nProcessing: {file_name}')
                    process_excel_file(file_path)
except KeyboardInterrupt:
    print("Execution stopped by user.")
