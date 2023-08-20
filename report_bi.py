'''
Prepare the data from xlsx files for Power Bi import:
- check all .xlsx files in selected path
- collect the data into a DataFrame
- drop some unnecessary lines; replace some text
- combine the data into a single DataFrame and export into .csv
'''

from pathlib import Path
from datetime import datetime

import easygui
import pandas as pd


def read_data(file_path) -> pd.DataFrame:
    try:
        # Read data from Excel sheet (Work List) with header at row 3
        data = pd.read_excel(file_path, sheet_name='Work List', header=2)

        # Drop rows where Planned and SR are empty
        data = data.dropna(subset=[data.columns[14]])
        data = data.dropna(subset=[data.columns[3]])

        # Rename columns'
        if 'WR Number' in data.columns:
            data.rename(columns={'WR Number': 'NCR_SR\nNo.'}, inplace=True)
        if 'WO Number' in data.columns:
            data.rename(columns={'WO Number': 'NCR_TASK\nNo.'}, inplace=True)
        if 'PIM Name' in data.columns:
            data['PIM Name'] = 'AT&T'
            data.rename(columns={'PIM Name': 'Customer'}, inplace=True)

        return data

    except PermissionError:
        # If file is open, prompt user to close it
        file_name = Path(file_path).name
        message = f'The file {file_name} might be open. \nPlease close it and press OK to continue processing.'
        keep_going = easygui.ynbox(message, 'File Open', ('OK', 'Cancel'))
        if not keep_going:  # User clicked 'Cancel'
            raise KeyboardInterrupt  # Stop the execution
    
    except ValueError as ve:
        if 'Worksheet named "Work List" not found' in str(ve):
            file_name = Path(file_path).name
            message = f'The file {file_name} is missing a worksheet named "Work List". \nPlease make sure the worksheet name is correct.'
            keep_going = easygui.ynbox(message, 'File Open', ('OK', 'Cancel'))
            if not keep_going:  # User clicked 'Cancel'
                raise KeyboardInterrupt  # Stop the execution
        else:
            raise # Re-raise the exception for other ValueError cases


def process_folder(folder):
    all_data = pd.DataFrame()
    
    for file_path in folder.glob('**/*.xlsx'):
        print('\nProcessing:', file_path.name)
        try:
            data = read_data(str(file_path))
            if data is not None:
                all_data = pd.concat([all_data, data], ignore_index=True)
        except KeyboardInterrupt:
            print('Execution stopped by user.')
            return None
    
    return all_data


# Initialize the DataFrame to accumulate all processed data
combined_data = pd.DataFrame()

# Main loop for multiple processing sessions
while True:
    selected_directory = easygui.diropenbox(title='Select the root directory')

    if selected_directory:
        # Process all files in all included folders using pathlib
        processed_data = process_folder(Path(selected_directory))
        if processed_data.empty:
            print('No valid data found.')
        else:
            combined_data = pd.concat([combined_data, processed_data], ignore_index=True)
    else:
        print('No directory selected.')
    
    keep_going = easygui.ynbox('Do you want to process more data?', title='Continue?')
    if not keep_going:
        break

if not combined_data.empty:
    easygui.msgbox('Choose the location to save the exported data.')
    save_directory = easygui.diropenbox(title='Select the root directory')

    # Export the combined processed data to a CSV file
    current_date = datetime.now().strftime('%Y%m%d')
    output_csv_path = save_directory + f'\combined_data_{current_date}.csv'
    combined_data.to_csv(output_csv_path, index=False, encoding='utf_8_sig')
    print('\nAll data exported to CSV:', output_csv_path)
else:
    print('No valid data to export.')