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

def select_directory() -> str:
    '''Prompt user to select a directory using easygui.'''
    selected_directory = easygui.diropenbox(title='Select the root directory')
    return selected_directory or ''  # Return an empty string if None is returned


def process_data_from_directory(directory: str, combined_data: pd.DataFrame) -> pd.DataFrame:
    '''Process all files in the provided directory and return combined DataFrame.'''
    if directory:
        processed_data = process_folder(Path(directory))
        if processed_data is not None:
            if not processed_data.empty:
                combined_data = pd.concat([combined_data, processed_data], ignore_index=True)
            else:
                print('No valid data found.')
        else:
            print('Error occurred while processing data.')
    else:
        print('No directory selected.')
    return combined_data


def process_folder(folder) -> pd.DataFrame | None:
    all_data = []
    for file_path in folder.glob('**/*.xlsx'):
        print('\nProcessing:', file_path.name)
        try:
            data = read_file(str(file_path))
            if data is not None:
                all_data.append(data)
        except KeyboardInterrupt:
            print('Execution stopped by user.')
            return None
    
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()


def read_file(file_path) -> pd.DataFrame | None:
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
    
    return None


def main():
    combined_data = pd.DataFrame()

    while True:
        selected_directory = select_directory()
        combined_data = process_data_from_directory(selected_directory, combined_data)

        keep_going = easygui.ynbox('Do you want to process more data?', title='Continue?')
        if not keep_going:
            break

    if not combined_data.empty:
        # Export the combined processed data to a CSV file
        easygui.msgbox('Choose the location to save the exported data.')
        save_directory = easygui.diropenbox(title='Select the root directory')

        if save_directory:
            current_date = datetime.now().strftime('%Y%m%d')
            output_csv_path = Path(save_directory) / f'combined_data_{current_date}.csv'
            combined_data.to_csv(output_csv_path, index=False, encoding='utf_8_sig')
            print('\nAll data exported to CSV:', output_csv_path)
        else:
            print('No valid directory selected.')
    else:
        print('No valid data to export.')


if __name__ == "__main__":
    main()