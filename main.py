import pandas as pd
import os

# Define the folder containing the Excel files and the output file
folder_path = "C:/Users/SAUDI/Documents/PROJECTS/SOUNDEO/source"


# Name of the sheet to consolidate
sheets = [
    "match_stats",
    "bat_stat",
    "innings_wkfall",
    "bowl_stat",
    "player_info",
    "mvp",
    "commentary",
    "overs"
]

for target_sheet in sheets:
    # Initialize an empty DataFrame to hold all the data for the target sheet
    consolidated_data = pd.DataFrame()

    # Loop through all Excel files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") or file_name.endswith(".xls"):  # Ensure only Excel files are processed
            file_path = os.path.join(folder_path, file_name)
            # Read the specific sheet
            try:
                sheet_data = pd.read_excel(file_path, sheet_name=target_sheet)
                # Append the data to the consolidated DataFrame
                consolidated_data = pd.concat([consolidated_data, sheet_data], ignore_index=True)
            except Exception as e:
                print(f"Skipping {file_name} due to error: {e}")

    # Write the consolidated data to a new Excel file
    output_file = f"{sheets.index(target_sheet)+1}_consolidated_data_{target_sheet}.xlsx"
    consolidated_data.to_excel(output_file, sheet_name=target_sheet, index=False)

    print(f"Data from the '{target_sheet}' sheet in all files has been consolidated into {output_file}.")
