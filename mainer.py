import pandas as pd
import os

# Define the folder containing the Excel files and the output file
folder_path = "C:/Users/SAUDI/Documents/PROJECTS/SOUNDEO/source_mainer"
output_file = "consolidated_data_with_tabs.xlsx"

files = [
    ("1_consolidated_data_match_stats.xlsx", "match_stats"),
    ("2_consolidated_data_bat_stat.xlsx", "bat_stat"),
    ("3_consolidated_data_innings_wkfall.xlsx", "innings_wkfall"),
    ("4_consolidated_data_bowl_stat.xlsx", "bowl_stat"),
    ("5_consolidated_data_player_info.xlsx", "player_info"),
    ("6_consolidated_data_mvp.xlsx", "mvp"),
    ("7_consolidated_data_commentary.xlsx", "commentary"),
    ("8_consolidated_data_overs.xlsx", "overs")
]

# Create a Pandas Excel writer to write data to the new file
with pd.ExcelWriter(output_file) as writer:
    # Loop through each file and corresponding sheet
    for file_name, sheet_name in files:
        file_path = os.path.join(folder_path, file_name)  # Construct the full path to the file
        
        # Try reading the sheet and writing it to the new Excel file
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Successfully added {sheet_name} from {file_name}")
        except Exception as e:
            print(f"Error reading {file_name} - {sheet_name}: {e}")

print(f"Data has been successfully consolidated into {output_file}.")