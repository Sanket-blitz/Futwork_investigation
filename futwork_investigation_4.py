import os
import pandas as pd
import json
import time
from datetime import datetime
import warnings
from pandas.errors import SettingWithCopyWarning

def create_loading_screen():
    """Display a graphical loading screen in the terminal."""
    for i in range(1, 101):
        time.sleep(0.02)  # Slightly faster loading screen
        progress_bar = "#" * (i // 2) + "-" * (50 - i // 2)
        print(f"\rLoading... Please Wait [{i}%] [{progress_bar}]", end="")
    print("\nProcessing complete!")

def process_data(file_path):
    """Filter all rows with valid recording links and save to a new Excel file."""
    
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Extract recording links from the 'Response' column
    def extract_recording_link(response):
        try:
            response_data = json.loads(response)
            proof_link = response_data.get("ndr_update", {}).get("proofAttachmentLink", "")
            return proof_link if proof_link.startswith("https://recordings.futwork.com") else None
        except (json.JSONDecodeError, TypeError):
            return None

    # Create a new column 'Recording' with valid recording links
    df['Recording'] = df['Response'].apply(extract_recording_link)

    # Filter only rows where 'Recording' is not null
    filtered_df = df[df['Recording'].notna()].copy()

    # Create output folder and file name
    current_date = datetime.now().strftime("%Y-%m-%d")
    base_folder_name = f"futwork_investigation_{current_date}"
    folder_suffix = 1

    while os.path.exists(f"{base_folder_name}-{folder_suffix}"):
        folder_suffix += 1

    output_folder = f"{base_folder_name}-{folder_suffix}"
    os.makedirs(output_folder, exist_ok=True)

    output_file_name = f"futwork_investigation_all_recordings_{current_date}.xlsx"
    output_file_path = os.path.join(output_folder, output_file_name)

    # Save the filtered data
    filtered_df.to_excel(output_file_path, index=False)
    print(f"\n✅ File saved at: {output_file_path}")

if __name__ == "__main__":
    input_file_path = input("Enter the path of the Excel file: ")
    create_loading_screen()
    process_data(input_file_path)
