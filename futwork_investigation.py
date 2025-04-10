import os
import pandas as pd
import json
import time
from datetime import datetime

def create_loading_screen():
    """Display a graphical loading screen in the terminal."""
    for i in range(1, 101):
        time.sleep(0.05)  # Simulate processing time
        progress_bar = "#" * (i // 2) + "-" * (50 - i // 2)
        print(f"\rLoading... Please Wait [{i}%] [{progress_bar}]", end="")
    print("\nProcessing complete!")

def process_data(file_path):
    """Filter rows with recording links and save them to a new file."""
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

    # Filter rows where 'Recording' is not null
    filtered_df = df[df['Recording'].notna()]

    # Convert 'Response Time' column to datetime for hourly grouping
    filtered_df['Response Time'] = pd.to_datetime(filtered_df['Response Time'])

    # Group by hour and select up to 5 random rows per hour
    grouped_df = (
        filtered_df.groupby(filtered_df['Response Time'].dt.hour)
        .apply(lambda x: x.sample(n=min(5, len(x)), random_state=42))
        .reset_index(drop=True)
    )

    # Create output folder and file names dynamically
    current_date = datetime.now().strftime("%Y-%m-%d")
    base_folder_name = f"futwork_investigation_{current_date}"
    folder_suffix = 1

    while os.path.exists(f"{base_folder_name}-{folder_suffix}"):
        folder_suffix += 1

    output_folder = f"{base_folder_name}-{folder_suffix}"
    os.makedirs(output_folder, exist_ok=True)

    output_file_name = f"futwork_investigation_data_{current_date}.xlsx"
    output_file_path = os.path.join(output_folder, output_file_name)

    # Save the filtered data to a new Excel file
    grouped_df.to_excel(output_file_path, index=False)

    print(f"File saved at: {output_file_path}")

if __name__ == "__main__":
    # Input file path from user
    input_file_path = input("Enter the path of the Excel file: ")

    # Create a loading screen while processing
    create_loading_screen()

    # Process the data and save the results
    process_data(input_file_path)
