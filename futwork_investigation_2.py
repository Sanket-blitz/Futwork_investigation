import os
import pandas as pd
import json
import time
from datetime import datetime
import warnings
from pandas.errors import SettingWithCopyWarning

warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)

def create_loading_screen():
    """Display a graphical loading screen in the terminal."""
    for i in range(1, 101):
        time.sleep(0.05)
        progress_bar = "#" * (i // 2) + "-" * (50 - i // 2)
        print(f"\rLoading... Please Wait [{i}%] [{progress_bar}]", end="")
    print("\nProcessing complete!")

def process_data(file_path):
    """Filter rows with recording links and save them to a new file."""
    df = pd.read_excel(file_path)

    # Extract recording link from JSON in 'response' column
    def extract_recording_link(response):
        try:
            response_data = json.loads(response)
            proof_link = response_data.get("ndr_update", {}).get("proofAttachmentLink", "")
            return proof_link if proof_link.startswith("https://recordings.futwork.com") else None
        except (json.JSONDecodeError, TypeError):
            return None

    # Add new column for valid links
    df["Recording"] = df["response"].apply(extract_recording_link)

    # Filter rows with valid links
    filtered_df = df[df["Recording"].notna()].copy()

    # Convert 'response_time' to datetime for grouping
    filtered_df.loc[:, "response_time"] = pd.to_datetime(filtered_df["response_time"])

    # Group by hour and sample up to 5 rows
    def sample_hour(hour_group):
        return hour_group.sample(min(5, len(hour_group)), random_state=42)

    hourly_samples = filtered_df.groupby(filtered_df["response_time"].dt.hour, group_keys=False).apply(sample_hour)

    # Dynamic folder & file naming
    current_date = datetime.now().strftime("%Y-%m-%d")
    base_folder = f"futwork_investigation_{current_date}"
    folder_suffix = 1
    while os.path.exists(f"{base_folder}-{folder_suffix}"):
        folder_suffix += 1
    output_folder = f"{base_folder}-{folder_suffix}"
    os.makedirs(output_folder, exist_ok=True)

    output_file = f"futwork_investigation_data_{current_date}.xlsx"
    output_path = os.path.join(output_folder, output_file)

    # Save final output
    hourly_samples.to_excel(output_path, index=False)
    print(f"File saved at: {output_path}")

if __name__ == "__main__":
    input_file_path = input("Enter the path of the Excel file: ")
    create_loading_screen()
    process_data(input_file_path)
