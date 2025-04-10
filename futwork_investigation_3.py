import os
import pandas as pd
import json
import time
from datetime import datetime

def create_loading_screen():
    """Display a graphical loading screen in the terminal."""
    for i in range(1, 101):
        time.sleep(0.01)  # You can adjust this for speed
        progress_bar = "#" * (i // 2) + "-" * (50 - i // 2)
        print(f"\rLoading... Please Wait [{i}%] [{progress_bar}]", end="")
    print("\nProcessing complete!")

def process_data(file_path):
    """Filter rows with duplicate AWBs (Trip IDs) and recording links."""

    # Load Excel
    df = pd.read_excel(file_path)

    # Print available columns for debugging
    print("\n📋 Columns found in the file:")
    for col in df.columns:
        print(f"- '{col}'")

    # Use 'Trip ID' as AWB
    awb_column = "Trip ID"
    if awb_column not in df.columns:
        raise ValueError(f"❌ Column '{awb_column}' not found in the Excel file. Please check column names above.")

    # Extract recording link
    def extract_recording_link(response):
        try:
            data = json.loads(response)
            return data.get("ndr_update", {}).get("proofAttachmentLink", "")
        except (json.JSONDecodeError, TypeError):
            return None

    df['Recording'] = df['Response'].apply(extract_recording_link)
    df = df[df['Recording'].str.startswith("https://recordings.futwork.com", na=False)]

    # Filter AWBs that occur more than once
    awb_counts = df[awb_column].value_counts()
    repeated_awbs = awb_counts[awb_counts > 1].index
    result_df = df[df[awb_column].isin(repeated_awbs)].copy()

    # Create output folder
    current_date = datetime.now().strftime("%Y-%m-%d")
    base_folder = f"futwork_investigation_{current_date}"
    suffix = 1
    while os.path.exists(f"{base_folder}-{suffix}"):
        suffix += 1
    final_folder = f"{base_folder}-{suffix}"
    os.makedirs(final_folder, exist_ok=True)

    # Save result
    output_path = os.path.join(final_folder, f"futwork_investigation_data_{current_date}.xlsx")
    result_df.to_excel(output_path, index=False)
    print(f"\n✅ File saved at: {output_path}")

if __name__ == "__main__":
    input_path = input("Enter the path of the Excel file: ")
    create_loading_screen()
    process_data(input_path)
