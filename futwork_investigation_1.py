import os
import json
import pandas as pd
import random
import time
from datetime import datetime

def extract_proof_link(response):
    """Extract proofAttachmentLink from JSON response."""
    try:
        data = json.loads(response)
        return data.get("ndr_update", {}).get("proofAttachmentLink", "")
    except (json.JSONDecodeError, TypeError):
        return ""

def show_loading_screen():
    """Displays a loading screen from 1% to 100%."""
    for i in range(1, 101):
        time.sleep(0.05)  # Adjust speed if needed
        print(f"\rProcessing... {i}%", end="", flush=True)
    print("\nProcessing complete!")

def detect_timestamp_column(df):
    """Automatically detects the timestamp column."""
    possible_columns = ["Timestamp", "created_at", "Date", "time"]
    for col in possible_columns:
        if col in df.columns:
            return col
    return None

def process_futwork_data(file_path):
    """Processes the given Excel file and extracts valid cases."""
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Extract proofAttachmentLink
    df["proofAttachmentLink"] = df["Response"].apply(extract_proof_link)
    
    # Filter cases where proofAttachmentLink contains a valid recording link
    filtered_df = df[df["proofAttachmentLink"].str.startswith("https://recordings.futwork.com")].copy()
    
    if filtered_df.empty:
        print("No valid cases found with recording links.")
        return
    
    # Add a new column 'recording' with all links
    filtered_df.loc[:, "recording"] = filtered_df["proofAttachmentLink"]
    
    # Detect timestamp column
    timestamp_col = detect_timestamp_column(df)
    if not timestamp_col:
        print("Error: No valid timestamp column found.")
        return
    
    # Convert timestamp column to datetime
    filtered_df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors='coerce')
    filtered_df.dropna(subset=[timestamp_col], inplace=True)
    filtered_df["hour"] = filtered_df[timestamp_col].dt.hour
    
    # Group by hour and select 5 random cases per hour
    sampled_df = filtered_df.groupby("hour").apply(lambda x: x.sample(n=min(5, len(x)), random_state=random.randint(1, 1000))).reset_index(drop=True)
    
    # Create a unique folder
    current_date = datetime.now().strftime("%Y-%m-%d")
    base_folder = f"futwork_investigation_{current_date}"
    folder_number = 1
    while os.path.exists(f"{base_folder}-{folder_number}"):
        folder_number += 1
    output_folder = f"{base_folder}-{folder_number}"
    os.makedirs(output_folder, exist_ok=True)
    
    # Define output filename
    output_file = os.path.join(output_folder, f"futwork_investigation_data_{current_date}.xlsx")
    
    # Save the filtered data to a new Excel file
    sampled_df.to_excel(output_file, index=False)
    print(f"Data successfully saved to {output_file}")

if __name__ == "__main__":
    input_file = input("Enter the path of the .xlsx file: ")
    show_loading_screen()
    process_futwork_data(input_file)
