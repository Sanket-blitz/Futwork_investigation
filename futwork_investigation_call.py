import os
import re
import requests
import pandas as pd
import torch
import whisper
from datetime import datetime
from urllib.parse import urlparse
from tqdm import tqdm

# Load model
print("🔍 Loading Whisper model (base)...")
model = whisper.load_model("base")

# Ask user for Excel path
excel_path = input("Enter the path of the Excel file to detect the futwork<>Cx call: ").strip()
if not os.path.isfile(excel_path):
    raise FileNotFoundError(f"❌ File not found: {excel_path}")

# Read Excel
df = pd.read_excel(excel_path)
recording_col = next((col for col in df.columns if re.search(r"recording", col, re.IGNORECASE)), None)
if not recording_col:
    raise Exception("❌ 'Recording Link' column not found in the Excel file.")

# Prepare output folder
output_dir = os.path.splitext(excel_path)[0]
os.makedirs(output_dir, exist_ok=True)

# Generate unique output filename
base_filename = os.path.join(output_dir, f"futwork_investigation_call_{datetime.today().date()}")
output_path = base_filename + ".xlsx"
counter = 1
while os.path.exists(output_path):
    output_path = f"{base_filename}-{counter}.xlsx"
    counter += 1

# Function to extract intent
def get_intent(text):
    lowered = text.lower()
    if "cancel" in lowered or "not available" in lowered:
        return "Cancellation"
    elif "delivery" in lowered or "receive" in lowered:
        return "Delivery related"
    elif "call back" in lowered or "busy" in lowered:
        return "Callback requested"
    elif "wrong number" in lowered:
        return "Wrong Number"
    return "Uncategorized"

# Process each recording
transcripts = []
for idx, row in enumerate(tqdm(df.itertuples(), total=len(df), desc="📞 Processing calls")):
    url = getattr(row, recording_col)
    if not isinstance(url, str) or not url.startswith("http"):
        print(f"⚠️ Skipping row {idx+1} — Invalid link.")
        transcripts.append(("", "Invalid or missing link"))
        continue

    try:
        filename = os.path.join(output_dir, f"call_{idx+1}.mp3")
        print(f"\n➡️ [{idx+1}] Downloading: {url}")
        with requests.get(url, stream=True, timeout=30) as r:
            with open(filename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)

        print(f"🎙️ [{idx+1}] Transcribing...")
        result = model.transcribe(filename)
        text = result["text"].strip()

        intent = get_intent(text)
        transcripts.append((text, intent))

        print(f"✅ [{idx+1}] Done — Intent: {intent}")
        os.remove(filename)

    except Exception as e:
        print(f"❌ [{idx+1}] Error: {e}")
        transcripts.append(("", f"Error: {e}"))

# Add results to DataFrame
df["Transcription"] = [t[0] for t in transcripts]
df["Intent"] = [t[1] for t in transcripts]

# Save results
df.to_excel(output_path, index=False)
print(f"\n✅ All done! Results saved to:\n{output_path}")
