import argparse
import json
from datetime import datetime
from pathlib import Path
import pandas as pd
import os
import configparser
import subprocess

# Read the configuration
config = configparser.ConfigParser()
config.read('config.properties')

# Get absolute paths from config
json_file_path = os.path.abspath(config['Export']['json_file_path'])
excel_file_path = os.path.abspath(config['Export']['excel_file_path'])
backup_file_path = os.path.abspath(config['Export']['backup_file_path'])
firestore_config_file = os.path.abspath(config['firestore']['config_file_path'])

print("JSON Path:", json_file_path)
print("Firestore Credential Path:", firestore_config_file)

def convert_json_to_excel(json_data, excel_file_path):
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as excel_writer:
        collections = json_data.get("__collections__", {})
        if not collections:
            print("Error: No collections found in the JSON data.")
            return

        for collection_name, collection_data in collections.items():
            if collection_data:
                df = pd.DataFrame.from_dict(collection_data, orient='index')
                df.to_excel(excel_writer, sheet_name=collection_name)

        if not excel_writer.sheets:
            raise ValueError("No sheets created in the Excel file. The JSON data might be empty.")

def main(backup):
    file_path = json_file_path
    if backup:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = f"{backup_file_path}_{timestamp}.json"  # ensure .json extension

    # Construct command to export Firestore data
    command = [
        "npx", "-p", "node-firestore-import-export", "firestore-export",
        "-a", firestore_config_file,
        "-b", file_path
    ]

    # Wrap arguments with spaces in double quotes
    command = [f'"{arg}"' if ' ' in arg else arg for arg in command]

    # Run the command
    try:
        subprocess.run(" ".join(command), shell=True, check=True)
    except subprocess.CalledProcessError as e:
        print("Error: Failed to export Firestore data.")
        print(e)
        return

    if not backup:
        try:
            with open(json_file_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)

            convert_json_to_excel(json_data, excel_file_path)
            print(f'✅ Excel file saved successfully at: {excel_file_path}')
        except FileNotFoundError:
            print("❌ Error: JSON file not found.")
        except json.JSONDecodeError:
            print("❌ Error: Failed to decode JSON file.")
        except ValueError as e:
            print(f"❌ Error: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Export Firestore data to Excel.')
    parser.add_argument('--backup', action='store_true', help='Create a backup of the JSON file before processing.')
    args = parser.parse_args()
    main(args.backup)

