import configparser
import pandas as pd
import json
import numpy as np
import os
import sys

config = configparser.ConfigParser()
config.read('config.properties')

excel_file_path = config['Import']['excel_file_path']
json_file_path = config['Import']['json_file_path']
firestore_config_file = config['firestore']['config_file_path']

def is_json_like(value):
    """
    Checks if a string resembles a JSON object.
    """
    if isinstance(value, str):
        value = value.strip()
        return (value.startswith('{') and value.endswith('}')) or (value.startswith('[') and value.endswith(']'))
    return False

def convert_to_json(value):
    """
    Converts a string representation of a JSON object to a dictionary.
    """
    if is_json_like(value):
        try:
            return json.loads(value.replace("'", "\""))
        except json.JSONDecodeError:
            pass
    return value

def convert_json_fields(data):
    """
    Converts fields containing JSON-like strings into JSON objects.
    """
    for key, value in data.items():
        if isinstance(value, str):
            json_data = convert_to_json(value)
            if json_data is not None:
                data[key] = json_data
        elif isinstance(value, float) and np.isnan(value):
            data[key] = ""
    return data

def convert_excel_to_json(excel_file_path):
    # Reading the Excel file
    excel_data = pd.read_excel(excel_file_path, sheet_name=None)

    # Initializing the JSON data dictionary
    json_data = {"__collections__": {}}

    # Iterating over sheets and convert each one to a collection in the JSON data
    for collection_name, df in excel_data.items():
        collection_data = {}
        for index, row in df.iterrows():
            item_data = row.to_dict()
            if "Unnamed: 0" in item_data:
                item_id = item_data["Unnamed: 0"]
                del item_data["Unnamed: 0"]
                collection_data[item_id] = convert_json_fields(item_data)
            else:
                collection_data[index] = convert_json_fields(item_data)

        json_data["__collections__"][collection_name] = collection_data

    return json_data

def main():
    # Checking if backup file is passed as an argument
    if len(sys.argv) > 1:
        backup_file = sys.argv[1]
        print(f"Using backup file: {backup_file}")
        json_file_path_to_use = backup_file
    else:
        # Converting Excel to JSON if no backup file is provided
        json_data = convert_excel_to_json(excel_file_path)

        with open(json_file_path, 'w') as json_file:
            json.dump(json_data, json_file, indent=4)

        print(f'JSON file saved successfully at: {json_file_path}')
        json_file_path_to_use = json_file_path

    command = [
        "npx", "-p", "node-firestore-import-export", "firestore-import",
        "-a", firestore_config_file,
        "-b", json_file_path_to_use
    ]

    os.system(" ".join(command))

if __name__ == "__main__":
    main()
