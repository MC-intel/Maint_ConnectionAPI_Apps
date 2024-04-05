# This script sends a POST request to generate a new Asset in our database. The script allows the user to input custom parameters and streamlines making requests:   

import requests
import json
import time

# Title
title = " ASSET GENERATOR "
delay = 0.1
for char in title:
    print(char, end='', flush=True)
    time.sleep(delay)
print('\n\nSelect the location of the Asset...\n')

# Function to get input for a dictionary key
def get_input_for_key(key, display_name=None):
    if display_name is None:
        display_name = key
    return input(f"Enter the value for '{display_name}': ")

# Load ParentRefs data from file
with open(r" ...FILEPATH... \assetParents.txt", "r") as file:
    parent_refs_data = json.load(file)
parent_refs = parent_refs_data["ParentRefs"]

# Display ParentRefs and get user selection
for i, ref in enumerate(parent_refs):
    time.sleep(.2)
    print(f"{i + 1}. {ref['Name']}\n")

# Ensuring valid selection
while True:
    print('\n')
    try:
        selected_index = int(input("  Select the Lab by list number: ")) - 1
        if 0 <= selected_index < len(parent_refs):
            break
        else:
            print("\nInvalid selection. Please choose Lab by list number (example: 10).")
    except ValueError:
        print("\nInvalid input. Please choose Lab by list number (example: 10).")

# Selected ParentRef
selected_parent_ref = parent_refs[selected_index]

print('\nNow input the required values for the new Asset\n')

# Constructing the URL
url = f'http:// ...URL... /assets?$filter=ParentRef.id%20eq%20"{selected_parent_ref["ID"]}"'

# Headers
headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ...APIkey... ',
    'Connection': 'keep-alive'
}

# Payload template
single_payload = {
    "ClassificationRef": {"PK": 34, "ID": "EQUIPMENT", "Name": "Equipment"},
    "ParentRef": selected_parent_ref,
    "DepartmentRef": {"PK": 50, "ID": "Transformation", "Name": "Transformation"},
    "RepairCenterRef": {"PK": 45, "ID": "CV-GG", "Name": "Chesterfield GG Labs"},
    "UDFChar10": "QUALITY CRITICAL",
}

# Mapping of user-friendly input names to API keys
key_mapping = {
    "SAP# (ex. 300...) ": "UDFChar3",
    "Metrology# (ex. M...) ": "UDFChar4",
    "Next Calibration Date (ex. 1/31/1999) ": "UDFDate5",
    "Last Calibration Date (ex. 1/31/1999) ": "UDFDate3",
}

# Keys for which user input is required
input_keys = ["ID", "Name", "Model"] + list(key_mapping.keys())

# Get user input for specified keys
for key in input_keys:
    api_key = key_mapping.get(key, key)
    single_payload[api_key] = get_input_for_key(api_key, key)

# Wrap the payload in list
payload = [single_payload]

print('\n')
time.sleep(1)

# Send request
json_payload = json.dumps(payload)

response = requests.post(url, headers=headers, data=json_payload)
print("\nResponse Status:", response.status_code)

if response.status_code == 201:
    print("    Asset Generated in Database. SUCCESS!")

elif response.status_code == 400 or 409:
    print("! ERROR | ERROR !")
    print(' ')

else:
  print("TRY AGAIN...")


response_dict = json.loads(response.text)
parts = response.text.split('"Message":')

# Split
if len(parts) > 1:
    message_part = parts[1].split(',', 1)[0]
    message = message_part.split('"')[1] 

    print("Error Message:", message)

else:
    print("__________")

time.sleep(33)