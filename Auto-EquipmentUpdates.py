# This script automates a documentation process by sending a PUT request that updates a large group of equipment based on User input. Script documents results on excel sheet and in database.

import time
import os
import requests
import json
import pandas as pd

# Define desktop directory
desktop_dir = os.path.expanduser("... FILEPATH ...\\Desktop")

# URL and headers
headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ... APIkey ...',
    'Connection': 'keep-alive'
}

url = 'http://... URL .../assets?$filter=ParentRef.ID%20eq%20"CV-GG-4-"'

# Payload for the initial request
payload = {
    "Results": [{}],
    "Total": 0
}

# Send request and save the response to file
response = requests.request("GET", url, headers=headers, data=json.dumps(payload))
file_path = os.path.join(desktop_dir, 'rawData.txt')
with open(file_path, 'w') as file:
    file.write(response.text)

# Title
title = "PIPETTE LOGGER"
delay = 0.1
for char in title:
    print(char, end='', flush=True)
    time.sleep(delay)
print(' ')

# User input for calibration dates
print("Enter Date Pipettes Were Calibrated: ")
calDate = str(input())
print(' ')

print("Enter Due-Date for the Next Calibration: ")
nextcalDate = str(input())
print(' ')

# Payload for PUT request
payloadPUT = {
    'UDFDate3': calDate,
    'UDFDate4': nextcalDate
}

# Read JSON data from file
with open(file_path, 'r') as file:
    json_data = file.read()

# JSON data dictionary
data_dict = json.loads(json_data)

print(' ')
print('Working...')
# Iterate over items and update data
for item in data_dict['Results']:
    PK = item.get("PK")
    print("Pipette " + str(PK))
    
    # Construct URL for PUT request
    url = 'http://... URL .../assets/' + str(PK)
    print("Requesting... " + url)
    time.sleep(.5)
    
    # Send the PUT request and handle the response
    responsePUT = requests.request("PUT", url, headers=headers, data=json.dumps(payloadPUT))

    if responsePUT.status_code == 200:
        print("PUT request successful.")
    else:
        print("PUT request failed with status code:", responsePUT.status_code)
        print("Response:", responsePUT.text)
    time.sleep(0.5)

print(' ')
print('Working...')
# GET with new info
url = 'http://... URL .../assets?$filter=ParentRef.ID%20eq%20"CV-GG-4-"'

payload = {
    "Results": [{}],
    "Total": 0
}

response = requests.request("GET", url, headers=headers, data=json.dumps(payload))

finalOutput = response.text
data_dict2 = json.loads(finalOutput)

Finaldf = pd.json_normalize(data_dict2['Results'])

csv_path = os.path.join(desktop_dir, 'Pipette_Report.csv')
Finaldf.to_csv(csv_path, index=False)

print(f"Updated Pipette List Saved to '{csv_path}'")
print(' ')
print(Finaldf)
time.sleep(3)

print(' ')
print("COMPLETE! See Updates in the Output File or in Maintenance Connection...")
print(' ')
print(f"Output File: '{csv_path}'")
time.sleep(5)
print("...TERMINATING")
time.sleep(2)
