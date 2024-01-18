# This script reads pdf report and updates Maintenance Connection to reflect the report. Script documents results on excel sheet. Automates process for monthly equipment calibration reports:   

import os
import time
import json
import requests
import pandas as pd
import pdfminer
from pdfminer.high_level import extract_text

title = "PROCAL PRO:"
print("___________________")
delay = 0.1
for char in title:
    print(char, end='', flush=True)
    time.sleep(delay)

desktop_dir = os.path.expanduser(r"... FILEPATH ...\MC Reports")
desktop_dir2 = os.path.expanduser(r"... FILEPATH ...")

headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ... APIkey ...',
    'Connection': 'keep-alive'
}

# Function to convert PDF to text
def pdf_to_text(path):
    text = extract_text(path)
    return text

print(' ')
print("Paste File Path of Procal Report pdf")
pdfFilePath = str(input())

# Convert the PDF to text and save it
pdf_path = pdfFilePath
text = pdf_to_text(pdf_path)

# Write the extracted text to a file
with open(os.path.join(desktop_dir2, 'doc.txt'), 'w') as f:
    f.write(text)

print(' ')
print("*Loading Report*")

# List of lab IDs
labs = ['4101', '4103', '4105', '4121', '4123', '4125', '4130', '4141', '4143', '4145', '4153', '4251', '4301', '4303', '4305', '4320-1', '4321', '4323', '4329', '4340', '4341', '4343', '4345', '4401', '4404', '4406', '4408', '4424', '4426', '4428', '4444', '4446', '4448', 'Misc.-']

# Collecting data
all_data = []
load = ''
for lab in labs:
    url = f'... URL .../assets?$filter=ParentRef.ID%20eq%20"CV-GG-4-{lab}"'
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        load += '*'
        print('\r'+ load, end='')
        data = response.json()
        results = data['Results']
        for item in results:
            all_data.append({
                'PK': item.get('PK', None),
                'ID': item.get('ID', None),
                'Name': item.get('Name', None),
                'Model': item.get('Model', None),
                'UDFChar4': item.get('UDFChar4', None),
                'UDFDate4': item.get('UDFDate4', None),
                'UDFDate3': item.get('UDFDate3', None)
            })
    else:
        print(f"Error: Request for lab {lab} failed with status code {response.status_code}")

df1 = pd.DataFrame(all_data)

# Renaming columns and converting date formats
df1.rename(columns={'UDFChar4': 'Serial#', 'UDFDate4': 'Calibration Date', 'UDFDate3': 'Last Verification'}, inplace=True)
df1['Calibration Date'] = pd.to_datetime(df1['Calibration Date']).dt.date
df1['Last Verification'] = pd.to_datetime(df1['Last Verification']).dt.date

# Modify the DataFrame
df1.rename(columns={'Calibration Date':'Calibration', 
                    'Serial#': 'M#'}, inplace=True)

# Save the modified DataFrame as CSV
datacsv_path = os.path.join(desktop_dir2, 'dataMC.csv')
df1.to_csv(datacsv_path, index=False)

# Read and process the text file 'doc.txt'
with open(os.path.join(desktop_dir2, "doc.txt"), "r") as file:
    document = file.read()

# Split the document into separate reports
reports = document.split("INSTRUMENT CALIBRATION REPORT")
parsed_sections = []

# Iterate over the reports
for report in reports:
    if "Instrument ID" in report and "Next Cal Date" in report:
        start = report.index("Instrument ID")
        end = report.index("Next Cal Date") + len("Next Cal Date")
        end_line = end
        for _ in range(4):
            end_line = report.index("\n", end_line + 1)
        parsed_section = report[start:end_line].strip().split("\n")
        parsed_section = [line.strip() for line in parsed_section if line.strip() != ""]
        parsed_sections.append(parsed_section)

# Create a DataFrame from the parsed sections
df = pd.DataFrame(parsed_sections)

# Drop unnecessary columns if there are more than 20 columns
num_columns = len(df.columns)
if num_columns > 20:
    cols_to_drop = list(range(0, 3)) + list(range(6, 16)) + list(range(22, 24)) + list(range(17, 18))
    cols_to_drop = [col for col in cols_to_drop if col < num_columns]
    df.drop(df.columns[cols_to_drop], axis=1, inplace=True)

df.rename(columns={3: 'M#'}, inplace=True)

procalcsv_path = os.path.join(desktop_dir, 'ProCal-Chart.csv')
df.to_csv(procalcsv_path, index=False)

# Merge the DataFrames
dfmc = pd.read_csv(datacsv_path)
dfpro = pd.read_csv(procalcsv_path)
merged_df = pd.merge(dfmc, dfpro, on='M#')

columns_to_drop = ['Serial', 'SAP Asset ID', 'Manufacturer Name', 'Vicinity', '4']
merged_df.drop(columns=columns_to_drop, inplace=True, errors='ignore')

print(' ')
print("PREVIEW ")
print(merged_df)
summarycsv = os.path.join(desktop_dir, 'CalibrationSummary.csv')
merged_df.to_csv(summarycsv, index=False)

time.sleep(1)
print(' ')

print(' ')
print(f"Full Procal Report File:{procalcsv_path}")
print(f"Maintenance Connection Report File:{summarycsv}")
time.sleep(3)

print(' ')
print('UPDATING MAINTENANCE CONNECTION')
print(' ')
time.sleep(1)

df = pd.read_csv(summarycsv)

for index, row in df.iterrows():
    pk = row['PK']
    next_cal_date = row['21']  # The column header for the Next Cal Date seems to be '21'
    last_verification_date = row['5']  # The column header for the Calibration Last Verification seems to be '5'

    # Check for missing or 'nan' values
    if pd.isnull(next_cal_date) or pd.isnull(last_verification_date):
        print(f"Skipping PK: {pk} due to missing date value.")
        continue

    # Ensure that the date values are valid dates, otherwise skip the PUT request
    try:
        # Attempt to convert the dates to a string in the expected date format
        payloadPUT = {
            'UDFDate5': pd.to_datetime(next_cal_date).strftime('%Y-%m-%d'),
            'UDFDate3': pd.to_datetime(last_verification_date).strftime('%Y-%m-%d'),
        }
    except ValueError as e:
        # If there is an error in date conversion, print the error and skip this row
        print(f"Error converting dates for PK: {pk}. Error: {e}")
        continue

    # Construct the URL for the PUT request using the PK value
    url = f'http:... URL .../assets/{pk}'
    print(f"Updating PK: {pk}")
    
    # Send the PUT request and handle the response
    responsePUT = requests.put(url, headers=headers, data=json.dumps(payloadPUT))

    # Check if the request was successful
    if responsePUT.status_code == 200:
        print("PUT request successful.")
    else:
        print(f"PUT request failed with status code: {responsePUT.status_code}")
        print(f"Response: {responsePUT.text}")

    time.sleep(0.5)

file_path = os.path.join(desktop_dir2, 'rawData.txt')
all_data = []

print(' ')
print('Creating Full Report...')
print(' ')

for lab in labs:
    url = f'... URL .../assets?$filter=ParentRef.ID%20eq%20"CV-GG-4-{lab}"'

    # Send request
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        print(f"Pulling Equipment Data From {lab}")
        data = response.json()
        results = data['Results']
        for item in results:
            # Extract keys and append to all_data
            all_data.append({
                'PK': item.get('PK', None),
                'ID': item.get('ID', None),
                'Name': item.get('Name', None),
                'Model': item.get('Model', None),
                'UDFChar4': item.get('UDFChar4', None),
                'UDFDate4': item.get('UDFDate4', None),
                'UDFDate3': item.get('UDFDate3', None)
            })
    else:
        print(f"Error: Request for lab {lab} failed with status code {response.status_code}")

df = pd.DataFrame(all_data)

df.rename(columns={'UDFChar4': 'Serial#',
                   'UDFDate4': 'Calibration Date',
                   'UDFDate3': 'Last Verification'}, inplace=True)

# Convert the 'UDFDate4' column to datetime and format to date only
df['Calibration Date'] = pd.to_datetime(df['Calibration Date']).dt.date
df['Last Verification'] = pd.to_datetime(df['Last Verification']).dt.date

# Save to file
df.to_csv(os.path.join(desktop_dir, 'TFN_Equipment_Report.csv'), index=False)

with open(file_path, 'a') as file:
    for entry in all_data:
        file.write(json.dumps(entry) + '\n')

print(' ')
print("   COMPLETE! See Updates in the Output Files or in Maintenance Connection...")
print(' ')
time.sleep(25)
print("...TERMINATING")
print("___________________")
time.sleep(2)
