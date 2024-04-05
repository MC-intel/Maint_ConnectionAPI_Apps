import re
import os
import time
import requests
import json
from datetime import datetime
import pandas as pd

import pdfminer
from pdfminer.high_level import extract_text


desktop_dir = os.path.expanduser(r"C:\Users\USER\OneDrive - ")
desktop_dir2 = os.path.expanduser(r"C:\Users\USER\OneDrive - ")

headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic APIKEY',
    'Connection': 'keep-alive'
}

labs = ['4101', '4103', '4105', '4121', '4123', '4125', '4130', '4141', '4143', '4145', '4153', '4251', '4301', '4303', '4305', '4320-1', '4321', '4323', '4329', '4340', '4341', '4343', 
          '4345', '4401', '4404', '4406', '4408', '4424', '4426', '4428', '4444', '4446', '4448', '3220']


def process_pdf(pdf_path): # Converts PDF to text and saves output to doc.txt.
    text = extract_text(pdf_path)
    with open(os.path.join(desktop_dir2, 'doc.txt'), 'w') as f:
        pdf = f.write(text)

    df = extract_pdf(pdf)
    return df

def extract_pdf(doc): # Parses through doc.txt and creates a dataframe from the extracted data.
    print('\nreadingPDF...')
    time.sleep(1)
    with open(os.path.join(desktop_dir2, 'doc.txt'), 'r') as f:
        document = f.read()

    data_dict = {'caldates': [], 'M': [], 'nextcal': []}

    reports = document.split("Bayer Metrology\n\nINSTRUMENT CALIBRATION REPORT")[1:]

    for x in range(len(reports)):
        clean = reports[x].split('Calibrated')[1:]
        clean = clean[0].split('Manufacturer')[0].replace('\n', ' ').strip(' ')
        clean = clean.split(' ')[:7]
        data_dict['M'].append(clean[0])

        date_pattern = r'\b\d{1,2}-[a-zA-Z]+-\d{4}\b'

        for string in clean:
            match = re.search(date_pattern, string)
            if match:
                date_str = match.group(0)
                date = datetime.strptime(date_str, '%d-%b-%Y')
                data_dict['caldates'].append(date)

        clean = reports[x].replace('\n\n',' ')
        clean = clean.replace('\n', ' ')
        clean = clean.split('Calibrated')
        clean = clean[1].split('Model Number')[:1]
        clean = clean[0].split('Next Cal Date')
        clean = clean[1].split(' ')
        clean = clean[1]

        try:
            datetime.strptime(clean, '%d-%B-%Y')
            datee = True
        except ValueError:
            datee = False

        if datee == True:
            data_dict['nextcal'].append(clean)
        else:
            data_dict['nextcal'].append('NoDate')

    PDFdf = pd.DataFrame(data_dict)
    print('\nThe following data was pulled from the Procal Report: ')
    print(PDFdf)
    PDFdf.to_csv(os.path.join(desktop_dir2, 'extractedList.csv'), index=False)
    time.sleep(3)
    return PDFdf

# Run functions
title = "PROCAL PRO:"
print("___________________")
delay = 0.1
for char in title:
    print(char, end='', flush=True)
    time.sleep(delay)

print("\nPaste File Path of Procal Report pdf")
pdfFilePath = str(input())

result_df = process_pdf(pdfFilePath)
print(' ')


def pullFromAPI(): # Pulls data from Maintenance Connection and creates a dataframe.
  print(' ')
  print("*Loading*")

  all_data = []
  load = ''
  for lab in labs:
      url = f'http://stlwcmmsprd01/v8/assets?$filter=ParentRef.ID%20eq%20"CV-GG-4-{lab}"' if lab != '3220' else f'http://stlwcmmsprd01/v8/assets?$filter=ParentRef.ID%20eq%20"CV-GG-3-3220"'
      try:
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
                      'M': item.get('UDFChar4', None)
                  })
          else:
              print(f"Error: Request for lab {lab} failed with status code {response.status_code}")
      except Exception as e:
          print(f"An error occurred while processing lab {lab}: {str(e)}")

  APIdf = pd.DataFrame(all_data)
  return APIdf
  print(' ')

# Run function and find which units in the Procal report are in Maintenance Connection by merging the two dataframes and making a list of matching values.
GETdf = pullFromAPI()

merged_df = pd.merge(GETdf, result_df[['M', 'nextcal', 'caldates']], how='inner', on='M')
final_df = merged_df[['PK', 'Name', 'M', 'nextcal', 'caldates']]

print('\nUpdating the following Equipment: ')
print(final_df)
final_df.to_csv(os.path.join(desktop_dir2, 'updatedUnits.csv'), index=False)
time.sleep(5)
print(' ')


def updateMC(): # Update Maintenance Connection with the new calibration dates based off the merged data.
  for index, row in final_df.iterrows():
    pk = row['PK']
    nextcal = row['nextcal']
    caldate = row['caldates']
    nm = row['Name']

    if pd.isnull(nextcal) or pd.isnull(caldate):
      print(f"Skipping: {pk} due to missing date value.")
      time.sleep(0.5)
      continue

    try:
        payloadPUT = {
            'UDFDate5': pd.to_datetime(nextcal).strftime('%Y-%m-%d'),
            'UDFDate3': pd.to_datetime(caldate).strftime('%Y-%m-%d'),
        }
    except ValueError as e:
        print(f"Error converting dates for PK: {pk}. Error: {e}")
        time.sleep(5)
        continue

    url = f'http://stlwcmmsprd01/v8/assets/{pk}'
    print(f"Updating: {nm}")

    responsePUT = requests.put(url, headers=headers, data=json.dumps(payloadPUT))

    if responsePUT.status_code == 200:
        print("PUT request successful.")
        time.sleep(0.5)
    else:
        print(f"PUT request failed with status code: {responsePUT.status_code}")
        print(f"Response: {responsePUT.text}")
        time.sleep(0.5)

updateMC()

time.sleep(1)
print('\nCreating Full Report...')
print(' ')


def tfnReport(): # Creates an excel report of all departments equipment reflecting the latest updates.
  all_data = []
  for lab in labs:
      url = f'http://stlwcmmsprd01/v8/assets?$filter=ParentRef.ID%20eq%20"CV-GG-4-{lab}"' if lab != '3220' else f'http://stlwcmmsprd01/v8/assets?$filter=ParentRef.ID%20eq%20"CV-GG-3-3220"'
      response = requests.get(url, headers=headers)

      if response.status_code == 200:
          print(f"Pulling Equipment Data From {lab}")
          data = response.json()
          results = data['Results']
          for item in results:
              all_data.append({
                  'LabID': lab, 
                  'PK': item.get('PK', None),
                  'ID': item.get('ID', None),
                  'Name': item.get('Name', None),
                  'Model': item.get('Model', None),
                  'M#': item.get('UDFChar4', None),
                  'LastVerification': item.get('UDFDate3', None),
                  'CalibrationDue': item.get('UDFDate5', None)
              })
      else:
          print(f"Error: Request for lab {lab} failed with status code {response.status_code}")

  df = pd.DataFrame(all_data)
  df['CalibrationDue'] = pd.to_datetime(df['CalibrationDue']).dt.date
  df['LastVerification'] = pd.to_datetime(df['LastVerification']).dt.date

  file_path = os.path.join(desktop_dir2, "TFN_Equipment_Report.xlsx")

  # Save to Excel with each lab on a separate sheet.
  with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
      unique_labs = df['LabID'].unique()
      for lab in unique_labs:
          lab_df = df[df['LabID'] == lab]
          lab_df.to_excel(writer, sheet_name=str(lab), index=False)

  print(f"\nData has been saved to '{file_path}', with separate sheets for each lab.")

tfnReport()

print(' ')
print("   COMPLETE! See Updates in the Output Files or in Maintenance Connection...")
print(' ')
time.sleep(25)
print("...TERMINATING")
time.sleep(2)
