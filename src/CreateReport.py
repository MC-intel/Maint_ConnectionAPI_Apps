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