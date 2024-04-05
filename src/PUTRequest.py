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

