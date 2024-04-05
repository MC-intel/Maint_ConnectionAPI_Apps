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

