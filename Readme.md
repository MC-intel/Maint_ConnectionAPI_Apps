## Info

This repo consists of modules built to streamline tedious documentation workflows involving a web-based equipment database. Maintenance Connection is a critical reference during audits. 
Brief descriptions of the functionalities are in the code files. Examples are of deployed scripts that are ran locally through a secure, onsite host-computer. 

## RUN

```python 

# Dependencies
import re
import os
import time
import requests
import json
from datetime import datetime
import pandas as pd

import pdfminer
from pdfminer.high_level import extract_text

# Imports: security policies may vary regaurding git modules
from Maint.Connection-API-Apps.FullScript import process_pdf
from Maint.Connection-API-Apps.FullScript import extract_pdf
from Maint.Connection-API-Apps.FullScript import pullFromAPI
from Maint.Connection-API-Apps.FullScript import updateMC
from Maint.Connection-API-Apps.FullScript import tfnReport


# Globals
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

GETdf = pullFromAPI()

merged_df = pd.merge(GETdf, result_df[['M', 'nextcal', 'caldates']], how='inner', on='M')
final_df = merged_df[['PK', 'Name', 'M', 'nextcal', 'caldates']]

print('\nUpdating the following Equipment: ')
print(final_df)
final_df.to_csv(os.path.join(desktop_dir2, 'updatedUnits.csv'), index=False)
time.sleep(5)
print(' ')

updateMC()

time.sleep(1)
print('\nCreating Full Report...')
print(' ')

tfnReport()

print(' ')
print("   COMPLETE! See Updates in the Output Files or in Maintenance Connection...")
time.sleep(25)
print("...TERMINATING")
time.sleep(2)

```
