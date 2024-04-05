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

