import pdfplumber
import sqlite3
sno1 = []

pdf_path = "C:/Users/PioneerGuest/Downloads/fedex inv1.pdf"

tracking_info ={}

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        lines = text.split('\n')
        for sno, line in enumerate(lines):
            if line.startswith('Ship Date'):
                #sno1.append(sno)
                #sno1.append(sno+7)
                for i in range(sno,sno+7):
                    if lines[i].startswith('Tracking'):
                        Traking_id = lines[i].split()
                        Traking_id = Traking_id[2]
                        # Add tracking ID and PDF path to the dictionary
                        tracking_info[Traking_id] = pdf_path


#print(tracking_info)




# Connect to SQLite database
conn = sqlite3.connect('tracking_data.db')
c = conn.cursor()

# Create table if it doesn't already exist
c.execute('''CREATE TABLE IF NOT EXISTS tracking (tracking_id TEXT, pdf_path TEXT)''')

# Insert tracking data into the database
for tracking_id, path in tracking_info.items():
    c.execute("INSERT INTO tracking (tracking_id, pdf_path) VALUES (?, ?)", (tracking_id, path))

# Commit changes and close the connection
conn.commit()
conn.close()

                        