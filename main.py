import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#Grabbing data from excel sheet
df = pd.read_excel('DataJewellNew.xlsm', sheet_name="Invoice", usecols="D,J")

#Create a dictionary with a list of emails and their invoice IDs
emailDict = {}
for index, row in df.iterrows():
    emailDict.setdefault(row['invoice_email'], [])
    emailDict[row['invoice_email']].append(row['sch_id'])

#print(emailDict)

#Email message
SUBJECT = ""
msg = MIMEMultipart()
msg['Subject'] = SUBJECT
msg['From'] = "sender@example.org"  #Email sender goes here
filename = "PW-2024-"

#Sending the email
for key, value in emailDict.items():
    files = []
    msg['To'] = key  #Email reciever goes here
    print(msg['To'])
    for i in value:
        files.append("{}{}.docx".format(filename, i))
        print(files)

# for i in data:
#     print(i)
# 
#     with smtplib.SMTP(host="localhost", port=25) as server:
#         server.ehlo()
#         server.sendmail("admin@example.org", i, "TEST EMAIL")
#
