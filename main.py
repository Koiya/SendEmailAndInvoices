from email.mime.application import MIMEApplication

import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#initialize settings
filename = "PW-2024-"
SUBJECT = "Sending invoices"
#Replace variable without outlook email and password
sender = "admin@example.org"
password = "123"

# Outlook SMTP details
# Domain: smtp-mail.outlook.com
# Port: 587
# Security protocol: TLS

#SMTP server settings
host = "localhost"
port = 25

#Grabbing data from excel sheet
df = pd.read_excel('DataJewellNew.xlsm', sheet_name="Invoice", usecols="D,J")

#Create a dictionary with a list of emails and their invoice IDs
emailDict = {}
for index, row in df.iterrows():
    emailDict.setdefault(row['invoice_email'], [])
    emailDict[row['invoice_email']].append(row['sch_id'])

print(emailDict)

#Sending the email
for key, value in emailDict.items():
    #Email message
    msg = MIMEMultipart()
    msg['Subject'] = SUBJECT
    msg['From'] = sender 
    #Add body message here
    body = MIMEText('Test', 'html', 'utf-8')
    msg.attach(body)

    for i in value:
        file = "{}{}.docx".format(filename, i)
        attachment = MIMEApplication(open(file, "rb").read(), _subtype="docx")
        attachment.add_header('Content-Disposition', 'attachment', filename=file)
        msg.attach(attachment)
        print(file)

    #connects to smtp server and send email
    with smtplib.SMTP(host=host, port=port) as server:
        server.starttls()
        server.login(sender,password)
        server.ehlo()
        server.sendmail(msg['From'], key, msg.as_string())
