
# import the required libraries
import pandas as pd
import smtplib
import email.mime
import mime
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# change these as per use
your_email = "yourmail"
your_password = "yourpass"

# connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(user="yourmail" , password="yourpass")

# reading the spreadsheet
email_list = pd.read_excel('Excel file location')
# getting the names and the emails from the excel file
names = email_list['NAME']
emails = email_list['EMAIL']
# HTML (you will use this later as the body in the mail
html = """
test
"""
# read through the excel
for i in range(len(emails)):
    # for every record get the name and the email addresses
    name = names[i]
    email = emails[i]
    # the message to be emailed
    message = MIMEMultipart()
    message['From'] = your_email
    message['To'] = email
    message['Subject'] = "Test"
    message.attach(MIMEText("Hey {}, ".format(name), 'plain'))
    message.attach(MIMEText(html, 'html'))

    # sending the email
    server.send_message(message)
# close the smtp server
server.close()
