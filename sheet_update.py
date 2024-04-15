#!/usr/bin/env python
# -*- coding: utf-8 -*-

#import all the packages
import imaplib, email, requests, quopri, re
import gspread
import time
import json
from collections import OrderedDict
import smtplib
import random
from google.oauth2 import service_account
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import google.auth.transport.requests
from google.oauth2.service_account import Credentials
import google.oauth2.credentials
import imaplib
import email
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()  # take environment variables from .env.

#Scratch email configuration
user = os.getenv("USER")
password = os.getenv("USER_APP_PASSWORD")
# sender = 'deals@buyformeretail.com'
sender = os.getenv("MAIL_SENDER")
# receive_email = 'winterpeas@gmail.com'
# receive_email= "sunnie.gpsh@gmail.com"
receive_email= "ladayodeji@gmail.com"

#  IMAP server conf
imap_ssl_host = 'imap.gmail.com'
imap_ssl_port = 993

# SMTP server configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587


SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Function to send an email with the sheet link URL
def send_email(body):
    # Prepare the email message
    subject = 'Updated Google Sheet Attached'
    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = receive_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    # Send the email via SMTP
    try:

        server = smtplib.SMTP(smtp_server, smtp_port)
        print(f"Starting mail server")
        server.starttls()
        server.login(user, password)
        print(f"{user} logged in")
        server.sendmail(user, receive_email, msg.as_string())
        print(f"Email sent successfully to {receive_email}")

    except Exception as e:
        print(f"An error occurred while sending the email: {e}")

    finally:
        server.quit()

def get_spreadsheet():
    start_time= time.perf_counter() 
    print("Connecting to mail Server")
    
    server = imaplib.IMAP4_SSL(imap_ssl_host, imap_ssl_port)
    server.login(user, password)
    print(f"{user}: Login Successfully")
    
    server.select('INBOX')

    # Query search based in Time Frame and particular sender
    days= 30
    date_since = (datetime.today() - timedelta(days=days)).strftime('%d-%b-%Y')
    
    search_query = f'(FROM "{sender}" SINCE "{date_since}")'
    
    result, msgs = server.search(None, search_query)


    sheet_url_id=[]
    if result == "OK":
        email_ids = msgs[0].split()

        if len(email_ids)<=0:
            print(f"No mail received from {sender}  in past {days} days.")

        print(f"{len(email_ids)} mail messages received from {sender} in past {days} days")
         
        for email_id in email_ids:
            url_maid_id=[]
            result, data= server.fetch(email_id, ("RFC822"))

            message= email.message_from_bytes(data[0][1])

            # print(f"Message ID: {email_id}")
            # print(f"From: {message.get('From')}")
            # print(f"Date: {message.get('Date')}")
            # print(f"To: {message.get('TO')}")
            # print(f"BCC: {message.get('BCC')``}")
            # print(f"Subject: {message.get('SUBJECT')}")
            # print("Content:")
            for part in message.walk():
                if part.get_content_type()== "text/plain" or part.get_content_type()== "text/html":
                    data= part.get_payload(decode=True).decode()
                    # print(data)
                    url_matches = re.findall(r'<https://docs\.google\.com/spreadsheets/.*', data)
                    sheet_url = str(url_matches[0]).replace("<", "").replace(">", "")
                    # print(email_id, sheet_url)
                    
                    email_id= email_id.decode('utf-8')
                    url_maid_id+=[email_id, sheet_url, message.get('Date')]
                    break
            # print()    
            sheet_url_id+=[url_maid_id]
        # print(sheet_url_id)

    end_time= time.perf_counter()

    server.close()
    print('Mail server Closed')
    print(f"The mail extraction  took {end_time - start_time:.2f} seconds to complete.\n")
    # print(sheet_url_id)
    return sheet_url_id

def is_sheet_processed(spreadsheet_id):
    
    # Procesed mail id  file name
    file_name = "processed_email_id.txt"

    # Use os.path.exists() to check if the file exists
    if not os.path.exists(file_name):
        print(f"Creating Processed mail id text file ")
        # If the file does not exist, create it using the 'with' statement
        with open(file_name, 'w') as f:
            pass

    # Check if  email_id exists in processed_email_id

    processed_mail_id=[]
    with open('processed_email_id.txt', 'r') as mail_ids:
        list= []
        # Iterate over the file line by line

        for id in mail_ids:
            cleaned_line= id.strip()
            list+=[cleaned_line]
        processed_mail_id+=list

        if spreadsheet_id in processed_mail_id:
            return True
        else:
            return False
   
     
    

def update_sheet(spreadsheet_list, values):
    start_time= time.perf_counter()
    sheet_url_id = [item for item in spreadsheet_list if item!=[]]
    # print(sheet_url_id)
    unprocessed_mail_id= [((bytes(sheet_url_id[i][0], 'ascii')),sheet_url_id[i][1], sheet_url_id[i][-1] ) for i in range(len(sheet_url_id))  if is_sheet_processed(sheet_url_id[i][0]) is False]

    sheet_count= 0
    

    # print(unprocessed_mail_id)
    print(f"Updating unprocessed mails")
    print(f"Found {len(unprocessed_mail_id)} unprocessed mails")

    processed_mail_id=[]
    if len(unprocessed_mail_id)<=0:
        print(f"No Mail to process")

    bfmr_id_col=[]
    processed_mail_id=[]

    for id, link, date in unprocessed_mail_id:

        
        credentials = Credentials.from_service_account_file(os.getenv("GOOGLE_APPLICATION_CREDENTIALS"),
                                                            scopes= SCOPES)
        sheet_count+=1
        
        print(f"Updating Mail id:{id}: Received on {date}")
        client= gspread.authorize(credentials)

        sheet_url = link
        workbook = client.open_by_url(sheet_url)
        workbook_sheet= workbook.worksheet("Sheet1")

        all_values= workbook_sheet.get_all_values()
        
        update_column_name= "Total Needed" 
        for col_index, column in enumerate(zip(*all_values), start=1):
            if update_column_name in column:
                # Calculate the 1-based index of the column where the value is found
                found_col_index = col_index
                found_col_alphabet = chr(65 + col_index - 1)
                bfmr_id_col+=[[found_col_index, found_col_alphabet]]
                print(f"'{update_column_name}' found in column {found_col_alphabet}")
                
            else:
                pass
        # Use OrderedDict to preserve the original order and remove duplicates
        bfmr_id_col = list(OrderedDict((tuple(item), item) for item in bfmr_id_col).values())
        # print(bfmr_id_col)
            
        values = values
        # values = [['New Value', 123, 234], ['New Value', 678, 345], ["New Vakue", 900, 789]]

        lst=[]
        processd_id=[]
        for columns_to_update in range(len(bfmr_id_col)):
            
            update_column= workbook_sheet.col_values(bfmr_id_col[columns_to_update][0])
            
            if values[columns_to_update][0]  in update_column:
                lst+=[str(id.decode('ascii'))]
                print(f"{values[columns_to_update][0]} exists in column {bfmr_id_col[columns_to_update][1]}: No data entry here.")
            else:
            
                # Specify the range of cells to update (e.g., 'A5:C5' for row 5)
                # range_to_update = f"F{len(first_col)+1}:{chr(ord('F') + len(values[0]) - 1)}{len(first_col)+1}"

                range_to_update = f"{bfmr_id_col[columns_to_update][1]}{len(update_column)+1}:{chr(ord(str(bfmr_id_col[columns_to_update][1])) + len(values[columns_to_update]) - 1)}{len(workbook_sheet.col_values(bfmr_id_col[columns_to_update][0]))+1}"
            

                print(f"{values[columns_to_update][0]} does not exists in column {bfmr_id_col[columns_to_update][1]}: Column Range: {range_to_update}")

                workbook_sheet.update(range_name=range_to_update, values=[values[columns_to_update]])


                print(f"Inserting {values[columns_to_update]} into Row {range_to_update} ")
                print(f"Mail id:{id} Updated Succesfully")
                lst+=[str(id.decode('ascii'))]
                # send_email(body= link)

                # send_email(body= str(bfmr_id_col[columns_to_update][-1]))

            processed_mail_id+=(lst)
        send_email(body= link)

    processed_mail_id = list(OrderedDict((tuple(item), item) for item in processed_mail_id).values())
    print()
    # print(processed_mail_id)


    try:
        # Open the file in append mode ('a')
        with open("processed_email_id.txt", 'a') as file:
            # Write the content to the file
            
            for ids in processed_mail_id:
                # processed_id=str(id.decode('ascii'))
                file.write(ids + "\n")  # Append a newline character after the content
        if len(processed_mail_id)<=0:
            print(f"No mail appended to Processed Mails.")
        else:
            print(f"Processed mails appended to Processed Mails.")
    except Exception as e:
        print(f"An error occurred: {e}")
    stop_time= time.perf_counter()
    print(f"Update Completed in {stop_time - start_time:.2f} seconds")


if __name__ == "__main__":
    sheet_list= get_spreadsheet()
    values= [['B-22456', 239], ['B-644321', 634], ['B-967060', 129]] # [[Deal 1], [Deal 2], [Deal 3]]
    update_sheet(sheet_list, values= values)