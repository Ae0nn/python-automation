import os
import json
import smtplib
#import ssl
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def server_configurations(configfile):
    '''
    Server_configurations reads into config.json file to retrieve smtp server configs.
    Since our SMTP server requires authentication, credentials are stored seperately.  
    '''
    with open(configfile) as f:
        return json.load(f)
    
def email_message(recipient_name):
    '''
    Email_message reads into html file containing email subject, body, and signature.
    Function utilizes MIME Class to create email message object to pass to function handling the mailing.
    '''
    #read into email html template and create multipart object
    with open('email_message.html') as content:
        email_content =  content.read()
    email = MIMEMultipart()
    email.attach(MIMEText(email_content, 'html'))

    #get path to excel file
    attachment_directory = os.path.join(os.getcwd(), 'email_attachments')
    file_path = os.path.join(attachment_directory, f'{recipient_name}.xlsx')
    print("File path:", file_path)

    #attach file to email
    if os.path.exists(file_path) and os.access(file_path, os.R_OK):
            try:
                with open(file_path, 'rb') as email_attachment:
                    attachment = MIMEBase('application', 'octet-stream')
                    attachment.set_payload(email_attachment.read())
                    #print(attachment.get_payload())
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename= {recipient_name}.xlsx')
                    email.attach(attachment)
            except (IOError, Exception) as exc:
                print('ERROR. Document was not attached. Something went wrong: ', exc)
    else:
        print(f"{recipient_name} has no associated document")
        print(f'{file_path} does not exist')
    return email
    
def automated_mailer():
    '''
    Automated_mailer, when called, builds then sends the email to the recipient mailbox. 
    '''
    subject= 'Test Subject'

    credentials = server_configurations('config.json')

    smtp_server = credentials.get('smtp_server')
    smtp_port =  credentials.get('smtp_port')
    smtp_username = credentials.get('smtp_username')
    smtp_password = credentials.get('smtp_password')

    with open('recipients.csv', newline='') as recipients:
        read_file = csv.reader(recipients)
        next(read_file)

        for item in read_file:
            recipient_name = item[0]
            recipient_email = item[1]
            #compiling the full message
            message = email_message(recipient_name)
            message['From'] = smtp_username
            message['To'] = recipient_email
            message['CC'] = 'sometestemail@test.com'
            message['Subject'] = subject
            #sending message to recipient(s)
            try:
                with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                    smtp.starttls()
                    smtp.login(smtp_username, smtp_password)
                    #smtp.send_message(message)
                    response = smtp.send_message(message)
                    print('SMTP Server Response:', response)
                    
                    print(f'Email sent successfully to {recipient_name} ({recipient_email})')
            except Exception as exc:
                print('ERROR. Email not sent. Something went wrong: ', exc)

def main():
    print('Initiating Mailer')
    automated_mailer()

if __name__ == '__main__':
    main()