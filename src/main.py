import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import os
import json
import argparse
import logging
import re

# Configuration des logs
email_log_error = logging.getLogger('email_errors')
email_log_error.setLevel(logging.ERROR)
email_handler_error = logging.FileHandler('email_errors.log')
email_handler_error.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
email_log_error.addHandler(email_handler_error)

email_log = logging.getLogger('email_log')
email_log.setLevel(logging.INFO)
email_handler = logging.FileHandler('email_sent.log')
email_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
email_log.addHandler(email_handler)

def is_valid_email(email):
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(email_regex, email) is not None

def send_email(sender_email, sender_password, receiver_email, subject, body, attachment_paths=None):
    try:
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject

        # Le corps du message est en HTML
        message.attach(MIMEText(body, 'html', 'utf-8'))
        if attachment_paths:
            for attachment_path in attachment_paths:
                filename = os.path.basename(attachment_path)
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                    message.attach(part)

        text = message.as_string()
        email_log.info(f"Attempting to send email to {receiver_email} with subject '{subject}' \n {body}\n")
        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, text)
        server.quit()

        email_log.info(f"Email successfully sent to {receiver_email} with subject '{subject}'")
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        email_log_error.error(f"Failed to send email to {receiver_email} with subject '{subject}' due to {str(e)}")
        print(f"Failed to send email to {receiver_email}, check log for details.")

def main(excel_file, auth_file, subject, body_file, attachment_paths=None):
    with open(auth_file, 'r', encoding='utf-8') as f:
        auth_data = json.load(f)
    sender_email = auth_data['email']
    sender_password = auth_data['password']
    
    with open(body_file, 'r', encoding='utf-8') as f:
        body_template = f.read()

    df = pd.read_excel(excel_file)
    df_filtered = df[df['PUSH'] == 'x']

    for index, row in df_filtered.iterrows():
        if pd.notna(row['mail']) and is_valid_email(row['mail']):
            receiver_email = row['mail']
            salutation = "Monsieur" if row['Genre'] == 'M' else "Madame"

            # Incorporation du template du corps de l'email
            body = f"""
            <html>
            <body>
                <p><b>{salutation}</b>,</p>
                <p>{body_template}</p>
                <p><u>Merci de votre attention.</u></p>
            </body>
            </html>
            """

            send_email(sender_email, sender_password, receiver_email, subject, body, attachment_paths)
            time.sleep(2) 
        else:
            print(f"Skipping row {index}: Invalid or missing email address.")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Send emails with attachments to a list of addresses from an Excel file.')
    parser.add_argument('excel_path', type=str, help='Path to the Excel file containing email addresses')
    parser.add_argument('auth_file', type=str, help='Path to the JSON file containing email authentication information')
    parser.add_argument('body_file', type=str, help='Path to the file containing the email body')
    parser.add_argument('--subject', type=str, default='Default Subject', help='Subject of the email (default: "Default Subject")')
    parser.add_argument('--attachments', type=str, nargs='*', default=[], help='Paths to the attachment files (optional)')

    args = parser.parse_args()
    
    main(args.excel_path, args.auth_file, args.subject, args.body_file, args.attachments)