import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os


def send_email(sender_email, sender_password, receiver_email, subject, message, attachments):
    # Set up the SMTP server
    smtp_server = "smtp-mail.outlook.com"
    smtp_port =587   # For SSL, use port 465

    # Create a MIMEText object with the message
    email_message = MIMEMultipart()
    email_message['From'] = sender_email
    email_message['To'] = receiver_email
    email_message['Subject'] = subject
    email_message.attach(MIMEText(message, 'plain'))

    # Anexar cada arquivo ao email
    # Anexar cada arquivo ao email
    for attachment in attachments:
        file_path = attachment['filename']
        if file_path.endswith('.zip') and os.path.exists(file_path):
            part = MIMEBase('application', 'zip')
            with open(file_path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            filename = os.path.basename(file_path)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            email_message.attach(part)
        else:
            print(f"O arquivo {file_path} não é um arquivo zip ou não existe. Não será anexado ao e-mail.")



    # Start the SMTP session
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        # Login to the SMTP server
        server.login(sender_email, sender_password)
        # Send the email
        server.sendmail(sender_email, receiver_email, email_message.as_string())
        

