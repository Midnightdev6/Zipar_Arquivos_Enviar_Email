from email import encoders
from email.mime.base import MIMEBase
import os
import zipfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

directory = r'C:\\Users\\vitor.moreira\\OneDrive - Grupo Trinus Co\\Documentos - PRESTAÇÃO DE CONTAS BACKOFFICE'
output_folder = r'C:\\Users\\vitor.moreira\\Desktop\\Zip_Files'

def nome_pastas() :
    subdirectories = []
    # Walk through the directory and get all subdirectories
    for root, dirs, _ in os.walk(directory):
        for dir_name in dirs:
            subdirectories.append(os.path.join(root, dir_name))

    return subdirectories

def zip_folders(subfolder_path, output_folder, folder_names):
    for folder_name in folder_names:
        folder_path = os.path.join(subfolder_path, folder_name)
        if os.path.exists(folder_path):
            zip_name = os.path.join(output_folder, f'{os.path.basename(subfolder_path)}_{folder_name}.zip')
            with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, os.path.relpath(file_path, folder_path))
            print(f"Folder '{folder_path}' has been zipped to '{zip_name}'.")
        else:
            print(f"Folder '{folder_path}' does not exist. Skipping...")

def send_email(sender_email, sender_password, receiver_email, subject, message, attachments=[]):
    # Set up the SMTP server
    smtp_server = "smtp-mail.outlook.com"
    smtp_port = 587  # For SSL, use port 465

    # Create a MIMEText object with the message
    email_message = MIMEMultipart()
    email_message['From'] = sender_email
    email_message['To'] = receiver_email
    email_message['Subject'] = subject
    email_message.attach(MIMEText(message, 'plain'))

    # Attach files
    for attachment in attachments:
        with open(attachment, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {attachment.split('/')[-1]}")
        email_message.attach(part)

    # Start the SMTP session
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        # Login to the SMTP server
        server.login(sender_email, sender_password)
        # Send the email
        server.sendmail(sender_email, receiver_email, email_message.as_string())

subdirectories = nome_pastas()

zip_files = []
for item in subdirectories:
    caminho_subpastas = os.path.join(item, '2024', '01 - Janeiro')
    if os.path.exists(caminho_subpastas):
        zip_folders(caminho_subpastas, output_folder, ['Contabilidade', 'Financeiro', 'Fiscal'])
        for folder in ['Contabilidade', 'Financeiro', 'Fiscal']:
            zip_files.append(os.path.join(output_folder, f'01 - Janeiro_{folder}.zip'))
    else:
        print(f"Subfolder '{caminho_subpastas}' does not exist. Skipping...")

# Send email after all folders have been zipped
sender_email = "vitor.moreira@trinusco.com.br"
sender_password = "B2luzinho"
receiver_email = "vitor.moreira@trinusco.com.br"
subject = "Teste Email"
message = "teste."
send_email(sender_email, sender_password, receiver_email, subject, message, zip_files)
