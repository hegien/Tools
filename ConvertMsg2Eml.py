# Script to convert all .msg files in a file structure to .eml format

# Following code was provided by Copilot.
# !!! check and adjust it before us !!!

import extract_msg
from email.message import EmailMessage

def msg_to_eml(msg_path, eml_path):
    msg = extract_msg.Message(msg_path)
    email_msg = EmailMessage()
    email_msg['Subject'] = msg.subject
    email_msg['From'] = msg.sender
    email_msg['To'] = ', '.join(msg.to)
    email_msg.set_content(msg.body)

    # Anhänge hinzufügen
    for attachment in msg.attachments:
        email_msg.add_attachment(attachment.data, maintype='application', subtype='octet-stream', filename=attachment.longFilename)

    # EML-Datei speichern
    with open(eml_path, 'wb') as f:
        f.write(email_msg.as_bytes())

# Beispielaufruf
msg_to_eml('beispiel.msg', 'beispiel.eml')
