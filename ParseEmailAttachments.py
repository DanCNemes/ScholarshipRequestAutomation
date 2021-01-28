import email
import getpass, imaplib
from datetime import date
from datetime import datetime
import os
import email.header
import sys

os.chdir('C:\\Users\\nemes\\OneDrive\\Desktop\\TestDirectory')
detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

userName = os.environ['INGINERIE_MAIL_USER']
passwd = os.environ['INGINERIE_MAIL_PASS']

imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
typ, accountDetails = imapSession.login(userName, passwd)

if typ != 'OK':
    print('Not able to sign in!')
    raise ConnectionError

imapSession.select('"[Gmail]/All Mail"')
typ, data = imapSession.search(None, 'UNSEEN', 'ALL')
if typ != 'OK':
    print('Error searching Inbox.')
    raise ConnectionError

    # Iterating over all emails
for msgId in data[0].split():
    typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
    if typ != 'OK':
        print('Error fetching mail.')
        raise ConnectionError

    emailBody = messageParts[0][1]
    mail = email.message_from_bytes(emailBody)
    mail_date = mail['Date']
    mail_date_formatted = mail_date[5:16]
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            # print part.as_string()
            continue
        if part.get('Content-Disposition') is None:
            # print part.as_string()
            continue
        fileNameDate = date.today()
        fileName = part.get_filename()

        if bool(fileName):
            file_extension = os.path.splitext(fileName)[1]
            filePath = os.path.join(detach_dir, 'attachments', mail_date_formatted + file_extension)
            if not os.path.isfile(filePath):
                print(fileName)
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

imapSession.close()
imapSession.logout()
