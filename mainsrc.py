import email
import imaplib
import datetime
import os
import email.header
from email.mime.text import MIMEText
import exceptions
import psycopg2
import docx
import smtplib
from email.mime.multipart import MIMEMultipart
import re

os.chdir('C:\\Users\\nemes\\OneDrive\\Desktop\\TestDirectory')
detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

searched_key_words = ['Prenume:', 'Nume:', 'Telefon:', 'E-mail:', 'Sectia:', 'Profil:', 'Media semestrul curent:',
                      'Media semestrul precedent (daca este cazul):']


# Read file and extract necessary information that follows key words from list, remove whitespace
def get_info_from_file(filename):
    doc = docx.Document(filename)
    extracted_info = []
    for para in doc.paragraphs:
        for key_word in searched_key_words:
            if key_word in para.text:
                extracted_info.append(para.text.split(key_word, 1)[1].strip())

    return extracted_info


def insert_info_into_db(extracted_info):
    # connect to db
    conn = psycopg2.connect(f'dbname=IngineriaSistemelor user=postgres password={os.environ["SQLPASS"]}')
    cursor = conn.cursor()
    # remove null values
    extracted_info = list(filter(None, extracted_info))
    # insert data without previous_gpa in case of students that are in the first semester
    if len(extracted_info) == 7:
        sql_query = f"INSERT INTO student_scholarships(first_name, last_name, phone_number, email, segment, profile, " \
                f"current_gpa) VALUES (%s, %s, %s, %s, %s, %s, %s)"
        cursor.execute(sql_query, extracted_info)
        conn.commit()
    # insert data with previous_gpa in case of students that are in semesters > 1
    elif len(extracted_info) == 8:
        sql_query = f"INSERT INTO student_scholarships(first_name, last_name, phone_number, email, segment, profile, current_gpa, previous_gpa) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
        cursor.execute(sql_query, extracted_info)
        conn.commit()
    else:
        print("Incorrect number of fields, the number of completed fields must be 7 or 8")
        raise exceptions.MissingValues


def process_info(filePath):
    fileName, file_extension = os.path.splitext(filePath)
    if file_extension == '.docx' or file_extension == '.doc':
        necessary_info = get_info_from_file(filePath)
        insert_info_into_db(necessary_info)
    else:
        print("The only accepted types of documents are .docx and .doc")
        raise exceptions.WrongTypeOfDocument


def send_mail(to_address, subject_mail, body_mail):
    gmail_user = os.environ['INGINERIE_MAIL_USER']
    gmail_password = os.environ['INGINERIE_MAIL_PASS']

    sent_from = gmail_user
    # Create mail content
    msg = MIMEMultipart()
    msg['From'] = 'sender_address'
    msg['To'] = to_address
    msg['Subject'] = subject_mail
    body_text = MIMEText(body_mail, 'plain')
    msg.attach(body_text)
    # Try to send mail. In case of error, mail the administrator
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to_address, msg.as_string())
        server.close()

        print('Email sent!')
    except:
        gmail_user = os.environ['INGINERIE_MAIL_USER']
        gmail_password = os.environ['INGINERIE_MAIL_PASS']

        sent_from = gmail_user
        to = "nemes.dan123@gmail.com"
        subject = "Error sending mail to " + to_address
        body = "Something went wrong. Please check."

        msg_except = MIMEMultipart()
        msg_except['From'] = gmail_user
        msg_except['To'] = to
        msg_except['Subject'] = subject
        body_text = MIMEText(body, 'plain')
        msg_except.attach(body_text)

        try:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            server.login(gmail_user, gmail_password)
            server.sendmail(sent_from, to, msg_except.as_string())
            server.close()
            print('Could not send mail to sender. Administrator notified.')
        except:
            print('Email could not be sent')


def main():
    try:
        userName = os.environ['INGINERIE_MAIL_USER']
        passwd = os.environ['INGINERIE_MAIL_PASS']

        # Login to GMAIL
        imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
        typ, accountDetails = imapSession.login(userName, passwd)

        if typ != 'OK':
            print("Cannot login to gmail")
            raise ConnectionError
        # Search All Mail folder
        imapSession.select('"[Gmail]/All Mail"')
        typ, data = imapSession.search(None, 'UNSEEN', 'ALL')
        if typ != 'OK':
            print('Error searching Inbox.')
            raise exceptions.ErrorSearchingMail

        # Iterate over all emails
        for msgId in data[0].split():
            typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
            if typ != 'OK':
                print('Error fetching mail.')
                raise exceptions.ErrorFetchingMail

            emailBody = messageParts[0][1]
            mail = email.message_from_bytes(emailBody)

            # Get and format the date when the mail was received in order to rename the file
            mail_date = mail['Date']
            mail_date_formatted = mail_date[5:25].replace(":", '.')
            num_of_attachments = 0

            # Check mail content
            for part in mail.walk():
                fileName = part.get_filename()

                # Check file attached and save it locally renamed as the date it was received + the sender of the mail
                if bool(fileName):
                    num_of_attachments += 1
                    mail_sender_regex = re.search("<(.*)>", mail['From'])
                    mail_sender = mail_sender_regex.group(1)
                    file_extension = os.path.splitext(fileName)[1]
                    filePath = os.path.join(detach_dir, 'attachments', mail_date_formatted + " (" + mail_sender + ")" + file_extension)
                    if not os.path.isfile(filePath):
                        print(fileName)
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                    try:
                        process_info(filePath)
                        send_mail(mail['From'], 'Cererea pentru bursa a fost primita',
                             'Datele au fost introduse in baza noastra de date. Multumim')
                    except exceptions.MissingValues:
                        send_mail(mail['From'], 'Campuri necompletate in cadrul cererii trimise',
                                 'Buna ziua, \nExista campuri necompletate in cadrul cererii trimise, cele 8 campuri trebuie completate'
                                 'cu exceptia ultimului camp care poate fi optional')
                    except exceptions.WrongTypeOfDocument:
                        send_mail(mail['From'], 'Documentul trimis nu este in format doc. sau docx.', 'Va rugam trimiteti modelul de pe site')

            if num_of_attachments == 0:
                send_mail(mail['From'], 'Nu ati atasat nici un document in email-ul trimis',
                         'Va rugam atasati modelul de cerere de bursa si trimiteti din nou un email')

        imapSession.close()
        imapSession.logout()
    except ConnectionError:
        send_mail('nemes.dan123@gmail.com', 'Could not login to gmail program executed at %s' % datetime.datetime.today(), 'Login Error')
    except exceptions.ErrorSearchingMail:
        send_mail('nemes.dan123@gmail.com', 'Could not search mail folder, program executed on %s' % datetime.datetime.today(),
                 'Error searching mail')
    except exceptions.ErrorFetchingMail:
        send_mail('nemes.dan123@gmail.com', 'Could not fetch any mail from folder, program executed on %s' % datetime.datetime.today(),
                 'Error fetching mail')


main()


