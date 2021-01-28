import email
import getpass, imaplib
from datetime import date
from datetime import datetime
import os
import email.header
import sys
from email.mime.text import MIMEText

import exceptions
import smtplib
from email.mime.multipart import MIMEMultipart

def SendMail(to_address, subject_mail, body_mail):
    gmail_user = os.environ['INGINERIE_MAIL_USER']
    gmail_password = os.environ['INGINERIE_MAIL_PASS']

    sent_from = gmail_user

    msg = MIMEMultipart()
    msg['From'] = 'sender_address'
    msg['To'] = to_address
    msg['Subject'] = subject_mail
    body_text = MIMEText(body_mail, 'plain')
    msg.attach(body_text)

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
            print('Email sent!')
        except:
            print('Email could not be sent')


