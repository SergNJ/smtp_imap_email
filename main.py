import smtplib
import ssl
import string
import random
import sys
import yagmail
import imaplib
import email
import os
import webbrowser
from email import encoders
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import decode_header

con_method = 5
# 1 - SMTP Gmail SSL secure via smtplib
# 2 - SMTP Gmail unsecure via smtplib
# 3 - built-in python SMTP server
# 4 - SMTP Gmail via yagmail
# 5 - IMAP Gmail

action = 'receive'
# 'send'
# 'receive'

message_type = 'mime'
# 'plain'
# 'mime'

# section below is not safe. It is just an exercise.
gmail_sender_pwd = 'sender password'
gmail_sender_addr = 'sender@gmail.com'
gmail_rec_addr = 'receiver@gmail.com'
dummy_file = 'dummy payload.pdf'
pic_file = 'robot.jpg'

MIME_text = """\
Hi,
How are you?
"""

MIME_html = """\
<html>
  <body>
    <p>Hi,<br>
       How are you?<br>
       <img src="cid:image1" width="128" height="128"><b>Here is a picture of a robot for you.</b><br>
       <a href="www.google.com">This is a link to Google</a> 
       Click the link
    </p>
  </body>
</html>
"""


def clean(text):
    return "".join(c if c.isalnum() else "_" for c in text)


def rnd_subject():
    return 'test id ' + ''.join(random.sample(string.ascii_uppercase + string.digits, 6)) + '\n'


def rnd_plain_message():
    s = 'Subject: test id ' + ''.join(random.sample(string.ascii_uppercase + string.digits, 6)) + '\n'
    s += 'This is a test message - line 1 \n line 2 \n last line 3.'
    return s


def parse_list_response(line):
    flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
    mailbox_name = mailbox_name.strip('"')
    return (flags, delimiter, mailbox_name)


if con_method == 1:  # SMTP Gmail SSL secure via smtplib
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp_serv:
            smtp_serv.login(gmail_sender_addr, gmail_sender_pwd)
        print('SMTP Method 1 - Connection successful.')
    except Exception as e:
        print(e)
        sys.exit('SMTP Method 1 - Connection to SMTP server failed.')

elif con_method == 2:  # SMTP Gmail unsecure via smtplib
    context = ssl.create_default_context()
    try:
        smtp_serv = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_serv.ehlo()
        smtp_serv.starttls(context=context)
        smtp_serv.ehlo()
        smtp_serv.login(gmail_sender_addr, gmail_sender_pwd)
        print('SMTP Method 2 - Connection successful.')
    except Exception as e:
        print(e)
        sys.exit('SMTP Method 2 - Connection to SMTP server failed.')

elif con_method == 3:  # built-in python SMTP server
    # built-in smtp server for testing. Default output to the shell. python -m smtpd -n -c DebuggingServer
    # 127.0.0.1:1025 requires adding 127.0.0.1 smtp.test.com to hosts filr located in Windows 10 at
    # C:\Windows\System32\drivers\etc\hosts
    context = ssl.create_default_context()
    try:
        smtp_serv = smtplib.SMTP('smtp.test.com', 1025)
        smtp_serv.sendmail('me@dummy.com', 'they@dummy.com', 'Hey there!')
        sys.exit('SMTP Method 3 - success! Bye!')
    except Exception as e:
        print(e)
        sys.exit('SMTP Method 3 - Connection to SMTP server failed.')

elif con_method == 4:  # SMTP Gmail via yagmail
    try:
        yagcon = yagmail.SMTP(user=gmail_sender_addr, password=gmail_sender_pwd)
        yagcon.send(to=gmail_rec_addr, subject='Test message', contents='Hello there!', attachments=dummy_file)
        print('Test message was sent OK')
        sys.exit('SMTP Method 4 - success! Bye!')
    except Exception as e:
        print(e)
        sys.exit('SMTP Method 4 - failed.')

elif con_method == 5:  # IMAP Gmail
    try:
        imap_serv = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        imap_serv.login(gmail_sender_addr, gmail_sender_pwd)
        print('IMAP Method 1 - Connection successful.')
    except Exception as e:
        print(e)
        sys.exit('IMAP Method 1 - failed.')

if action == 'send':
    try:  # SMTP sending email
        if message_type == 'plain':
            print('Server status: ' + str(smtp_serv.noop()))
            smtp_serv.sendmail(gmail_sender_addr, gmail_rec_addr, rnd_plain_message())
            print('Plain message was sent OK')

        elif message_type == 'mime':
            print('Server status: ' + str(smtp_serv.noop()))
            MIME_message = MIMEMultipart('alternative')
            MIME_message['Subject'] = rnd_subject()
            MIME_message['From'] = gmail_sender_addr
            MIME_message['To'] = gmail_rec_addr
            MIME_message['Bcc'] = gmail_rec_addr
            part_text = MIMEText(MIME_text, 'plain')
            part_html = MIMEText(MIME_html, 'html')

            with open(pic_file, 'rb') as f:
                part_pic = MIMEImage(f.read())
            part_pic.add_header('Content-ID', '<image1>')

            with open(dummy_file, 'rb') as attachment:
                part_attachment = MIMEBase('application', 'octet-stream')
                part_attachment.set_payload(attachment.read())
            encoders.encode_base64(part_attachment)
            part_attachment.add_header('Content-Disposition', 'attachment', filename=dummy_file)

            MIME_message.attach(part_text)
            MIME_message.attach(part_html)
            MIME_message.attach(part_pic)
            MIME_message.attach(part_attachment)  # The email client will try to render the last part first
            smtp_serv.sendmail(gmail_sender_addr, gmail_rec_addr, MIME_message.as_string())
            print('MIME message was sent OK')
    except Exception as e:
        print(e)
    finally:
        smtp_serv.quit()

elif action == 'receive':
    status, messages = imap_serv.select('INBOX')
    msg_count = int(messages[0])

    print('Folder selection status: %s, Messages: %s' % (status, msg_count))
    for item in imap_serv.list()[1]:  # getting a list of folders on the server
        l = item.decode().split(' "/" ')
        print(' ' * 10 + l[0] + " = " + l[1])
    print()

    for i in range(msg_count, 0, -1):
        res, msg = imap_serv.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):

                msg = email.message_from_bytes(response[1])

                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)

                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                recd, encoding = decode_header(msg.get("Received"))[0]
                if isinstance(recd, bytes):
                    recd = recd.decode(encoding)

                print('Fetch: ' + res + ', Subject: %s, From: %s, Received: %s ' % (subject, From, recd))

                if msg.is_multipart():
                    print(' ' * 10 + 'Structure: Multipart')
                else:
                    print(' ' * 10 + 'Structure: Plain text')

                print('Available fields: ' + '| '.join(msg.keys()))

                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))

                        try:
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass

                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            print(body)

                        elif "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    content_type = msg.get_content_type()
                    body = msg.get_payload(decode=True).decode()

                    if content_type == "text/plain":
                        print(body)
                    if content_type == "text/html":
                        folder_name = clean(subject)
                        if not os.path.isdir(folder_name):
                            os.mkdir(folder_name)
                        filename = "index.html"
                        filepath = os.path.join(folder_name, filename)
                        open(filepath, "w").write(body)
                        webbrowser.open(filepath)
        print('=' * 50 + ' MESSAGE END ' + '=' * 50)
    print(imap_serv.close())
    print(imap_serv.logout())
