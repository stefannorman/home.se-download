import json

import requests
import os
from bs4 import BeautifulSoup
import re
import mailbox
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.message import MIMEMessage
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

LOGIN_URL = "http://idlogin.spray.se/home.se/mail"
USERNAME = os.getenv('HOME_SE_USERNAME')
PASSWORD = os.getenv('HOME_SE_PASSWORD')
FOLDER_NAME = 'Inbox'
SERVER_URL = "http://nymail.spray.se"


def login():
    # Login to home.se
    res = requests.post(
        LOGIN_URL,
        headers={
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        data={
            'username': USERNAME,
            'password': PASSWORD,
        })

    # Get cookies with session data
    return res.request._cookies


def get_messages(cookies, page=1):
    # max messages to get is 10,000
    msgno = 10000
    res = requests.get(
        "{}/mail/ms_ajax.asp?folder=/{}&pg={}&msgno={}&sortby=Received&sort_order=DESC&dtTS=full".format(
            SERVER_URL, FOLDER_NAME, page, msgno),
        cookies=cookies,
    )
    _messages = []
    msg_data = res.text.split('{')
    for msg in msg_data:
        m = msg.split('_#c|-')
        if len(m) > 2 and 'Spam!' not in m[0]:
            _messages.append({
                'id': m[0].replace('}', ''),
                'subject': m[1],
                'date': datetime.strptime(m[2], '%m/%d/%Y %H:%M:%S %p'),
                'from_name': m[3],
                'from_email': m[4],
            })
    if len(msg_data) >= msgno:
        # call recursively if more than 10,000
        _messages += get_messages(cookies, page + 1)
    return _messages


cookies = login()
messages = get_messages(cookies)

# DEBUG
# messages = messages[2700:]
debug_ids = []
# debug_ids = ['0D7DA972-1275-436C-8A99-8DD487C9FEE4']

filename = '{}.mbox'.format(FOLDER_NAME)
mbox = mailbox.mbox(filename)

print('Exporting {} to {}'.format(len(messages), filename))

mbox.lock()

for message in messages:

    if len(debug_ids) and message['id'] not in debug_ids:
        continue

    r = requests.get(
        "{}/mail/ms_message.asp?MsgID={}&SM=F&FolderName=/{}".format(
            SERVER_URL,
            '{' + message['id'] + '}',
            FOLDER_NAME),
        cookies=cookies,
    )
    soup = BeautifulSoup(r.content, 'html.parser')

    date_tag = soup.find(id='sDateA')

    # Weird datetime bug
    msg_date = message['date'] + timedelta(hours=8, minutes=59)

    msg_body = soup.find(id='QComposerMSB').decode_contents()
    # Remove mysterious ? at the start of body
    if msg_body.startswith('?'):
        msg_body = msg_body[1:]

    # Attachments are listed in JavaScript.
    attachments = []
    for script_tag in soup.findAll('script', {'language': 'Javascript'}):
        if 'CATTACH' in script_tag.decode_contents():
            for line in script_tag.decode_contents().splitlines():
                line = line.strip()
                # Attachment lines start with something like:
                # m_aCAtt[0] =  new CATTACH(
                if line.startswith('m_aCAtt'):

                    # Remove JS code and construct array
                    line = re.sub(r'^m_aCAtt\[\d+\] =  new CATTACH\(', '[',
                                             re.sub(r'\);$', ']', line))
                    # Fix string to be JSON friendly.
                    # Tricky, since url can contain single quote and comma
                    line = line.replace(", '", ", \"").replace("\',", "\",").replace("\']", "\"]")
                    # # Remove double backslash
                    line = line.strip().replace('\\', '')
                    # Load into JSON
                    line = json.loads(line)

                    attachment = {
                        'name': line[3].strip(),
                        'url': line[6].replace("location = ", "").replace("'", ""),
                    }
                    # Get attachment folder from username.
                    # I e stefan.norman@home.se >> _stefan.norman_home_se/
                    u_parts = USERNAME.split('@')
                    attach_folder = '_{}_{}/'.format(u_parts[0], u_parts[1].replace('.', '_'))
                    # Change url in body replace folder in img src with a content ID (cid)
                    msg_body = re.sub(r'/Attach/[0-9, A-Z, /, -]+', '', msg_body).replace(
                        attach_folder, 'cid:')

                    attachments.append(attachment)

    try:

        msg = MIMEMultipart()
        msg['From'] = '{} <{}>'.format(message['from_name'], message['from_email'])
        msg['To'] = USERNAME
        msg['Date'] = msg_date.strftime("%a, %d %b %Y %H:%M:%S +0100")
        msg['Subject'] = message['subject']

        body = mailbox.mboxMessage()
        body.set_type('text/html')
        body.set_payload(msg_body, charset='utf-8')
        msg.attach(MIMEMessage(body))

        # add attachments
        for attachment in attachments:
            _filename, _extension = os.path.splitext(attachment['name'].lower())
            if _extension == '.vcf':
                # Getting VCF contact info is broken, skip it.
                continue
            url = '{}{}'.format(SERVER_URL, attachment['url'])
            data = requests.get(url).content
            if attachment['name'].lower().endswith(('.jpg', '.jpeg', '.gif', '.png')):
                part = MIMEImage(data, name=attachment['name'], _subtype=_extension[1:])
                # Add CID to enable internal img src reference in html mail
                part.add_header('Content-ID', '<%s>' % attachment['name'])
            else:
                part = MIMEApplication(data, name=attachment['name'])

            part.add_header('Content-Disposition',
                            'attachment; filename="%s"' % attachment['name'])
            msg.attach(part)

        mbox.add(msg)
        mbox.flush()

    except TypeError as e:
        print('Error adding message {}'.format(message['id']))
        print(e)

    finally:
        mbox.unlock()
