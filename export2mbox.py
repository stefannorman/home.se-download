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


def get_message_ids(cookies, page=1):
    # max messages to get is 10,000
    msgno = 10000
    res = requests.get(
        "{}/mail/ms_ajax.asp?folder=/{}&pg={}&msgno={}&sortby=Received&sort_order=DESC&dtTS=full".format(
            SERVER_URL, FOLDER_NAME, page, msgno),
        cookies=cookies,
    )
    msg_ids = []
    for msg in res.text.split('{'):
        message_id = msg[:msg.find('}')]
        if message_id:
            msg_ids.append(msg[:msg.find('}')])
    if len(msg_ids) >= msgno:
        # call recurseviley if more than 10,000
        msg_ids += get_message_ids(cookies, page + 1)
    return msg_ids


cookies = login()
# message_ids = get_message_ids(cookies)
message_ids = ['597C0A55-D9F9-4DAD-B1A7-E8171E883CE0', '0FF29262-A6BD-4507-97DA-750B1C4E7B42']

mbox = mailbox.mbox('{}.mbox'.format(FOLDER_NAME))

mbox.lock()

for msg_ID in message_ids:

    r = requests.get(
        "{}/mail/ms_message.asp?MsgID={}&SM=F&FolderName=/{}".format(
            SERVER_URL,
            '{' + msg_ID + '}',
            FOLDER_NAME),
        cookies=cookies,
    )
    soup = BeautifulSoup(r.content, 'html.parser')
    from_tag = soup.find(id='FromA')
    from_addr = from_tag['title']
    from_name = from_tag.text

    subject_tag = soup.find(id='SubjectA')

    date_tag = soup.find(id='sDateA')

    msg_date = datetime.strptime(date_tag.text, '%d/%m/%Y %H:%M:%S %p')
    # Weird datetime bug when
    msg_date = msg_date + timedelta(hours=8, minutes=59)

    msg_body = soup.find(id='QComposerMSB').decode_contents()
    # Remove mysterious ? at the start of body
    if msg_body.startswith('?'):
        msg_body = msg_body[1:]

    # Attachments are listed in JavaScript in this format:
    # m_aCAtt[0] =  new CATTACH(0, '1', false, 'image001.png', '44', false, 'location = \'/tools/getFile.asp?GUID=40307d26-8ea5-4881-b67f-54e10cb37617&MsgID=%7B4FD62DD5-CB96-4DD3-9542-6E52604C2F39%7D&name=X*1\'', '/FileCabinet/images/icoIMG.gif');
    # m_aCAtt[1] =  new CATTACH(1, '2', true, 'image003.jpg', '3', false, 'location = \'/tools/getFile.asp?GUID=40307d26-8ea5-4881-b67f-54e10cb37617&MsgID=%7B4FD62DD5-CB96-4DD3-9542-6E52604C2F39%7D&name=X*2\'', '/FileCabinet/images/icoIMG.gif');
    # m_aCAtt[2] =  new CATTACH(2, '3', false, 'God Jul och ett Gott Nytt &#197;r_h&#228;lsning.pdf', '201', false, 'location = \'/tools/getFile.asp?GUID=40307d26-8ea5-4881-b67f-54e10cb37617&MsgID=%7B4FD62DD5-CB96-4DD3-9542-6E52604C2F39%7D&name=X*3\'', '/FileCabinet/images/icoPDF.gif');

    attachments = []
    for script_tag in soup.findAll('script', {'language': 'Javascript'}):
        if 'CATTACH' in script_tag.decode_contents():
            for line in script_tag.decode_contents().splitlines():
                line = line.strip().replace('\\', '').replace('\'', '')
                if line.startswith('m_aCAtt'):
                    parts = line.split(',')
                    attachment = {
                        'name': parts[3].strip(),
                        'url': parts[6].replace('location = ', '').strip(),
                    }
                    # stefan.norman@home.se >> _stefan.norman_home_se/
                    u_parts = USERNAME.split('@')
                    attach_folder = '_{}_{}/'.format(u_parts[0], u_parts[1].replace('.', '_'))
                    # Change url in body replace folder in img src with a content ID (cid)
                    msg_body = re.sub(r'/Attach/[0-9, A-Z, /, -]+', '', msg_body).replace(
                        attach_folder, 'cid:')

                    attachments.append(attachment)

    try:

        msg = MIMEMultipart()
        msg['From'] = '{} <{}>'.format(from_name, from_addr)
        msg['To'] = USERNAME
        msg['Date'] = msg_date.strftime("%a, %d %b %Y %H:%M:%S +0100")
        msg['Subject'] = subject_tag.text

        body = mailbox.mboxMessage()
        body.set_type('text/html')
        body.set_payload(msg_body, charset='utf-8')
        msg.attach(MIMEMessage(body))

        # add attachments
        for attachment in attachments:
            url = '{}{}'.format(SERVER_URL, attachment['url'])
            data = requests.get(url).content
            if attachment['name'].lower().endswith(('.jpg', '.jpeg', '.gif', '.png')):
                part = MIMEImage(data, name=attachment['name'])
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
        print('Error adding message {}'.format(msg_ID))
        print(e)

    finally:
        mbox.unlock()
