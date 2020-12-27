# Export emails from home.se

My webmail, home.se (Spray), has been neglected far to long.
You can't change password, use IMAP, etc. 

Downloading your emails is also broken. Hence this python script that exports all emails in a home.se folder into an mbox file.

## Usage

- Set environment variables `HOME_SE_USERNAME` and `HOME_SE_PASSWORD` to your credentials.
- Create a virtual env, activate it and run `pip install -r requirements.txt`.
- Run `python export2mbox.py`.

Per default the _Inbox_ is scraped and an `Inbox.mbox` is created.
You can use this mbox file to import into Apple Mail or Mozilla Firebird.
Gmail does not have import capabilities from mbox, but by importing into the mentioned email clients, you are able to bypass that. 

If you want to export another folder alter the `FOLDER_NAME` variable.