import sys
import imaplib
import getpass
import email
import email.header
import datetime

EMAIL_ACCOUNT = "kevan_tan@mymail.sutd.edu.sg"
EMAIL_FOLDER = "INBOX"

def process_mailbox(M):
    """
    Do something with emails messages in the folder.  
    For the sake of this example, print some headers.
    """

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    print(data)
    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        subject = str(hdr)
        print('Message %s: %s' % (num, subject))
        print('Raw Date:', msg['Date'])
        # Now convert to local date-time
        date_tuple = email.utils.parsedate_tz(msg['Date'])
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
            print ("Local Date:", \
                local_date.strftime("%a, %d %b %Y %H:%M:%S"))

M = imaplib.IMAP4_SSL('imap-mail.outlook.com')

try:
    rv, data = M.login(EMAIL_ACCOUNT, getpass.getpass())
    print(rv, data)
except imaplib.IMAP4.error:
    print("LOGIN FAILED")
    sys.exit(1)

rv, mailboxes = M.list()
if rv == 'OK':
    print("Mailboxes:")
    print(mailboxes)

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("processing mailbox...\n")
    process_mailbox(M)
    M.close()
else:
    print("unable to open mailbox ", rv)

M.logout()
