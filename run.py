from exchangelib import Credentials, Account, Configuration, DELEGATE
from mailbox import MH, MHMessage
from email.message import EmailMessage
from unidecode import unidecode
import keyring
import argparse

# python 3.6 seems to choke on unicode characters in headers sometimes
def sanitize_header_value(value):
    return unidecode(value)

def format_mailbox(mb):
    return sanitize_header_value('"%s" <%s>' % (mb.name, mb.email_address))

def format_mailbox_list(mb_list):
    return ', '.join(map(format_mailbox, mb_list))

parser = argparse.ArgumentParser(description = "Fetch e-mails from Office 365 inbox into a local mailbox.")
parser.add_argument('username', type = str, nargs = 1, help = "Account e-mail address")
parser.add_argument('mailbox', type = str, nargs = 1, help = "Local mailbox path")
parser.add_argument('--limit', type = int, nargs = 1, help = "Maximum number of e-mails to fetch")
parser.add_argument('--verbose', action = "store_true", help = "Print lots of messages")
parser.add_argument('--pretend', action = "store_true", help = "Don't write files or mark messages as read")
parser.add_argument('--no-mark', action = "store_true", help = "Don't mark messages as read")
parser.add_argument('--all', action = "store_true", help = "Fetch all e-mails instead of only unread e-mails")
args = parser.parse_args()

username = args.username[0]
password = keyring.get_password("exchange", username)
credentials = Credentials(username, password)
config = Configuration(server = 'outlook.office365.com', credentials = credentials)
account = Account(primary_smtp_address = username, config = config,
                  autodiscover = False, access_type = DELEGATE)
mh = MH(args.mailbox[0])

it = account.inbox.all().order_by('-datetime_received')
if not args.all:
    it = it.filter('IsRead:false')

count = len(it)

if args.limit is not None:
    limit = args.limit[0]
    if limit < count:
        count = limit
        it = it[:limit]

for index, item in enumerate(it):
    if args.verbose:
        print(f"Fetching message {index+1} of {count}: {item.message_id}")

    email = EmailMessage()
    for header in item.headers:
        name = header.name

        # deal with content type later on, since the email module gets fussy
        # when you do things out of order
        if name.lower() == 'content-type':
            continue

        value = sanitize_header_value(header.value)
        email.add_header(name, value)

    if item.author is not None:
        email['From'] = format_mailbox(item.author)

    if item.sender is not None:
        email['Sender'] = format_mailbox(item.sender)

    if item.reply_to is not None:
        email['Reply-To'] = format_mailbox(item.reply_to)

    if item.to_recipients is not None:
        email['To'] = format_mailbox_list(item.to_recipients)

    if item.cc_recipients is not None:
        email['Cc'] = format_mailbox_list(item.cc_recipients)

    if item.bcc_recipients is not None:
        email['Bcc'] = format_mailbox_list(item.bcc_recipients)

    # separate inline attachments
    inline_attachments = []
    attachments = []
    if item.attachments is not None:
        for attachment in item.attachments:
            if attachment.is_inline:
                inline_attachments.append(attachment)
            else:
                attachments.append(attachment)

    if len(inline_attachments) > 0:
        email.add_related(item.text_body, subtype = 'plain')
        if item.body.body_type == 'HTML':
            email.add_related(str(item.body), subtype = 'html')
    else:
        email.set_content(item.text_body, subtype = 'plain')
        if item.body.body_type == 'HTML':
            email.add_alternative(str(item.body), subtype = 'html')

    # add any inline attachments first
    for attachment in inline_attachments:
        maintype, subtype = attachment.content_type.split('/')
        with attachment.fp as fp:
            data = fp.read()
            email.add_related(data, maintype = maintype, subtype = subtype,
                              cid = attachment.content_id)

    # add any non-inline attachments (will convert to multipart/mixed)
    for attachment in attachments:
        maintype, subtype = attachment.content_type.split('/')
        with attachment.fp as fp:
            data = fp.read()
            email.add_attachment(data, maintype = maintype, subtype = subtype)

    if not args.pretend:
        mh.add(MHMessage(email))

        if not args.no_mark:
            item.is_read = True
            item.save()
