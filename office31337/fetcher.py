from exchangelib import Credentials, Account, Configuration, DELEGATE
from mailbox import MH, MHMessage
from email.message import EmailMessage
from unidecode import unidecode
from enum import Enum, auto

# python 3.6 seems to choke on unicode characters in headers sometimes
def sanitize_header_value(value):
    return unidecode(value)

def format_mailbox(mb):
    return sanitize_header_value('"%s" <%s>' % (mb.name, mb.email_address))

def format_mailbox_list(mb_list):
    return ', '.join(map(format_mailbox, mb_list))

class FetchType(Enum):
    UNREAD = auto()
    ALL = auto()

class Fetcher:
    def __init__(self, username, password, destination):
        credentials = Credentials(username, password)
        config = Configuration(server = 'outlook.office365.com',
                credentials = credentials)
        self.account = Account(primary_smtp_address = username,
                config = config, autodiscover = False, access_type = DELEGATE)
        self.mh = MH(destination)

    def fetch(self, which = FetchType.UNREAD, limit = None, verbose = False, mark_read = True, pretend = False):
        it = self.account.inbox.all().order_by('-datetime_received')
        if which == FetchType.UNREAD:
            it = it.filter('IsRead:false')

        count = len(it)

        if limit is not None:
            if limit < count:
                count = limit
                it = it[:limit]

        for index, item in enumerate(it):
            if verbose:
                print(f"Fetching message {index+1} of {count}: {item.message_id}")

            email = EmailMessage()
            for header in item.headers:
                name = header.name

                # deal with content type later on, since the email module gets
                # fussy when you do things out of order
                if name.lower() == 'content-type':
                    continue

                # parse the header using the e-mail's policy class, since
                # sometimes unicode characters get created.
                # i.e. '=?UTF-8?Q?Fakult=c3=a4t_Statistik=2c_Technische_Universit?= =?UTF-8?Q?=c3=a4t_Dortmund?='
                value = email.policy.header_factory(name, header.value)

                # remove unicode characters to workaround python bug
                value = sanitize_header_value(value)

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

            if not pretend:
                self.mh.add(MHMessage(email))

                if mark_read and item.is_read:
                    item.is_read = True
                    item.save()
