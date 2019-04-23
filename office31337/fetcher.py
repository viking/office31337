from exchangelib import Credentials, Account, Configuration, Message, FileAttachment, ItemAttachment, DELEGATE
from mailbox import MH, mbox
from email.message import EmailMessage
from unidecode import unidecode
from enum import Enum, auto
from sys import stdout

class FetchType(Enum):
    UNREAD = auto()
    ALL = auto()

class MailboxType(Enum):
    MH = auto()
    MBOX = auto()

class Fetcher:
    def __init__(self, username, password, mailbox_path, mailbox_type = MailboxType.MH, verbose = False):
        self.verbose = verbose
        if verbose:
            stdout.write("Logging in...")
            stdout.flush()

        credentials = Credentials(username, password)
        config = Configuration(server = 'outlook.office365.com',
                credentials = credentials)
        self.account = Account(primary_smtp_address = username,
                config = config, autodiscover = False, access_type = DELEGATE)

        if verbose:
            print("done.")

        self.mailbox_type = mailbox_type
        if mailbox_type == MailboxType.MH:
            self.mailbox = MH(mailbox_path)
        elif mailbox_type == MailboxType.MBOX:
            self.mailbox = mbox(mailbox_path)
        else:
            raise RuntimeError("invalid mailbox type")

    def fetch(self, which = FetchType.UNREAD, limit = None, mark_read = True, pretend = False, check_dupes = False):
        message_ids = []
        if check_dupes:
            if self.verbose:
                stdout.write("Collecting message IDs for duplication checks...")
                stdout.flush()

            message_ids = []
            for (i, m) in self.mailbox.items():
                m_id = m['Message-ID']
                if m_id is not None:
                    message_ids.append(m_id)

            if self.verbose:
                print("done.")

        it = self.account.inbox.all().order_by('-datetime_received')
        if which == FetchType.UNREAD:
            it = it.filter('IsRead:false')

        count = len(it)

        if limit is not None:
            if limit < count:
                count = limit
                it = it[:limit]

        for index, item in enumerate(it):
            if check_dupes and item.message_id is not None and item.message_id in message_ids:
                if self.verbose:
                    print(f"Skipping message {index+1} of {count}: {item.subject} ({item.message_id})")
                continue

            if self.verbose:
                print(f"Processing message {index+1} of {count}: {item.subject} ({item.message_id})")

            email = EmailMessage()

            # process headers
            if item.headers is None:
                if self.verbose:
                    print("> No headers! Skipping.")
                continue

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
                value = self._sanitize_header_value(value)

                email.add_header(name, value)

            if item.author is not None:
                email['From'] = self._format_mailbox(item.author)

            if item.sender is not None:
                email['Sender'] = self._format_mailbox(item.sender)

            if item.reply_to is not None:
                email['Reply-To'] = self._format_mailbox(item.reply_to)

            if item.to_recipients is not None:
                email['To'] = self._format_mailbox_list(item.to_recipients)

            if item.cc_recipients is not None:
                email['Cc'] = self._format_mailbox_list(item.cc_recipients)

            if item.bcc_recipients is not None:
                email['Bcc'] = self._format_mailbox_list(item.bcc_recipients)

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
                if item.text_body is not None:
                    email.add_related(item.text_body, subtype = 'plain')
                if item.body.body_type == 'HTML':
                    email.add_related(str(item.body), subtype = 'html')
            else:
                if item.text_body is not None:
                    email.set_content(item.text_body, subtype = 'plain')
                if item.body.body_type == 'HTML':
                    email.add_alternative(str(item.body), subtype = 'html')

            # add any inline attachments first
            self._add_attachments(email, inline_attachments, True)

            # add any non-inline attachments (will convert to multipart/mixed)
            self._add_attachments(email, attachments, False)

            if not pretend:
                self.mailbox.lock()
                self.mailbox.add(email)
                self.mailbox.unlock()

                if mark_read and not item.is_read:
                    item.is_read = True
                    item.save(update_fields=['is_read'])

                if check_dupes and email['Message-ID'] is not None:
                    message_ids.append(email['Message-ID'])

    def _sanitize_header_value(self, value):
        return unidecode(value)

    def _format_mailbox(self, mb):
        return self._sanitize_header_value('"%s" <%s>' % (mb.name, mb.email_address))

    def _format_mailbox_list(self, mb_list):
        return ', '.join(map(self._format_mailbox, mb_list))

    def _add_attachments(self, email, attachments, inline):
        for attachment in attachments:
            if isinstance(attachment, FileAttachment):
                maintype, subtype = attachment.content_type.split('/')
                with attachment.fp as fp:
                    data = fp.read()
                    if inline:
                        email.add_related(data, maintype = maintype, subtype = subtype,
                                          cid = attachment.content_id)
                    else:
                        email.add_attachment(data, maintype = maintype, subtype = subtype)
            elif isinstance(attachment, ItemAttachment):
                print(attachment)
