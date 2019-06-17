from exchangelib import Credentials, Account, Configuration, DELEGATE
from mailbox import MH, mbox
from enum import Enum, auto
from sys import stdout
from .message import Message

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

            if item.headers is None:
                if self.verbose:
                    print("> No headers! Skipping.")
                continue

            email = Message(item=item)

            if not pretend:
                self.mailbox.lock()
                self.mailbox.add(email)
                self.mailbox.unlock()

                if mark_read and not item.is_read:
                    item.is_read = True
                    item.save(update_fields=['is_read'])

                if check_dupes and email['Message-ID'] is not None:
                    message_ids.append(email['Message-ID'])
