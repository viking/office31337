from email import policy
from email.message import EmailMessage
from exchangelib import FileAttachment, ItemAttachment
from unidecode import unidecode

class Message(EmailMessage):
    def __init__(self, policy = policy.default, item = None):
        EmailMessage.__init__(self, policy);

        if item is None:
            return

        # process headers
        for header in item.headers:
            name = header.name

            # deal with content type later on, since the email module gets
            # fussy when you do things out of order
            if name.lower() == 'content-type':
                continue

            # parse the header using the e-mail's policy class, since
            # sometimes unicode characters get created.
            # i.e. '=?UTF-8?Q?Fakult=c3=a4t_Statistik=2c_Technische_Universit?= =?UTF-8?Q?=c3=a4t_Dortmund?='
            value = self.policy.header_factory(name, header.value)

            # remove unicode characters to workaround python bug
            value = self._sanitize_header_value(value)

            self.add_header(name, value)

        if item.author is not None:
            self['From'] = self._format_mailbox(item.author)

        if item.sender is not None:
            self['Sender'] = self._format_mailbox(item.sender)

        if item.reply_to is not None:
            self['Reply-To'] = self._format_mailbox(item.reply_to)

        if item.to_recipients is not None:
            self['To'] = self._format_mailbox_list(item.to_recipients)

        if item.cc_recipients is not None:
            self['Cc'] = self._format_mailbox_list(item.cc_recipients)

        if item.bcc_recipients is not None:
            self['Bcc'] = self._format_mailbox_list(item.bcc_recipients)

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
                self.add_related(item.text_body, subtype = 'plain')
            if item.body.body_type == 'HTML':
                self.add_related(str(item.body), subtype = 'html')
        else:
            if item.text_body is not None:
                self.set_content(item.text_body, subtype = 'plain')
            if item.body.body_type == 'HTML':
                self.add_alternative(str(item.body), subtype = 'html')

        # add any inline attachments first
        self._add_attachments(inline_attachments, True)

        # add any non-inline attachments (will convert to multipart/mixed)
        self._add_attachments(attachments, False)

    def _sanitize_header_value(self, value):
        return unidecode(value)

    def _format_mailbox(self, mb):
        return self._sanitize_header_value('"%s" <%s>' % (mb.name, mb.email_address))

    def _format_mailbox_list(self, mb_list):
        return ', '.join(map(self._format_mailbox, mb_list))

    def _add_attachments(self, attachments, inline):
        for attachment in attachments:
            if isinstance(attachment, FileAttachment):
                maintype, subtype = attachment.content_type.split('/')
                with attachment.fp as fp:
                    data = fp.read()
                    if inline:
                        self.add_related(data, maintype = maintype, subtype = subtype,
                                         cid = attachment.content_id)
                    else:
                        self.add_attachment(data, maintype = maintype, subtype = subtype)
            elif isinstance(attachment, ItemAttachment):
                print(attachment)
