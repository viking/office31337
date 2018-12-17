office31337
-----------

office31337 is a Python 3 package for fetching e-mail from Office 365 and
storing it in a local mailbox (MH or mbox formats). It is designed to simulate
POP3 for those whose companies use Office 365 for e-mail and have disabled IMAP
and POP3 for vague 'security' reasons. MH and mbox formats can be consumed by
various e-mail clients, notably [Claws Mail](https://www.claws-mail.org/).

### Installation

Install office31337 by running the following shell command from inside the
package source directory:

```
python setup.py install
```

Add the `--user` argument to the above command if you wish to install the
package for your current user only.

### Usage

office31337 will try to use a local keyring (like GNOME keyring, for example)
for credential storage. The first time you run office31337, you should use the
`--password` argument to specify your password. If you don't want your password
to show up in your shell history, execute the command with a space in front of
it, like so:

```
$  office31337 --password hunter2 AzureDiamond@example.com ~/Mail/inbox
```

office31337 will store your password in the keyring for future usage, so you
only have to do this once.

Below is a full list of command line arguments accepted by office31337:

```
usage: office31337 [-h] [--password PASSWORD] [--limit LIMIT] [--verbose]
                   [--pretend] [--no-mark] [--all] [--mbox] [--no-dupes]
                   username mailbox

Fetch e-mails from Office 365 inbox into a local mailbox.

positional arguments:
  username             Account e-mail address
  mailbox              Local mailbox path

optional arguments:
  -h, --help           show this help message and exit
  --password PASSWORD  Password for account (optional, uses keyring otherwise)
  --limit LIMIT        Maximum number of e-mails to fetch
  --verbose            Print lots of messages
  --pretend            Don't write files or mark messages as read
  --no-mark            Don't mark messages as read
  --all                Fetch all e-mails instead of only unread e-mails
  --mbox               Use mbox format instead of MH
  --no-dupes           Make sure no duplicate e-mails are stored
```

### Notes

Python currently has a bug (likely related to
https://bugs.python.org/issue34424) that makes processing e-mail headers with
unicode characters impossible. To work around this, office31337 transliterates
unicode characters in headers via the
[Unidecode](https://pypi.org/project/Unidecode/) package.
