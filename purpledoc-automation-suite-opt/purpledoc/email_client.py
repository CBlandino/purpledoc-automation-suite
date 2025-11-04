from O365 import Account
from O365.utils import FileSystemTokenBackend
from .config import CLIENT_ID, CLIENT_SECRET, TENANT_ID, O365_TOKEN_FILE, SMTP_SERVER, SMTP_PORT
from typing import List, Optional

def create_account():
    credentials = (CLIENT_ID, CLIENT_SECRET)
    token_backend = FileSystemTokenBackend(token_path='.', token_filename=O365_TOKEN_FILE)
    account = Account(credentials, auth_flow_type='public', tenant_id=TENANT_ID, token_backend=token_backend)
    if not account.is_authenticated:
        account.authenticate(scopes=[
            'offline_access',
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Mail.ReadWrite',
            'https://graph.microsoft.com/Mail.Send',
            'https://graph.microsoft.com/Files.Read.All',
            'https://graph.microsoft.com/Sites.Read.All',
        ])
    return account

class EmailClient:
    def __init__(self, account):
        self.account = account
        self.mailbox = account.mailbox()
        self.inbox = self.mailbox.inbox_folder()

    def fetch_unread_pd_messages(self, limit=10):
        messages = self.inbox.get_messages(limit=limit)
        return [m for m in messages if m.subject and m.subject.strip().lower() == 'pd' and not m.is_read]

    def send_message(self, to_addr: str, subject: str, body: str, attachments: Optional[List[str]] = None):
        m = self.account.new_message()
        m.to.add(to_addr)
        m.subject = subject
        m.body = body
        if attachments:
            for a in attachments:
                m.attachments.add(a)
        m.send()
