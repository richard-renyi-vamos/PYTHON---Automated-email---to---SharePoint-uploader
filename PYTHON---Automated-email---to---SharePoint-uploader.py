import imaplib
import email
from email.header import decode_header
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File

# Email credentials and IMAP server
username = "your-email@example.com"
password = "your-password"
imap_server = "imap.gmail.com"

# SharePoint credentials
client_id = "your-client-id"
client_secret = "your-client-secret"
site_url = "https://yourdomain.sharepoint.com/sites/yoursite"
library_name = "Documents"

# Connect to email server
mail = imaplib.IMAP4_SSL(imap_server)
mail.login(username, password)
mail.select("inbox")

# Search for specific subject
subject_to_search = "Specific Subject"
status, messages = mail.search(None, '(SUBJECT "{}")'.format(subject_to_search))
messages = messages[0].
