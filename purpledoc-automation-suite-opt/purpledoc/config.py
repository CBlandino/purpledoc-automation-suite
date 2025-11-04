import os
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
SMARTSHEET_TOKEN = os.getenv('SMARTSHEET_TOKEN')
SHEET_ID = int(os.getenv('SHEET_ID')) if os.getenv('SHEET_ID') else None
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.office365.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
TENANT_ID = os.getenv('TENANT_ID') or 'common'

# Local files
SMARTSHEET_CACHE_FILE = os.getenv('SMARTSHEET_CACHE_FILE', 'smartsheet_cache.json')
PROCESSED_FORM_TRACKER = os.getenv('PROCESSED_FORM_TRACKER', 'processed_form_rows.json')
O365_TOKEN_FILE = os.getenv('O365_TOKEN_FILE', 'o365_token.txt')
PDF_TEMPLATE = os.getenv('PDF_TEMPLATE', '000000 - Template.pdf')
