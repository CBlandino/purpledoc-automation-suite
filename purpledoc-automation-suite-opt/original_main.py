import os
import re
import json
from dotenv import load_dotenv
from O365 import Account
from O365.utils import FileSystemTokenBackend
import smartsheet
from rapidfuzz import fuzz
from datetime import datetime
from collections import defaultdict
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfString, PdfObject

# --------------------- Load Environment Variables ---------------------
load_dotenv()

CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
SMARTSHEET_TOKEN = os.getenv('SMARTSHEET_TOKEN')
SHEET_ID = int(os.getenv('SHEET_ID'))
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
TENANT_ID = os.getenv('TENANT_ID') or 'common'

# --------------------- Authenticate with Microsoft Graph ---------------------
credentials = (CLIENT_ID,)
token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
account = Account(credentials, auth_flow_type='public', tenant_id=TENANT_ID, token_backend=token_backend)

if not account.is_authenticated:
    account.authenticate(
        scopes = [
            'offline_access',
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Mail.ReadWrite',
            'https://graph.microsoft.com/Mail.Send',
            'https://graph.microsoft.com/Files.Read.All',
            'https://graph.microsoft.com/Sites.Read.All',
            ]
    )

mailbox = account.mailbox()
inbox = mailbox.inbox_folder()
messages = inbox.get_messages(limit=10)


# --------------------- Smartsheet Caching ---------------------

SMARTSHEET_CACHE_FILE = "smartsheet_cache.json"

def fetch_smartsheet_conversations(ss_client, sheet_id, row_ids):
    """
    Fetch comments (conversations) for each row in the sheet.
    Returns a dict: {row_id: [comment_dict, ...], ...}
    """
    conversations = {}
    for row_id in row_ids:
        try:
            comments = ss_client.Sheets.list_row_comments(sheet_id, row_id).data
            conversations[str(row_id)] = [
                {
                    "id": c.id,
                    "text": c.text,
                    "created_by": getattr(c.created_by, "email", ""),
                    "created_at": c.created_at.isoformat() if hasattr(c, "created_at") else ""
                }
                for c in comments
            ]
        except Exception as e:
            conversations[str(row_id)] = []
    return conversations

def load_smartsheet_cache():
    if os.path.exists(SMARTSHEET_CACHE_FILE):
        try:
            with open(SMARTSHEET_CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return (
                    data.get("columns", []),
                    data.get("rows", []),
                    data.get("timestamp", 0),
                    data.get("conversations", {})
                )
        except (json.JSONDecodeError, OSError, TypeError, ValueError):
            return [], [], 0, {}
    return [], [], 0, {}

def save_smartsheet_cache(columns, rows, conversations):
    import time
    with open(SMARTSHEET_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "columns": columns,
            "rows": rows,
            "conversations": conversations,
            "timestamp": int(time.time())
        }, f, ensure_ascii=False, indent=2)

def fetch_smartsheet_data_with_conversations():
    ss_client = smartsheet.Smartsheet(SMARTSHEET_TOKEN)
    sheet = ss_client.Sheets.get_sheet(SHEET_ID)
    columns = [
        {"id": col.id, "title": col.title.strip().lower()} for col in sheet.columns
    ]
    rows = []
    row_ids = []
    for row in sheet.rows:
        row_dict = {sheet.columns[i].title.lower(): cell.value for i, cell in enumerate(row.cells)}
        row_dict["_row_id"] = row.id
        rows.append(row_dict)
        row_ids.append(row.id)
    conversations = fetch_smartsheet_conversations(ss_client, SHEET_ID, row_ids)
    return columns, rows, conversations


# --------------------- Initialize Smartsheet Data ---------------------
columns, rows, _, conversations = load_smartsheet_cache()
if not rows:
    print("⚠️ Smartsheet cache empty. Fetching live data...")
    columns, rows, conversations = fetch_smartsheet_data_with_conversations()
    save_smartsheet_cache(columns, rows, conversations)

def html_to_clean_text(html):
    import re
    # Add newlines after block tags for better formatting
    html = re.sub(r'(?i)<br\s*/?>', '\n', html)
    html = re.sub(r'(?i)</p\s*>', '\n', html)
    html = re.sub(r'(?i)</div\s*>', '\n', html)
    html = re.sub(r'(?i)</li\s*>', '\n', html)
    html = re.sub(r'(?i)</tr\s*>', '\n', html)
    html = re.sub(r'(?i)<(p|div|li|tr|table)[^>]*>', '\n', html)
    # Remove all other tags
    text = re.sub(r'<[^>]+>', '', html)
    # Normalize spaces and newlines
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n+', '\n', text)
    lines = [line.strip() for line in text.split('\n')]
    return '\n'.join([line for line in lines if line])

def get_clean_email_body(msg):
    # Prefer HTML body if available
    if msg.body and msg.body_type and msg.body_type.lower() == 'html':
        return html_to_clean_text(msg.body)
    elif msg.body and msg.body_type and msg.body_type.lower() == 'text':
        return msg.body
    
    # Fallback: check raw message content for HTML or Text body parts
    if hasattr(msg, '_raw_message') and msg._raw_message:
        body_content = msg._raw_message.get('body', {})
        content_type = body_content.get('contentType', '').lower()
        content = body_content.get('content', '')
        if content_type == 'html':
            return html_to_clean_text(content)
        elif content_type == 'text':
            return content
    
    return ''  # if no body found

def normalize_ticket(ticket_str):
    if not ticket_str:
        return ''
    ticket_str = str(ticket_str).strip().lower()
    if ticket_str.endswith('.0'):
        ticket_str = ticket_str[:-2]
    ticket_str = re.sub(r'\W+', '', ticket_str)
    return ticket_str

def get_ticket_data(ticket_number):
    normalized_target = normalize_ticket(ticket_number)
    for row in rows:
        ticket_val = row.get("ticket number", "")
        if normalize_ticket(ticket_val) == normalized_target:
            return row, row  # Return the dict for both values for compatibility
    return None, None

def strip_signature(body):
    signature_keywords = [
        r'^--\s*$',
        r'^thanks[\s,]*$', r'^regards[\s,]*$',
        r'^t:\s*\(?\d{3}\)?',
        r'^m:\s*\(?\d{3}\)?',
        r'@\w+\.\w+',
        r'www\.',
        r'chris(topher)? blandino',
        r'Providing innovative Engineered Solutions',
        r'Get Outlook for iOS',
        r'Get Outlook for Android',
        r'Powered by O365',
        r'Powered by Microsoft 365',
    ]
    lines = body.strip().splitlines()
    stripped = []
    for line in lines:
        if any(re.search(pat, line.strip(), re.IGNORECASE) for pat in signature_keywords):
            break
        stripped.append(line)
    return '\n'.join(stripped)

# --- Snippet to highlight key parts with comments and minor polish ---

def send_email(to_addr, subject, body, pdf_path):
    m = account.new_message()
    m.to.add(to_addr)
    m.subject = subject
    m.body = body
    m.attachments.add(pdf_path)
    m.send()

def fill_pdf(input_pdf_path, output_pdf_path, data_dict):
    template_pdf = PdfReader(input_pdf_path)
    annotations = template_pdf.pages[0]['/Annots']

    if annotations:
        for annotation in annotations:
            if annotation['/Subtype'] == '/Widget' and annotation.get('/T'):
                key = annotation['/T'][1:-1]
                if key in data_dict:
                    value = data_dict[key] or ''
                    annotation.update(PdfDict(V=PdfString.encode(value)))
                    annotation.update(PdfDict(AS=PdfName('Yes')))

        if template_pdf.Root.AcroForm:
            template_pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
        else:
            template_pdf.Root.update(PdfDict(AcroForm=PdfDict(NeedAppearances=PdfObject('true'))))

    PdfWriter().write(output_pdf_path, template_pdf)

def send_error_email(sender_email, ticket, message):
    subject = f"Issue Processing Ticket #{ticket or 'Unknown'}"
    body = (
        f"There was an issue processing your Purple Doc form:\n\n"
        f"{message}\n\n"
        f"Please review your email format and try again."
    )
    m = account.new_message()
    m.to.add(sender_email)
    m.subject = subject
    m.body = body
    m.send()

def get_clean_email_body(msg):
    # Prioritize HTML body and convert to clean text preserving line breaks
    if msg.body and msg.body_type and msg.body_type.lower() == 'html':
        return html_to_clean_text(msg.body)
    elif msg.body and msg.body_type and msg.body_type.lower() == 'text':
        return msg.body
    # Fallback to raw message body content
    if hasattr(msg, '_raw_message') and msg._raw_message:
        body_content = msg._raw_message.get('body', {})
        content_type = body_content.get('contentType', '').lower()
        content = body_content.get('content', '')
        if content_type == 'html':
            return html_to_clean_text(content)
        elif content_type == 'text':
            return content
    return ''

import re
from rapidfuzz import fuzz

def fuzzy_contains(text, keywords, threshold=80):
    return any(fuzz.partial_ratio(text.lower(), kw.lower()) >= threshold for kw in keywords)

import re
from collections import defaultdict
from rapidfuzz import fuzz

def parse_email_body(body, default_tech_name="Unknown"):
    # Normalize newlines and strip signatures
    body = re.sub(r'(?i)<br\s*/?>', '\n', body)
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = strip_signature(body)  # Assuming you have this implemented

    # Remove service/update prefixes
    body = re.sub(r'(?im)^(service|update|note)\s*[:\-]?\s*', '', body)
    lines = [line.strip() for line in body.strip().split('\n') if line.strip()]
    flattened = ' '.join(lines).lower()

    data = {
        'ticket': None,
        'time_spent': '',
        'tech_notes': '',
        'additional_notes': '',
        'techs': defaultdict(lambda: {'notes': '', 'time': ''}),
        'error': None
    }

    # ---- Ticket: fuzzy match or 6-digit fallback
    for line in lines:
        if fuzz.partial_ratio("ticket", line.lower()) > 80:
            match = re.search(r'\d{6}', line)
            if match:
                data['ticket'] = match.group(0)
                break
    if not data['ticket']:
        match = re.search(r'\b\d{6}\b', flattened)
        if match:
            data['ticket'] = match.group(0)

    # ---- Status
    is_closed = 'close' in flattened
    is_ongoing = 'ongoing' in flattened

    # ---- Group lines by @Tech
    current_tech = default_tech_name
    tech_sections = defaultdict(list)

    for line in lines:
        mention_match = re.match(r'@([A-Za-z\s.]+)', line)
        if mention_match:
            current_tech = mention_match.group(1).strip()
            continue
        tech_sections[current_tech].append(line)

    for tech, tech_lines in tech_sections.items():
        notes = []
        time_spent = ''

        for l in tech_lines:
            # Accept both 'Time 1.5' and just '1.5'
            time_match = re.search(r'\btime\s*[:\-]?\s*(\d+(?:\.\d+)?|\d{1,2}:\d{2})', l, re.IGNORECASE)
            if not time_match:
                loose_time = re.match(r'^\s*(\d+(?:\.\d+)?|\d{1,2}:\d{2})\s*$', l)
                if loose_time:
                    time_match = loose_time

            if time_match:
                raw = time_match.group(1)
                try:
                    if ':' in raw:
                        h, m = map(int, raw.split(':'))
                        time_spent = f"{round(h + m/60.0, 2):.2f}"
                    else:
                        time_spent = f"{float(raw):.2f}"
                except:
                    data['error'] = f"Invalid time format '{raw}' for tech {tech}"
            else:
                notes.append(l)

        data['techs'][tech]['notes'] = '\n'.join(notes).strip()
        data['techs'][tech]['time'] = time_spent

    # Fallback if no techs found
    if not data['techs']:
        data['techs'][default_tech_name]['notes'] = '\n'.join(lines)

    # ---- One-time status in additional_notes
    if is_closed:
        data['additional_notes'] = 'close'
    elif is_ongoing:
        data['additional_notes'] = 'ongoing'
    else:
        data['additional_notes'] = 'ongoing'

    return data

# Main message processing loop snippet:
for msg in messages:
    if msg.subject.strip().lower() == 'pd' and not msg.is_read:
        print(f"\U0001f4e8 Processing email from: {msg.sender.address}")

        body = get_clean_email_body(msg)
        print(f"DEBUG: Body after cleaning:\n{body}")

        parsed = parse_email_body(body)
        ticket_number = parsed['ticket']

        if parsed['error']:
            send_error_email(msg.sender.address, ticket_number, parsed['error'])
            msg.mark_as_read()
            continue

        if not ticket_number:
            send_error_email(msg.sender.address, None, "No ticket number was found in your message.")
            msg.mark_as_read()
            continue

        # Pull matching ticket data from Smartsheet
        row_data, row_obj = get_ticket_data(ticket_number)
        if not row_data:
            send_error_email(msg.sender.address, ticket_number, f"Ticket #{ticket_number} not found in Smartsheet.")
            msg.mark_as_read()
            continue

        # Compose PDF field mapping with Smartsheet data + parsed notes
        sender_name = msg.sender.address.split('@')[0].replace('.', ' ').title()
        sent_date = msg.received.strftime('%m/%d/%Y')

        field_map = {
            'SERVICE TICKET': ticket_number,
            'COMPANY': row_data.get('site', ''),
            'SITE NAME': row_data.get('site', ''),
            'REQUESTED BY': row_data.get('requestor', ''),
            'SITE ADDRESS': row_data.get('address', ''),
            'TICKET REQUESTRow1': row_data.get('problem', ''),
            'TECHRow1': sender_name,
            'TECHNICIAN NOTESRow1': parsed['tech_notes'],
            'ADDITIONAL NOTESRow1': parsed['additional_notes'],
            'HOURSRow1': parsed['time_spent'],
            'DATERow1': sent_date
        }

        site_name = row_data.get('site', 'UNKNOWN').strip().upper()
        clean_site = re.sub(r'[^\w\s\-]', '', site_name).replace(' ', '_')
        short_date = datetime.strptime(sent_date, '%m/%d/%Y').strftime('%m-%d-%y')
        pdf_filename = f"{ticket_number} - {clean_site} - {short_date} - PurpleDoc.pdf"

        print("✅ Filled PDF fields:", field_map)

        fill_pdf('000000 - Template.pdf', pdf_filename, field_map)

        send_email(
            to_addr=msg.sender.address,
            subject=f'Purple Doc Report for Ticket #{ticket_number}',
            body='Here is your pre-filled Purple Doc form. Please complete any remaining fields and return it as needed.',
            pdf_path=pdf_filename
        )

        msg.mark_as_read()
        print(f"✅ Replied to ticket #{ticket_number}")

# --------------------- Main Processing ---------------------

for msg in messages:
    if msg.subject.strip().lower() == 'pd' and not msg.is_read:
        print(f"\U0001f4e8 Processing email from: {msg.sender.address}")

        body = get_clean_email_body(msg)
        print(f"DEBUG: Body after signature strip and cleaning:\n{body}")

        parsed = parse_email_body(body)
        ticket_number = parsed['ticket']

        if parsed['error']:
            send_error_email(msg.sender.address, ticket_number, parsed['error'])
            msg.mark_as_read()
            continue

        if not ticket_number:
            send_error_email(msg.sender.address, None, "No ticket number was found in your message.")
            msg.mark_as_read()
            continue

        row_data, row_obj = get_ticket_data(ticket_number)
        if not row_data:
            send_error_email(msg.sender.address, ticket_number, f"Ticket #{ticket_number} not found in Smartsheet.")
            msg.mark_as_read()
            continue

        sender_name = msg.sender.address.split('@')[0].replace('.', ' ').title()
        sent_date = msg.received.strftime('%m/%d/%Y')

        field_map = {
            'SERVICE TICKET': ticket_number,
            'COMPANY': row_data.get('site', ''),
            'SITE NAME': row_data.get('site', ''),
            'REQUESTED BY': row_data.get('requestor', ''),
            'SITE ADDRESS': row_data.get('address', ''),
            'TICKET REQUESTRow1': row_data.get('problem', ''),
            'TECHRow1': sender_name,
            'TECHNICIAN NOTESRow1': parsed['tech_notes'],
            'ADDITIONAL NOTESRow1': parsed['additional_notes'],
            'HOURSRow1': parsed['time_spent'],
            'DATERow1': sent_date
        }

        site_name = row_data.get('site', 'UNKNOWN').strip().upper()
        clean_site = re.sub(r'[^\w\s\-]', '', site_name).replace(' ', '_')
        short_date = datetime.strptime(sent_date, '%m/%d/%Y').strftime('%m-%d-%y')
        pdf_filename = f"{ticket_number} - {clean_site} - {short_date} - PurpleDoc.pdf"

        print("✅ Filled PDF fields:", field_map)

        fill_pdf('000000 - Template.pdf', pdf_filename, field_map)

        send_email(
            to_addr=msg.sender.address,
            subject=f'Purple Doc Report for Ticket #{ticket_number}',
            body='Here is your pre-filled Purple Doc form. Please complete any remaining fields and return it as needed.',
            pdf_path=pdf_filename
        )

        msg.mark_as_read()
        print(f"✅ Replied to ticket #{ticket_number}")

# --------------------- Microsoft Forms Excel Integration ---------------------
import json

PROCESSED_FORM_TRACKER = "processed_form_rows.json"
if not os.path.exists(PROCESSED_FORM_TRACKER):
    with open(PROCESSED_FORM_TRACKER, "w") as f:
        json.dump([], f)

def load_processed_form_rows():
    if not os.path.exists(PROCESSED_FORM_TRACKER) or os.path.getsize(PROCESSED_FORM_TRACKER) == 0:
        return set()  # no processed rows yet
    with open(PROCESSED_FORM_TRACKER, "r") as f:
        try:
            return set(json.load(f))
        except json.JSONDecodeError:
            return set()

def save_processed_form_rows(row_ids):
    with open(PROCESSED_FORM_TRACKER, "w") as f:
        json.dump(sorted(list(row_ids)), f)


import requests

import json

import json
import requests
import urllib.parse

def get_excel_form_rows(filename="Purple Doc _Online Form.xlsx", worksheet_name="Sheet1"):
    # Load access token
    try:
        with open('o365_token.txt', 'r') as f:
            token_file = json.load(f)
        access_token_entry = next(iter(token_file["AccessToken"].values()))
        access_token = access_token_entry["secret"]
    except Exception as e:
        raise RuntimeError("❌ Failed to read or parse o365_token.txt or extract access token") from e

    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    # Your OneDrive business drive ID (replace with yours)
    drive_id = "b!vrOA81dFbk2x7LBmIfNVmP2_IbG3Lk9MjOyvtNnftIHQYuYcqVNjQ4Zt57_A_toG"

    # Get file metadata by filename at drive root
    encoded_filename = urllib.parse.quote(filename)
    file_metadata_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_filename}"
    res = requests.get(file_metadata_url, headers=headers)
    res.raise_for_status()
    file_metadata = res.json()
    file_id = file_metadata['id']

    # Get worksheets list
    worksheets_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/workbook/worksheets"
    res_ws = requests.get(worksheets_url, headers=headers)
    res_ws.raise_for_status()
    worksheets = res_ws.json().get('value', [])

    # Find worksheet by name, or fallback to first worksheet
    worksheet_names = [ws['name'] for ws in worksheets]
    if worksheet_name in worksheet_names:
        ws_name = worksheet_name
    elif worksheets:
        ws_name = worksheets[0]['name']
        print(f"⚠️ Worksheet '{worksheet_name}' not found. Using first worksheet '{ws_name}' instead.")
    else:
        raise RuntimeError("❌ No worksheets found in workbook")

    # Get used range in the worksheet
    encoded_ws_name = urllib.parse.quote(ws_name)
    used_range_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/workbook/worksheets('{encoded_ws_name}')/usedRange"
    res_range = requests.get(used_range_url, headers=headers)
    res_range.raise_for_status()
    data = res_range.json()

    values = data.get("values", [])
    if not values or len(values) < 2:
        return []

    # Use first row as headers, normalize them
    headers_row = [str(h).strip().lower() for h in values[0]]

    # Return list of dicts for each data row
    return [dict(zip(headers_row, row)) for row in values[1:] if any(row)]


# Example usage:
# rows = get_excel_form_rows()
# print(rows)

# PDF field mapping helper
def build_pdf_field_map(ticket_number, row_data, sender_name, sent_date, parsed=None):
    # parsed is from parse_email_body, can be None for form rows
    return {
        'SERVICE TICKET': ticket_number,
        'COMPANY': row_data.get('site', ''),
        'SITE NAME': row_data.get('site', ''),
        'REQUESTED BY': row_data.get('requestor', ''),
        'SITE ADDRESS': row_data.get('address', ''),
        'TICKET REQUESTRow1': row_data.get('problem', ''),
        'TECHRow1': sender_name,
        'TECHNICIAN NOTESRow1': parsed['tech_notes'] if parsed else '',
        'ADDITIONAL NOTESRow1': parsed['additional_notes'] if parsed else '',
        'HOURSRow1': parsed['time_spent'] if parsed else '',
        'DATERow1': sent_date
    }

def fill_from_form_row(form_row, smartsheet_data):
    ticket_number = str(form_row.get("ticket number", "")).strip()
    if not ticket_number:
        return None
    email = form_row.get("email", "")
    name = form_row.get("name", "")
    work_done = form_row.get("work done", "")
    hours = str(form_row.get("time spent", "")).strip()
    status = str(form_row.get("ticket status", "")).strip().lower()
    completion = str(form_row.get("completion time", "")).strip()
    other_techs_raw = form_row.get("additional tech names", "")
    other_times_raw = form_row.get("other techs time spent", "")
    try:
        sent_date = datetime.strptime(completion, "%m/%d/%Y %I:%M:%S %p").strftime('%m/%d/%Y')
        short_date = datetime.strptime(sent_date, '%m/%d/%Y').strftime('%m-%d-%y')
    except:
        sent_date = datetime.today().strftime('%m/%d/%Y')
        short_date = datetime.today().strftime('%m-%d-%y')
    site_raw = smartsheet_data.get("site", "")
    site_name = site_raw.strip().upper() if isinstance(site_raw, str) else ""
    clean_site = re.sub(r'[^\w\s\-]', '', site_name).replace(' ', '_') or "NO_SITE"
    # Use helper for main fields
    field_map = build_pdf_field_map(ticket_number, smartsheet_data, name, sent_date)
    # Add form-specific fields
    field_map['TECHNICIAN NOTESRow1'] = work_done
    field_map['HOURSRow1'] = hours
    field_map['ADDITIONAL NOTESRow1'] = status if status in ['ongoing', 'close', 'closed'] else 'ongoing'
    # Add secondary techs
    other_techs = [t.strip() for t in other_techs_raw.split(',') if t.strip()]
    other_times = [t.strip() for t in other_times_raw.split(',') if t.strip()]
    for i, tech in enumerate(other_techs):
        idx = i + 2  # because TECHRow1 is primary
        field_map[f'TECHRow{idx}'] = tech
        field_map[f'TECHNICIAN NOTESRow{idx}'] = ""
        field_map[f'HOURSRow{idx}'] = other_times[i] if i < len(other_times) else ""
    pdf_filename = f"{ticket_number} - {clean_site} - {short_date} - PurpleDoc.pdf"
    print("✅ [FORM] Filled PDF fields:", field_map)
    fill_pdf("000000 - Template.pdf", pdf_filename, field_map)
    send_email(
        to_addr=email,
        subject=f"Purple Doc Report for Ticket #{ticket_number}",
        body="Here is your pre-filled Purple Doc form. Please complete any remaining fields and return it as needed.",
        pdf_path=pdf_filename
    )
    return ticket_number

def process_new_form_rows():
    seen_ids = load_processed_form_rows()
    new_ids = set()
    form_rows = get_excel_form_rows()

    for row in form_rows:
        row_id = str(row.get("id", "")).strip()
        if not row_id or row_id in seen_ids:
            continue

        ticket_number = row.get("ticket number", "")
        if not ticket_number:
            continue

        smartsheet_data, _ = get_ticket_data(ticket_number)
        if not smartsheet_data:
            print(f"❌ [FORM] Ticket {ticket_number} not found in Smartsheet.")
            continue

        result = fill_from_form_row(row, smartsheet_data)
        if result:
            new_ids.add(row_id)

    seen_ids.update(new_ids)
    save_processed_form_rows(seen_ids)

# Run this to process any new form entries
process_new_form_rows()

import time

def check_for_new_emails():
    messages = inbox.get_messages(limit=10)
    for msg in messages:
        if msg.subject.strip().lower() == 'pd' and not msg.is_read:
            print(f"📨 Processing email from: {msg.sender.address}")
            body = get_clean_email_body(msg)
            parsed = parse_email_body(body)
            ticket_number = parsed['ticket']
            if parsed['error']:
                send_error_email(msg.sender.address, ticket_number, parsed['error'])
                msg.mark_as_read()
                continue
            if not ticket_number:
                send_error_email(msg.sender.address, None, "No ticket number was found in your message.")
                msg.mark_as_read()
                continue
            row_data, row_obj = get_ticket_data(ticket_number)
            if not row_data:
                send_error_email(msg.sender.address, ticket_number, f"Ticket #{ticket_number} not found in Smartsheet.")
                msg.mark_as_read()
                continue
            sender_name = msg.sender.address.split('@')[0].replace('.', ' ').title()
            sent_date = msg.received.strftime('%m/%d/%Y')
            short_date = datetime.strptime(sent_date, '%m/%d/%Y').strftime('%m-%d-%y')
            field_map = build_pdf_field_map(ticket_number, row_data, sender_name, sent_date, parsed)
            site_name = row_data.get('site', 'UNKNOWN').strip().upper()
            clean_site = re.sub(r'[^\w\s\-]', '', site_name).replace(' ', '_')
            pdf_filename = f"{ticket_number} - {clean_site} - {short_date} - PurpleDoc.pdf"
            print("✅ Filled PDF fields:", field_map)
            fill_pdf('000000 - Template.pdf', pdf_filename, field_map)
            send_email(
                to_addr=msg.sender.address,
                subject=f'Purple Doc Report for Ticket #{ticket_number}',
                body='Here is your pre-filled Purple Doc form. Please complete any remaining fields and return it as needed.',
                pdf_path=pdf_filename
            )
            msg.mark_as_read()
            print(f"✅ Replied to ticket #{ticket_number}")
            
def refresh_smartsheet_cache():
    global columns, rows, conversations
    try:
        latest_columns, latest_rows, latest_conversations = fetch_smartsheet_data_with_conversations()
        cached_row_ids = set(str(r.get("_row_id", "")) for r in rows)
        latest_row_ids = set(str(r.get("_row_id", "")) for r in latest_rows)

        new_rows = [r for r in latest_rows if str(r.get("_row_id", "")) not in cached_row_ids]

        if new_rows or latest_conversations != conversations:
            print(f"🆕 Found {len(new_rows)} new rows or updated conversations in Smartsheet. Updating cache...")
            merged_rows = rows + new_rows
            columns = latest_columns
            conversations = latest_conversations
            save_smartsheet_cache(columns, merged_rows, conversations)
            rows = merged_rows
        else:
            print("ℹ️ No new Smartsheet rows or conversations this cycle.")
    except Exception as e:
        print(f"⚠️ Error refreshing Smartsheet cache: {e}")
# --------------------- Continuous Loop ---------------------
if __name__ == "__main__":
    print("🔁 Starting continuous watch loop...")
    while True:
        try:
            
            print("🔍 Checking for new emails, form entries, and Smartsheet updates...")
            
            refresh_smartsheet_cache()
            check_for_new_emails()
            process_new_form_rows()
            print("✅ Idle. Waiting 30 seconds before next check.\n")
        except Exception as e:
            print(f"⚠️ Error occurred: {e}")

        time.sleep(30)

