import time
import os
from purpledoc.config import PDF_TEMPLATE, PROCESSED_FORM_TRACKER, O365_TOKEN_FILE
from purpledoc.smartsheet_client import load_smartsheet_cache, save_smartsheet_cache, fetch_smartsheet_data_with_conversations, get_ticket_by_number
from purpledoc.email_client import create_account, EmailClient
from purpledoc.parser import get_clean_email_body, parse_email_body, strip_signature
from purpledoc.pdf_util import fill_pdf
from purpledoc.forms import get_excel_form_rows
from purpledoc.config import SMARTSHEET_CACHE_FILE

def ensure_processed_tracker():
    if not os.path.exists(PROCESSED_FORM_TRACKER):
        with open(PROCESSED_FORM_TRACKER, 'w') as f:
            f.write('[]')

def process_email(msg, rows):
    body = get_clean_email_body(msg)
    parsed = parse_email_body(body)
    ticket_number = parsed.get('ticket')
    if parsed.get('error'):
        # send error email
        acct = create_account()
        client = EmailClient(acct)
        client.send_message(msg.sender.address, f'Issue Processing Ticket {ticket_number or ""}', parsed['error'], None)
        msg.mark_as_read()
        return
    if not ticket_number:
        acct = create_account()
        client = EmailClient(acct)
        client.send_message(msg.sender.address, 'Issue Processing Ticket', 'No ticket number found in your message.', None)
        msg.mark_as_read()
        return
    row = get_ticket_by_number(ticket_number, rows)
    if not row:
        acct = create_account()
        client = EmailClient(acct)
        client.send_message(msg.sender.address, f'Ticket {ticket_number} Not Found', f'Ticket #{ticket_number} not found in Smartsheet.', None)
        msg.mark_as_read()
        return

    sender_name = msg.sender.address.split('@')[0].replace('.', ' ').title()
    sent_date = msg.received.strftime('%m/%d/%Y')
    short_date = msg.received.strftime('%m-%d-%y')

    field_map = {
        'SERVICE TICKET': ticket_number,
        'COMPANY': row.get('site', ''),
        'SITE NAME': row.get('site', ''),
        'REQUESTED BY': row.get('requestor', ''),
        'SITE ADDRESS': row.get('address', ''),
        'TICKET REQUESTRow1': row.get('problem', ''),
        'TECHRow1': sender_name,
        'TECHNICIAN NOTESRow1': parsed.get('tech_notes',''),
        'ADDITIONAL NOTESRow1': parsed.get('additional_notes',''),
        'HOURSRow1': parsed.get('time_spent',''),
        'DATERow1': sent_date
    }

    clean_site = (row.get('site') or 'NO_SITE').strip().upper()
    clean_site = ''.join(ch for ch in clean_site if ch.isalnum() or ch in (' ','-')).replace(' ','_') or 'NO_SITE'
    pdf_filename = f"{ticket_number} - {clean_site} - {short_date} - PurpleDoc.pdf"
    fill_pdf(PDF_TEMPLATE, pdf_filename, field_map)
    acct = create_account()
    client = EmailClient(acct)
    client.send_message(msg.sender.address, f'Purple Doc Report for Ticket #{ticket_number}', 'Attached is your Purple Doc form.', [pdf_filename])
    msg.mark_as_read()

def process_form_row(form_row, rows, drive_id):
    ticket_number = str(form_row.get('ticket number', '')).strip()
    if not ticket_number:
        return None
    email = form_row.get('email', '')
    name = form_row.get('name', '')
    work_done = form_row.get('work done', '')
    hours = str(form_row.get('time spent', '')).strip()
    status = str(form_row.get('ticket status', '')).strip().lower()
    completion = form_row.get('completion time', '')
    try:
        sent_date = completion.split()[0]
        short_date = sent_date.replace('/', '-')[2:]
    except Exception:
        sent_date = time.strftime('%m/%d/%Y')
        short_date = time.strftime('%m-%d-%y')

    row = get_ticket_by_number(ticket_number, rows)
    if not row:
        return None

    field_map = {
        'SERVICE TICKET': ticket_number,
        'COMPANY': row.get('site', ''),
        'SITE NAME': row.get('site', ''),
        'REQUESTED BY': row.get('requestor', ''),
        'SITE ADDRESS': row.get('address', ''),
        'TICKET REQUESTRow1': row.get('problem', ''),
        'TECHRow1': name,
        'TECHNICIAN NOTESRow1': work_done,
        'ADDITIONAL NOTESRow1': status if status in ['ongoing','close','closed'] else 'ongoing',
        'HOURSRow1': hours,
        'DATERow1': sent_date
    }
    clean_site = (row.get('site') or 'NO_SITE').strip().upper()
    clean_site = ''.join(ch for ch in clean_site if ch.isalnum() or ch in (' ','-')).replace(' ','_') or 'NO_SITE'
    pdf_filename = f"{ticket_number} - {clean_site} - {short_date} - PurpleDoc.pdf"
    fill_pdf(PDF_TEMPLATE, pdf_filename, field_map)
    acct = create_account()
    client = EmailClient(acct)
    client.send_message(email, f'Purple Doc Report for Ticket #{ticket_number}', 'Attached is your Purple Doc form.', [pdf_filename])
    return ticket_number

def main_loop(drive_id=None):
    columns, rows, conversations, ts = load_smartsheet_cache()
    if not rows:
        columns, rows, conversations = fetch_smartsheet_data_with_conversations()
        save_smartsheet_cache(columns, rows, conversations)
    acct = create_account()
    client = EmailClient(acct)
    ensure_processed_tracker()
    while True:
        try:
            # refresh smartsheet cache
            latest_cols, latest_rows, latest_conv = fetch_smartsheet_data_with_conversations()
            if latest_rows and len(latest_rows) != len(rows):
                rows = latest_rows
                save_smartsheet_cache(latest_cols, rows, latest_conv)

            # process emails
            msgs = client.fetch_unread_pd_messages(limit=20)
            for m in msgs:
                process_email(m, rows)

            # process form rows if drive_id provided
            if drive_id:
                form_rows = get_excel_form_rows(drive_id)
                # load tracker
                try:
                    with open(PROCESSED_FORM_TRACKER, 'r') as f:
                        seen = set(json.load(f))
                except Exception:
                    seen = set()
                new_ids = set()
                for fr in form_rows:
                    rid = str(fr.get('id','')).strip()
                    if not rid or rid in seen:
                        continue
                    if process_form_row(fr, rows, drive_id):
                        new_ids.add(rid)
                if new_ids:
                    seen.update(new_ids)
                    with open(PROCESSED_FORM_TRACKER, 'w') as f:
                        json.dump(sorted(list(seen)), f)
            print('Idle. Sleeping 30s...')
        except Exception as e:
            print('Error in loop:', e)
        time.sleep(30)

if __name__ == '__main__':
    # Optionally pass drive_id via env var or leave None
    drive = os.getenv('ONEDRIVE_DRIVE_ID')
    main_loop(drive)
