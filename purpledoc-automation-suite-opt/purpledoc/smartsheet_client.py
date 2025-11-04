import os, json, time, re
from .config import SMARTSHEET_TOKEN, SHEET_ID, SMARTSHEET_CACHE_FILE
import smartsheet

def fetch_smartsheet_conversations(ss_client, sheet_id, row_ids):
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
                } for c in comments
            ]
        except Exception:
            conversations[str(row_id)] = []
    return conversations

def fetch_smartsheet_data_with_conversations():
    ss_client = smartsheet.Smartsheet(SMARTSHEET_TOKEN)
    sheet = ss_client.Sheets.get_sheet(SHEET_ID)
    columns = [{"id": col.id, "title": col.title.strip().lower()} for col in sheet.columns]
    rows = []
    row_ids = []
    for row in sheet.rows:
        row_dict = {sheet.columns[i].title.lower(): cell.value for i, cell in enumerate(row.cells)}
        row_dict["_row_id"] = row.id
        rows.append(row_dict)
        row_ids.append(row.id)
    conversations = fetch_smartsheet_conversations(ss_client, SHEET_ID, row_ids)
    return columns, rows, conversations

def load_smartsheet_cache():
    if os.path.exists(SMARTSHEET_CACHE_FILE):
        try:
            with open(SMARTSHEET_CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('columns', []), data.get('rows', []), data.get('conversations', {}), data.get('timestamp', 0)
        except Exception:
            return [], [], {}, 0
    return [], [], {}, 0

def save_smartsheet_cache(columns, rows, conversations):
    with open(SMARTSHEET_CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump({
            'columns': columns,
            'rows': rows,
            'conversations': conversations,
            'timestamp': int(time.time())
        }, f, ensure_ascii=False, indent=2)

def get_ticket_by_number(ticket_number, rows):
    def normalize_ticket(ticket_str):
        if not ticket_str:
            return ''
        ticket_str = str(ticket_str).strip().lower()
        if ticket_str.endswith('.0'):
            ticket_str = ticket_str[:-2]
        ticket_str = re.sub(r'\W+', '', ticket_str)
        return ticket_str
    normalized_target = normalize_ticket(ticket_number)
    for row in rows:
        if normalize_ticket(row.get('ticket number', '')) == normalized_target:
            return row
    return None
