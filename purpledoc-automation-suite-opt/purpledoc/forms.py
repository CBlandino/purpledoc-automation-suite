import urllib.parse, json
from .config import O365_TOKEN_FILE
import requests

def get_excel_form_rows(drive_id: str, filename="Purple Doc _Online Form.xlsx", worksheet_name="Sheet1"):
    try:
        with open(O365_TOKEN_FILE, 'r') as f:
            token_file = json.load(f)
        access_token_entry = next(iter(token_file.get('AccessToken', {}).values()))
        access_token = access_token_entry['secret']
    except Exception as exc:
        raise RuntimeError('Failed to read access token') from exc

    headers = {'Authorization': f'Bearer {access_token}'}
    encoded_filename = urllib.parse.quote(filename)
    file_metadata_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_filename}"
    res = requests.get(file_metadata_url, headers=headers)
    res.raise_for_status()
    file_id = res.json()['id']

    worksheets_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/workbook/worksheets"
    res_ws = requests.get(worksheets_url, headers=headers)
    res_ws.raise_for_status()
    worksheets = res_ws.json().get('value', [])
    ws_names = [ws['name'] for ws in worksheets]
    ws_name = worksheet_name if worksheet_name in ws_names else (ws_names[0] if ws_names else worksheet_name)

    encoded_ws_name = urllib.parse.quote(ws_name)
    used_range_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/workbook/worksheets('{encoded_ws_name}')/usedRange"
    res_range = requests.get(used_range_url, headers=headers)
    res_range.raise_for_status()
    data = res_range.json()
    values = data.get('values', [])
    if not values or len(values) < 2:
        return []

    headers_row = [str(h).strip().lower() for h in values[0]]
    return [dict(zip(headers_row, row)) for row in values[1:] if any(row)]
