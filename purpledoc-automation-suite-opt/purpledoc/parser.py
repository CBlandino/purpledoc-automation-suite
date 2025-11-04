import re
from collections import defaultdict
from rapidfuzz import fuzz

SIGNATURE_PATTERNS = [
    r'^--\s*$', r'^thanks[\s,]*$', r'^regards[\s,]*$',
    r'Get Outlook for iOS', r'Get Outlook for Android', r'Powered by O365',
]

def html_to_clean_text(html: str) -> str:
    html = re.sub(r'(?i)<br\s*/?>', '\n', html)
    html = re.sub(r'(?i)</p\s*>', '\n', html)
    html = re.sub(r'(?i)</div\s*>', '\n', html)
    html = re.sub(r'(?i)<(p|div|li|tr|table)[^>]*>', '\n', html)
    text = re.sub(r'<[^>]+>', '', html)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n+', '\n', text)
    lines = [line.strip() for line in text.split('\n')]
    return '\n'.join([line for line in lines if line])

def strip_signature(body: str) -> str:
    lines = body.strip().splitlines()
    stripped = []
    for line in lines:
        if any(re.search(pat, line.strip(), re.IGNORECASE) for pat in SIGNATURE_PATTERNS):
            break
        stripped.append(line)
    return '\n'.join(stripped)

def get_clean_email_body(msg) -> str:
    if getattr(msg, 'body', None) and getattr(msg, 'body_type', None):
        if msg.body_type.lower() == 'html':
            return html_to_clean_text(msg.body)
        return msg.body
    # fallback to raw body (graph)
    if hasattr(msg, '_raw_message') and msg._raw_message:
        body_content = msg._raw_message.get('body', {})
        content_type = body_content.get('contentType', '').lower()
        content = body_content.get('content', '')
        if content_type == 'html':
            return html_to_clean_text(content)
        return content
    return ''

def fuzzy_contains(text: str, keywords, threshold=80) -> bool:
    return any(fuzz.partial_ratio(text.lower(), kw.lower()) >= threshold for kw in keywords)

def parse_email_body(body: str, default_tech_name='Unknown'):
    body = re.sub(r'(?i)<br\s*/?>', '\n', body)
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = strip_signature(body)
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

    # Ticket number detection (6-digit fallback)
    for line in lines:
        if 'ticket' in line.lower():
            m = re.search(r'\d{6}', line)
            if m:
                data['ticket'] = m.group(0)
                break
    if not data['ticket']:
        m = re.search(r'\b\d{6}\b', flattened)
        if m:
            data['ticket'] = m.group(0)

    is_closed = 'close' in flattened or 'closed' in flattened
    is_ongoing = 'ongoing' in flattened

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
                except Exception:
                    data['error'] = f"Invalid time format '{raw}' for tech {tech}"
            else:
                notes.append(l)
        data['techs'][tech]['notes'] = '\n'.join(notes).strip()
        data['techs'][tech]['time'] = time_spent

    if not data['techs']:
        data['techs'][default_tech_name]['notes'] = '\n'.join(lines)

    if is_closed:
        data['additional_notes'] = 'close'
    elif is_ongoing:
        data['additional_notes'] = 'ongoing'
    else:
        data['additional_notes'] = 'ongoing'

    # Provide friendly summary fields
    # primary tech is the first key
    primary = next(iter(data['techs']), default_tech_name)
    data['tech_notes'] = data['techs'][primary]['notes']
    data['time_spent'] = data['techs'][primary]['time']
    return data
