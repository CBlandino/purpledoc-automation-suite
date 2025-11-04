# PurpleDoc Automation Suite

This project automates the generation and delivery of **Purple Doc service reports** by integrating **Microsoft Outlook (O365)**, **Smartsheet**, and **Microsoft Forms/Excel**.

It processes technician emails and online form submissions to automatically:
- Parse technician-submitted notes and ticket data.
- Fetch ticket metadata from Smartsheet.
- Fill PDF templates with accurate details.
- Reply to the sender with the completed Purple Doc form.
- Maintain local caches and prevent duplicate processing.

## 🚀 Features
- Microsoft Graph API Integration — Reads and replies to Outlook emails, and fetches Excel form data from OneDrive.
- Smartsheet Integration — Retrieves ticket metadata and comments for mapping.
- Automated PDF Filling — Uses `pdfrw` to populate template fields from email or form data.
- Fuzzy Parsing Engine — Extracts ticket numbers, notes, and time spent using `rapidfuzz`.
- Caching System — Reduces API calls with locally stored Smartsheet data.
- Continuous Monitoring Loop — Automatically polls for new emails, form entries, and Smartsheet updates.

## 🧩 Requirements
Install dependencies with:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- O365
- python-dotenv
- smartsheet-python-sdk
- rapidfuzz
- pdfrw
- requests

## ⚙️ Configuration

Create a `.env` file in the project root with the following variables:

```env
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
SMARTSHEET_TOKEN=your_smartsheet_api_token
SHEET_ID=1234567890123456
EMAIL_ADDRESS=you@domain.com
EMAIL_PASSWORD=your_email_password
SMTP_SERVER=smtp.office365.com
SMTP_PORT=587
TENANT_ID=your_tenant_id
```

## 🧠 How It Works

1. **O365 Authentication**: The script authenticates with Microsoft Graph using OAuth (stored in `o365_token.txt`).
2. **Smartsheet Data Caching**: Data is fetched once and cached locally in `smartsheet_cache.json`.
3. **Email Parsing**: Detects unread emails with subject `PD`, extracts details, and replies with a filled Purple Doc PDF.
4. **Microsoft Form Integration**: Reads a designated Excel workbook from OneDrive and generates Purple Docs for new rows.
5. **Continuous Loop**: Checks every 30 seconds for new Smartsheet updates, PD emails, and form submissions.

## 📄 File Structure

```
purpledoc-automation-suite/
├── main.py
├── .env
├── requirements.txt
├── smartsheet_cache.json
├── processed_form_rows.json
├── 000000 - Template.pdf
└── README.md
```

## 🧰 Running the Script

```bash
python main.py
```

## 🧑‍💻 Author
**Christopher Blandino**  
📧 [ChristopherBlandino0@gmail.com](mailto:ChristopherBlandino0@gmail.com)  
🔗 [GitHub: CBlandino](https://github.com/CBlandino)

## 📜 License
MIT License © 2025 Christopher Blandino
