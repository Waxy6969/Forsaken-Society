# Forsaken Creative Request Portal

This is a small web form that writes new creative requests into:

`G:/downloads/Creative_Request_Portal_Forsaken_Form.xlsx`

It also sends an email notification after each saved request when SMTP settings are configured.

Clients upload project assets through the `UPLOAD_FOLDER` Google Drive link. If they do not paste a specific asset link into the form, that upload folder is saved to the tracker automatically.

## Run

```powershell
python -m pip install -r requirements.txt
copy .env.example .env
python app.py
```

Open:

```text
http://127.0.0.1:8000
```

## Email Setup

Edit `.env` and set:

```text
ADMIN_EMAIL=your-request-inbox@example.com
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_TLS=true
SMTP_USER=your-sending-account@example.com
SMTP_PASS=your-app-password
SMTP_FROM=your-sending-account@example.com
```

For Gmail, use an app password instead of your normal account password.

## Publishing Online

The app is ready to run on a Windows server, VPS, or platform that supports Python. Keep the workbook path available to the server, or update `WORKBOOK_PATH` in `.env` to the server-side copy of the spreadsheet.

Each submission creates a timestamped backup in `backups/` before the workbook is saved.
