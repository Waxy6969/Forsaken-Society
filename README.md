# Forsaken Creative Request Portal

This is a small web form that writes new creative requests into:

`G:/downloads/Creative_Request_Portal_Forsaken_Form.xlsx`

On Vercel, requests are recorded in this Google Sheet:

`https://docs.google.com/spreadsheets/d/1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU/edit?usp=drivesdk`

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

## Vercel Google Sheets Setup

Set these Vercel environment variables:

```text
GOOGLE_SHEET_ID=1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU
GOOGLE_SHEET_URL=https://docs.google.com/spreadsheets/d/1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU/edit?usp=drivesdk
GOOGLE_SERVICE_ACCOUNT_JSON={...service account json...}
```

Share the Google Sheet with the service account email as an editor.

After downloading the service account key JSON, add it to Vercel without printing it:

```powershell
.\scripts\set_google_key_to_vercel.ps1 -KeyPath "C:\path\to\service-account.json"
npx vercel deploy --prod --yes
```

## Publishing Online

The app is ready to run on a Windows server, VPS, or platform that supports Python. Keep the workbook path available to the server, or update `WORKBOOK_PATH` in `.env` to the server-side copy of the spreadsheet.

Each submission creates a timestamped backup in `backups/` before the workbook is saved.
