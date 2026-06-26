# Forsaken Creative Request Portal

This is a small web form that writes new creative requests into:

`https://docs.google.com/spreadsheets/d/1FRv0kfJc10hZLygUnApg5kbbFQ1Bnvxt2aE3-HQn5jQ/edit?usp=drive_link`

It also sends an email notification after each saved request when SMTP settings are configured.

Clients can paste an asset link or upload photos/videos directly from the form. Direct uploads save into the Google Drive folder configured in Apps Script.

The admin dashboard is available at `/admin`. Set `ADMIN_DASHBOARD_PASSWORD` in Vercel to protect it.

The simplified custom form is available locally at `/simple`. After redeploying Apps Script, the Google-hosted simplified form is available at the Apps Script web app URL with `?view=simpleForm`.

Clients can check request progress at `/process` with their request ID and email. Progress states use red for not started, yellow for in process, and green for done.

The public work board is available at `/work-in-process` and only shows member name, design type, and stage.

Admins can manage designer names and emails from the `Designers` tab in `/admin`. Assigned Designer uses that list as a dropdown.

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

## Admin Dashboard

The dashboard reads and updates the `Request Tracker`, `Approved`, and `Disapproved` tabs through Apps Script. It manages request status, seen status, approval status, assigned designer, design dates, final links, admin notes, single deletion, and bulk deletion. Approved requests move to the `Approved` tab, and `Not Approved` or `Needs Info` requests move to the `Disapproved` tab.

After changing `google-apps-script/Code.gs` or `google-apps-script/appsscript.json`, redeploy the Apps Script web app so the form and `/admin` dashboard can use the latest actions.

## Vercel Google Sheets Setup

Recommended no-key setup: use Google Apps Script.

1. Open your Apps Script project.
2. Paste the contents of `google-apps-script/Code.gs`.
3. Open Project Settings and enable `Show "appsscript.json" manifest file in editor`.
4. Paste the contents of `google-apps-script/appsscript.json` into `appsscript.json`.
5. Change `SECRET` in Apps Script to a private phrase.
6. Click `Deploy` -> `New deployment`.
7. Select type `Web app`.
8. Set `Execute as` to `Me`.
9. Set `Who has access` to `Anyone`.
10. Authorize the requested Google Sheets and Google Drive permissions.
11. Copy the Web App URL ending in `/exec`.
12. Add it to Vercel:

```powershell
.\scripts\set_apps_script_webhook_to_vercel.ps1 -WebhookUrl "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec" -Secret "same-private-phrase"
npx vercel deploy --prod --yes
```

Alternative service-account setup:

```text
GOOGLE_SHEET_ID=1FRv0kfJc10hZLygUnApg5kbbFQ1Bnvxt2aE3-HQn5jQ
GOOGLE_SHEET_URL=https://docs.google.com/spreadsheets/d/1FRv0kfJc10hZLygUnApg5kbbFQ1Bnvxt2aE3-HQn5jQ/edit?usp=drive_link
GOOGLE_SERVICE_ACCOUNT_JSON={...service account json...}
GOOGLE_APPS_SCRIPT_WEBHOOK_URL=
GOOGLE_APPS_SCRIPT_SECRET=change-this-secret
```

Share the Google Sheet with the service account email as an editor.

After downloading the service account key JSON, add it to Vercel without printing it:

```powershell
.\scripts\set_google_key_to_vercel.ps1 -KeyPath "C:\path\to\service-account.json"
npx vercel deploy --prod --yes
```

## Publishing Online

Deploy with Vercel after updating the webhook or email settings:

```powershell
npx vercel deploy --prod --yes --scope naxystudiosllc-1953s-projects
```
