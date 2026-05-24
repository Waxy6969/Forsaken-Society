from __future__ import annotations

import html
import hashlib
import json
import mimetypes
import os
import re
import smtplib
import urllib.error
import urllib.request
from datetime import datetime
from email.message import EmailMessage
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlencode, unquote, urlparse


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_GOOGLE_SHEET_ID = "1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU"
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU/edit?usp=drivesdk"
ENV_PATH = BASE_DIR / ".env"

TRACKER_SHEET = "Request Tracker"
DEFAULT_UPLOAD_FOLDER = "https://drive.google.com/drive/folders/17Elym_RLgFLL2EPOOgS2FKhwNA3ikd-f?usp=sharing"
ADMIN_COOKIE_NAME = "forsaken_admin"

ADMIN_STATUS_OPTIONS = ["Submitted", "In Progress", "Revision", "Completed", "Cancelled", "Deleted"]
SEEN_STATUS_OPTIONS = ["Not Seen", "Seen"]
APPROVAL_STATUS_OPTIONS = ["Pending Review", "Approved", "Not Approved", "Needs Info"]

FIELD_LABELS = {
    "member_name": "Member Name",
    "email": "Email",
    "project_name": "Project Name",
    "design_type": "Design Type",
    "description": "Design Description",
    "requested_deadline": "Requested Deadline",
    "priority": "Priority Level",
    "rush_option": "Expedited Option",
    "uploaded_files_link": "File or URL Link",
    "notes": "Notes for Designer",
}

REQUIRED_FIELDS = [
    "member_name",
    "email",
    "project_name",
    "design_type",
    "description",
    "requested_deadline",
    "priority",
    "rush_option",
]


def load_env() -> dict[str, str]:
    env = dict(os.environ)
    if ENV_PATH.exists():
        for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            value = value.strip().strip('"').strip("'")
            env.setdefault(key.strip(), value)
    return env


def get_config() -> dict[str, str]:
    env = load_env()
    return {
        "google_sheet_id": env.get("GOOGLE_SHEET_ID", DEFAULT_GOOGLE_SHEET_ID),
        "google_sheet_url": env.get("GOOGLE_SHEET_URL", DEFAULT_GOOGLE_SHEET_URL),
        "google_service_account_json": env.get("GOOGLE_SERVICE_ACCOUNT_JSON", ""),
        "google_apps_script_webhook_url": env.get("GOOGLE_APPS_SCRIPT_WEBHOOK_URL", ""),
        "google_apps_script_secret": env.get("GOOGLE_APPS_SCRIPT_SECRET", ""),
        "admin_email": env.get("ADMIN_EMAIL", ""),
        "smtp_host": env.get("SMTP_HOST", ""),
        "smtp_port": env.get("SMTP_PORT", "587"),
        "smtp_user": env.get("SMTP_USER", ""),
        "smtp_pass": env.get("SMTP_PASS", ""),
        "smtp_from": env.get("SMTP_FROM", env.get("SMTP_USER", "")),
        "smtp_tls": env.get("SMTP_TLS", "true").lower(),
        "company_team": env.get("COMPANY_TEAM", "Forsaken"),
        "upload_folder": env.get("UPLOAD_FOLDER", DEFAULT_UPLOAD_FOLDER),
        "admin_dashboard_password": env.get("ADMIN_DASHBOARD_PASSWORD", ""),
        "host": env.get("HOST", "127.0.0.1"),
        "port": env.get("PORT", "8000"),
    }


def cell_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def read_choices() -> dict[str, list[str]]:
    return {
        "designTypes": [
            "AVIs",
            "Twitter Headers",
            "YouTube Headers",
            "TikTok Banners",
            "Social Media Packages",
            "3D Animations",
            "Intros or Outros",
            "Simple Editing",
            "Advanced Editing",
            "Other",
        ],
        "priorities": ["Standard", "Expedited"],
        "rushOptions": ["No Rush", "24 Hour Rush +$15", "Same Day Rush +$35"],
    }


def parse_form(body: bytes) -> dict[str, str]:
    parsed = parse_qs(body.decode("utf-8"), keep_blank_values=True)
    return {key: values[0].strip() for key, values in parsed.items()}


def validate_form(data: dict[str, str]) -> list[str]:
    errors = []
    for field in REQUIRED_FIELDS:
        if not data.get(field):
            errors.append(f"{FIELD_LABELS[field]} is required.")
    if data.get("email") and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", data["email"]):
        errors.append("Enter a valid email address.")
    return errors


def rush_fee_text(rush_option: str) -> str:
    if "Same Day" in rush_option:
        return "$35 rush fee"
    if "24 Hour" in rush_option:
        return "$15 rush fee"
    return "No additional fee"


def combined_admin_notes(data: dict[str, str]) -> str:
    parts = []
    if data.get("notes"):
        parts.append(f"Notes: {data['notes']}")
    if data.get("other_design_type"):
        parts.append(f"Other design type: {data['other_design_type']}")
    if data.get("rush_payment_confirmed"):
        parts.append(f"Rush payment confirmed: {data['rush_payment_confirmed']}")
    if data.get("rush_payment_link"):
        parts.append(f"Rush payment link: {data['rush_payment_link']}")
    return "\n".join(parts)


def build_tracker_values(data: dict[str, str], request_id: str, now: datetime, config: dict[str, str], *, as_text: bool) -> list[Any]:
    uploaded_link = data.get("uploaded_files_link") or ""
    notes = combined_admin_notes(data)
    submitted = now.strftime("%Y-%m-%d %I:%M %p") if as_text else now
    deadline = data["requested_deadline"]
    return [
        request_id,
        submitted,
        data["member_name"],
        config["company_team"],
        data["email"],
        data["project_name"],
        data["design_type"],
        data["description"],
        deadline,
        data["priority"],
        data["rush_option"],
        rush_fee_text(data["rush_option"]),
        "Submitted",
        "Not Seen",
        "",
        "Pending Review",
        "",
        "",
        "",
        "",
        uploaded_link,
        "",
        submitted,
        notes,
    ]


def next_google_request_id(values: list[str]) -> str:
    highest = 0
    for value in values:
        match = re.fullmatch(r"CRP-(\d+)", cell_text(value))
        if match:
            highest = max(highest, int(match.group(1)))
    return f"CRP-{highest + 1:04d}"


def save_submission_to_apps_script(data: dict[str, str], config: dict[str, str]) -> dict[str, str]:
    if not config["google_apps_script_webhook_url"]:
        raise RuntimeError(
            "Requests are set to record in Google Sheets, but the Apps Script webhook URL is not configured in Vercel."
        )
    now = datetime.now()
    payload = {
        "secret": config["google_apps_script_secret"],
        "submitted_at": now.strftime("%Y-%m-%d %I:%M %p"),
        "values": build_tracker_values(data, "", now, config, as_text=True),
        "fields": data,
        "files": json.loads(data.get("uploaded_files_json") or "[]"),
    }
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        config["google_apps_script_webhook_url"],
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=55) as response:
            response_payload = json.loads(response.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google Apps Script rejected the request: {message}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach the Google Apps Script webhook: {exc.reason}") from exc

    if not response_payload.get("ok"):
        raise RuntimeError(response_payload.get("error") or "Google Apps Script did not confirm the request was saved.")

    return {
        "request_id": response_payload.get("request_id", "CRP-SAVED"),
        "submitted_at": now.strftime("%Y-%m-%d %I:%M %p"),
        "email_sent": "false",
        "storage": "apps_script",
    }


def call_apps_script(payload: dict[str, Any], config: dict[str, str]) -> dict[str, Any]:
    if not config["google_apps_script_webhook_url"]:
        raise RuntimeError("The Apps Script webhook is not configured.")
    payload = {"secret": config["google_apps_script_secret"], **payload}
    request = urllib.request.Request(
        config["google_apps_script_webhook_url"],
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=55) as response:
            response_payload = json.loads(response.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google Apps Script rejected the admin request: {message}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach the Google Apps Script webhook: {exc.reason}") from exc
    if not response_payload.get("ok"):
        raise RuntimeError(response_payload.get("error") or "Google Apps Script did not confirm the admin request.")
    return response_payload


def get_apps_script(payload: dict[str, str], config: dict[str, str]) -> dict[str, Any]:
    if not config["google_apps_script_webhook_url"]:
        raise RuntimeError("The Apps Script webhook is not configured.")
    query = urlencode({"secret": config["google_apps_script_secret"], **payload})
    url = f"{config['google_apps_script_webhook_url']}?{query}"
    try:
        with urllib.request.urlopen(url, timeout=55) as response:
            response_payload = json.loads(response.read().decode("utf-8") or "{}")
    except json.JSONDecodeError as exc:
        raise RuntimeError("Redeploy the Google Apps Script web app so the admin dashboard can read requests.") from exc
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google Apps Script rejected the admin request: {message}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach the Google Apps Script webhook: {exc.reason}") from exc
    if not response_payload.get("ok"):
        raise RuntimeError(response_payload.get("error") or "Google Apps Script did not confirm the admin request.")
    return response_payload


def fetch_admin_requests(config: dict[str, str]) -> list[dict[str, Any]]:
    payload = get_apps_script({"action": "listRequests"}, config)
    if "requests" not in payload:
        raise RuntimeError("Redeploy the Google Apps Script web app so the admin dashboard can read requests.")
    requests = payload.get("requests", [])
    if not isinstance(requests, list):
        return []
    return [item for item in requests if cell_text(item.get("request_status")) != "Deleted"]


def update_admin_request(data: dict[str, str], config: dict[str, str]) -> None:
    version_payload = get_apps_script({"action": "version"}, config)
    if not version_payload.get("admin_dashboard"):
        raise RuntimeError("Redeploy the Google Apps Script web app so the admin dashboard can save updates.")
    call_apps_script(
        {
            "action": "updateRequest",
            "request_id": data.get("request_id", ""),
            "updates": {
                "request_status": data.get("request_status", ""),
                "seen_status": data.get("seen_status", ""),
                "seen_by": data.get("seen_by", ""),
                "approval_status": data.get("approval_status", ""),
                "assigned_designer": data.get("assigned_designer", ""),
                "design_start_date": data.get("design_start_date", ""),
                "design_due_date": data.get("design_due_date", ""),
                "final_deliverables_link": data.get("final_deliverables_link", ""),
                "admin_notes": data.get("admin_notes", ""),
            },
        },
        config,
    )


def delete_admin_request(data: dict[str, str], config: dict[str, str]) -> None:
    version_payload = get_apps_script({"action": "version"}, config)
    if not version_payload.get("admin_dashboard"):
        raise RuntimeError("Redeploy the Google Apps Script web app so the admin dashboard can delete requests.")
    if version_payload.get("supports_delete"):
        call_apps_script(
            {
                "action": "deleteRequest",
                "request_id": data.get("request_id", ""),
            },
            config,
        )
        return
    call_apps_script(
        {
            "action": "updateRequest",
            "request_id": data.get("request_id", ""),
            "updates": {
                "request_status": "Deleted",
                "admin_notes": "Deleted from admin dashboard.",
            },
        },
        config,
    )


def save_submission_to_google_sheet(data: dict[str, str], config: dict[str, str]) -> dict[str, str]:
    if not config["google_service_account_json"]:
        raise RuntimeError(
            "Requests are set to record in Google Sheets, but the site still needs the Google Sheets connection added in Vercel."
        )
    try:
        import gspread
    except ImportError as exc:
        raise RuntimeError("Google Sheets support requires gspread from requirements.txt.") from exc

    credentials = json.loads(config["google_service_account_json"])
    client = gspread.service_account_from_dict(credentials)
    spreadsheet = client.open_by_key(config["google_sheet_id"])
    worksheet = spreadsheet.worksheet(TRACKER_SHEET)
    request_id = next_google_request_id(worksheet.col_values(1))
    now = datetime.now()
    values = build_tracker_values(data, request_id, now, config, as_text=True)
    worksheet.append_row(values, value_input_option="USER_ENTERED")
    return {
        "request_id": request_id,
        "submitted_at": now.strftime("%Y-%m-%d %I:%M %p"),
        "email_sent": "false",
        "storage": "google_sheet",
    }


def save_submission(data: dict[str, str]) -> dict[str, str]:
    config = get_config()
    if config["google_apps_script_webhook_url"]:
        return save_submission_to_apps_script(data, config)
    if config["google_service_account_json"]:
        return save_submission_to_google_sheet(data, config)
    raise RuntimeError("Google Sheets recording is not configured yet.")


def send_email(data: dict[str, str], result: dict[str, str]) -> bool:
    config = get_config()
    required = ["admin_email", "smtp_host", "smtp_port", "smtp_from"]
    if any(not config[key] for key in required):
        return False

    msg = EmailMessage()
    msg["Subject"] = f"New creative request: {result['request_id']} - {data['project_name']}"
    msg["From"] = config["smtp_from"]
    msg["To"] = config["admin_email"]
    if data.get("email"):
        msg["Reply-To"] = data["email"]

    lines = [
        f"Request ID: {result['request_id']}",
        f"Submitted: {result['submitted_at']}",
        "",
    ]
    for key, label in FIELD_LABELS.items():
        if data.get(key):
            lines.append(f"{label}: {data[key]}")
    msg.set_content("\n".join(lines))

    rows = "".join(
        f"<tr><th>{html.escape(label)}</th><td>{html.escape(data.get(key, ''))}</td></tr>"
        for key, label in FIELD_LABELS.items()
        if data.get(key)
    )
    msg.add_alternative(
        f"""
        <html>
          <body>
            <h2>New creative request: {html.escape(result['request_id'])}</h2>
            <p>Submitted {html.escape(result['submitted_at'])}</p>
            <table cellpadding="8" cellspacing="0" border="1">{rows}</table>
          </body>
        </html>
        """,
        subtype="html",
    )

    port = int(config["smtp_port"])
    if port == 465:
        server: smtplib.SMTP = smtplib.SMTP_SSL(config["smtp_host"], port, timeout=20)
    else:
        server = smtplib.SMTP(config["smtp_host"], port, timeout=20)
    try:
        if config["smtp_tls"] == "true" and port != 465:
            server.starttls()
        if config["smtp_user"] and config["smtp_pass"]:
            server.login(config["smtp_user"], config["smtp_pass"])
        server.send_message(msg)
    finally:
        server.quit()
    return True


def page_template(content: str, status: str = "") -> bytes:
    choices_json = json.dumps(read_choices())
    config = get_config()
    html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Forsaken Creative Request Portal</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #151515;
      --muted: #6d6d6d;
      --line: #e4e0da;
      --field: #f8f5ef;
      --accent: #e74719;
      --accent-dark: #b93412;
      --gold: #f7b733;
      --warn: #a24112;
      --ok: #166534;
      --page: #111111;
      --panel: #ffffff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      color: var(--ink);
      background:
        linear-gradient(90deg, rgba(10, 10, 10, .94), rgba(10, 10, 10, .56)),
        url("/static/forsaken-background.png") center top / cover fixed,
        var(--page);
    }}
    body::before {{
      content: "";
      position: fixed;
      inset: -7%;
      z-index: 0;
      pointer-events: none;
      background:
        radial-gradient(circle at 16% 24%, rgba(247, 183, 51, .30), transparent 20%),
        radial-gradient(circle at 70% 18%, rgba(231, 71, 25, .34), transparent 24%),
        radial-gradient(circle at 88% 76%, rgba(255, 255, 255, .16), transparent 18%),
        linear-gradient(120deg, rgba(231, 71, 25, .16), transparent 44%, rgba(247, 183, 51, .14));
      filter: blur(24px) saturate(1.35);
      opacity: .92;
      animation: glowShift 11s ease-in-out infinite alternate;
    }}
    @keyframes glowShift {{
      0% {{ transform: translate3d(-2%, -1%, 0) scale(1); opacity: .78; }}
      100% {{ transform: translate3d(2%, 2%, 0) scale(1.04); opacity: .96; }}
    }}
    main {{
      position: relative;
      z-index: 2;
      width: min(1180px, calc(100% - 32px));
      margin: 0 auto;
      min-height: 100vh;
      display: grid;
      grid-template-columns: minmax(280px, .75fr) minmax(0, 1.25fr);
      align-items: center;
      gap: 34px;
      padding: 34px 0;
    }}
    header {{
      color: #ffffff;
      padding: 18px 0;
    }}
    .eyebrow {{
      display: inline-block;
      color: var(--gold);
      font-size: .78rem;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: uppercase;
      margin-bottom: 13px;
    }}
    h1 {{
      margin: 0;
      font-size: clamp(2.65rem, 6vw, 5.6rem);
      line-height: .9;
      text-transform: uppercase;
      max-width: 8ch;
    }}
    p {{ margin: 0; color: var(--muted); line-height: 1.5; }}
    .form-shell {{
      background: rgba(255, 255, 255, .96);
      border: 1px solid rgba(255, 255, 255, .45);
      border-radius: 8px;
      box-shadow: 0 22px 70px rgba(0, 0, 0, .36);
      overflow: hidden;
    }}
    .form-title {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      padding: 22px 26px;
      background: #161616;
      color: #ffffff;
      border-bottom: 4px solid var(--accent);
    }}
    .form-title h2 {{
      margin: 0;
      font-size: clamp(1.25rem, 2vw, 1.8rem);
      text-transform: uppercase;
    }}
    .tag {{
      color: var(--gold);
      font-weight: 800;
      font-size: .85rem;
      white-space: nowrap;
    }}
    .top-links {{
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }}
    .admin-login-link {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 36px;
      border: 1px solid rgba(255, 255, 255, .42);
      border-radius: 6px;
      padding: 0 12px;
      color: #ffffff;
      text-decoration: none;
      font-size: .78rem;
      font-weight: 900;
      text-transform: uppercase;
      white-space: nowrap;
    }}
    form {{ padding: 26px; }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    label {{ display: grid; gap: 7px; font-weight: 700; font-size: .92rem; }}
    .hidden {{ display: none; }}
    input, select, textarea {{
      width: 100%;
      border: 1px solid #d2ccc2;
      border-radius: 6px;
      background: var(--field);
      color: var(--ink);
      font: inherit;
      padding: 11px 12px;
      min-height: 44px;
    }}
    input[type="file"] {{
      padding: 9px 12px;
      cursor: pointer;
    }}
    input:focus, select:focus, textarea:focus {{
      outline: 2px solid rgba(231, 71, 25, .3);
      border-color: var(--accent);
      background: #ffffff;
    }}
    textarea {{ min-height: 132px; resize: vertical; }}
    .full {{ grid-column: 1 / -1; }}
    .hint {{ color: var(--muted); font-size: .82rem; font-weight: 400; }}
    .upload-row {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 10px;
      align-items: end;
    }}
    .upload-link {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 44px;
      padding: 0 14px;
      border-radius: 6px;
      border: 1px solid var(--accent);
      color: #ffffff;
      background: var(--accent);
      font-weight: 700;
      text-decoration: none;
      white-space: nowrap;
    }}
    .upload-link:hover {{ background: var(--accent-dark); }}
    .rush-payment {{
      grid-column: 1 / -1;
      border: 1px solid #f3b08e;
      border-radius: 8px;
      background: #fff7ed;
      padding: 14px;
    }}
    .rush-payment strong {{
      display: block;
      color: var(--accent-dark);
      margin-bottom: 8px;
    }}
    .rush-payment .pay-button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      margin: 12px 0;
      padding: 0 14px;
      border-radius: 6px;
      background: var(--accent);
      color: #ffffff;
      font-weight: 800;
      text-decoration: none;
      text-transform: uppercase;
    }}
    .rush-payment label {{
      display: flex;
      align-items: center;
      gap: 10px;
      font-weight: 700;
    }}
    .rush-payment input[type="checkbox"] {{
      width: 18px;
      min-height: 18px;
    }}
    .actions {{
      display: flex;
      align-items: center;
      gap: 14px;
      margin-top: 24px;
      flex-wrap: wrap;
    }}
    button {{
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: white;
      font-weight: 700;
      text-transform: uppercase;
      padding: 12px 18px;
      min-height: 44px;
      cursor: pointer;
    }}
    button:hover {{ background: var(--accent-dark); }}
    .status {{
      padding: 12px 14px;
      border-radius: 6px;
      border: 1px solid var(--line);
      margin-bottom: 18px;
      color: var(--ok);
      background: #f0fdf4;
    }}
    .status.error {{
      color: var(--warn);
      background: #fff7ed;
      border-color: #fed7aa;
    }}
    .pricing-guide {{
      border-top: 1px solid var(--line);
      background: #fbfaf7;
    }}
    .pricing-guide summary {{
      display: flex;
      justify-content: space-between;
      gap: 16px;
      padding: 20px 26px;
      cursor: pointer;
      font-weight: 800;
      text-transform: uppercase;
      color: #161616;
      list-style: none;
    }}
    .pricing-guide summary::-webkit-details-marker {{ display: none; }}
    .pricing-guide summary span {{
      color: var(--accent);
      font-size: .9rem;
      white-space: nowrap;
    }}
    .pricing-body {{
      max-height: 520px;
      overflow: auto;
      padding: 0 26px 26px;
    }}
    .pricing-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }}
    .service-card {{
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #ffffff;
      padding: 16px;
    }}
    .service-card h3 {{
      margin: 0 0 10px;
      font-size: 1rem;
      text-transform: uppercase;
      color: #161616;
    }}
    .service-card h4 {{
      margin: 12px 0 8px;
      color: var(--accent-dark);
      font-size: .92rem;
    }}
    .service-card ul {{
      margin: 0;
      padding-left: 18px;
      color: #3f3f3f;
      line-height: 1.55;
      font-size: .9rem;
    }}
    .client-notice {{
      margin-top: 14px;
      border-left: 4px solid var(--accent);
      background: #fff7ed;
    }}
    @media (max-width: 900px) {{
      body {{ background-attachment: scroll; }}
      main {{
        grid-template-columns: 1fr;
        min-height: auto;
        padding: 20px 0;
      }}
      h1 {{ max-width: 11ch; }}
    }}
    @media (max-width: 760px) {{
      .grid {{ grid-template-columns: 1fr; }}
      .pricing-grid {{ grid-template-columns: 1fr; }}
      .upload-row {{ grid-template-columns: 1fr; }}
      main {{ width: min(100% - 18px, 1180px); }}
      form, .form-title {{ padding-left: 18px; padding-right: 18px; }}
      .form-title {{ display: block; }}
      .tag {{ display: block; margin-top: 8px; }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <span class="eyebrow">The Forsaken Society</span>
      <h1>Creative Request Portal</h1>
    </header>
    <section class="form-shell">
      <div class="form-title">
        <h2>Start a Design Request</h2>
        <div class="top-links">
          <span class="tag">NAXYSTUDIOS LLC</span>
          <a class="admin-login-link" href="/admin">Admin Login</a>
        </div>
      </div>
      {content}
    </section>
  </main>
  <script>
    const choices = {choices_json};
    function fillSelect(id, values, selected) {{
      const select = document.getElementById(id);
      select.innerHTML = "";
      values.forEach((value) => {{
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        if (value === selected) option.selected = true;
        select.appendChild(option);
      }});
    }}
    fillSelect("design_type", choices.designTypes, "AVIs");
    fillSelect("priority", choices.priorities, "Standard");
    fillSelect("rush_option", choices.rushOptions, "No Rush");

    const form = document.querySelector("form");
    const designTypeSelect = document.getElementById("design_type");
    const otherDesignWrap = document.getElementById("other_design_wrap");
    const otherDesignInput = document.getElementById("other_design_type");
    const rushSelect = document.getElementById("rush_option");
    const rushPayment = document.getElementById("rush_payment");
    const rushPayButton = document.getElementById("rush_pay_button");
    const rushPaymentText = document.getElementById("rush_payment_text");
    const rushPaymentConfirmed = document.getElementById("rush_payment_confirmed");
    const rushPaymentLink = document.getElementById("rush_payment_link");
    const fileInput = document.getElementById("upload_files");
    const filePayload = document.getElementById("uploaded_files_json");
    const submitButton = form.querySelector("button[type='submit']");
    const maxUploadBytes = 4 * 1024 * 1024;

    function syncOtherDesignType() {{
      const isOther = designTypeSelect.value === "Other";
      otherDesignWrap.classList.toggle("hidden", !isOther);
      otherDesignInput.required = isOther;
      if (!isOther) otherDesignInput.value = "";
    }}
    designTypeSelect.addEventListener("change", syncOtherDesignType);
    syncOtherDesignType();

    function syncRushPayment() {{
      const value = rushSelect.value;
      let link = "";
      let label = "";
      if (value.includes("24 Hour")) {{
        link = "https://square.link/u/Btov81yL";
        label = "Pay the 24 Hour Rush fee ($15) before submitting.";
      }} else if (value.includes("Same Day")) {{
        link = "https://square.link/u/TIl0ygTT";
        label = "Pay the Same Day Rush fee ($35) before submitting.";
      }}
      const needsPayment = Boolean(link);
      rushPayment.classList.toggle("hidden", !needsPayment);
      rushPayButton.href = link || "#";
      rushPayButton.textContent = value.includes("Same Day") ? "Pay Same Day Rush Fee" : "Pay 24 Hour Rush Fee";
      rushPaymentText.textContent = label;
      rushPaymentLink.value = link;
      rushPaymentConfirmed.required = needsPayment;
      if (!needsPayment) rushPaymentConfirmed.checked = false;
    }}
    rushSelect.addEventListener("change", syncRushPayment);
    syncRushPayment();

    function readFileAsPayload(file) {{
      return new Promise((resolve, reject) => {{
        const reader = new FileReader();
        reader.onload = () => {{
          const dataUrl = String(reader.result || "");
          resolve({{
            name: file.name,
            type: file.type || "application/octet-stream",
            data: dataUrl.split(",", 2)[1] || ""
          }});
        }};
        reader.onerror = () => reject(reader.error || new Error("Could not read file"));
        reader.readAsDataURL(file);
      }});
    }}

    form.addEventListener("submit", async (event) => {{
      const files = Array.from(fileInput.files || []);
      const totalBytes = files.reduce((sum, file) => sum + file.size, 0);
      if (totalBytes > maxUploadBytes) {{
        event.preventDefault();
        alert("Uploads must be 4 MB total or less. For larger videos, paste a share link instead.");
        return;
      }}
      if (!rushPayment.classList.contains("hidden") && !rushPaymentConfirmed.checked) {{
        event.preventDefault();
        alert("Please pay the selected rush fee and check the payment confirmation box before submitting.");
        return;
      }}
      if (!files.length) return;
      event.preventDefault();
      submitButton.disabled = true;
      submitButton.textContent = "Uploading...";
      try {{
        filePayload.value = JSON.stringify(await Promise.all(files.map(readFileAsPayload)));
        form.submit();
      }} catch (error) {{
        alert("Could not prepare the upload. Please try again or paste a share link.");
        submitButton.disabled = false;
        submitButton.textContent = "Submit Request";
      }}
    }});
  </script>
</body>
</html>"""
    return html_doc.encode("utf-8")


def pricing_guide_html() -> str:
    return """
    <details class="pricing-guide">
      <summary>NAXYSTUDIOS LLC - Service Menu & Pricing Guide <span>View pricing</span></summary>
      <div class="pricing-body">
        <div class="pricing-grid">
          <section class="service-card client-notice full">
            <h3>All prices for design work</h3>
            <ul>
              <li>Rush Orders: 24 Hour Rush adds a flat $15 fee.</li>
              <li>Same Day Rush adds a flat $35 fee.</li>
              <li>Pricing may vary depending on complexity, revisions, licensing, and turnaround time.</li>
            </ul>
          </section>
          <section class="service-card">
            <h3>Social Media Revamp</h3>
            <ul>
              <li>Full Social Revamp - $65-$190: revamp multiple platforms like Twitter, YouTube, Twitch, and more</li>
              <li>Social Revamp (Single Platform) - $35-$65: revamp one specific social platform of your choice</li>
            </ul>
          </section>
          <section class="service-card">
            <h3>Stream Packages</h3>
            <ul>
              <li>Stream Package - $70-$250</li>
              <li>Includes overlays for camera and stream, stream starting/ended panels, BRB panel, and custom emotes</li>
              <li>Animated Stream Package - $85-$350</li>
              <li>Includes animated overlays for camera and stream, stream starting/ended panels, BRB panel, and custom emotes</li>
            </ul>
          </section>
          <section class="service-card">
            <h3>Graphics and Visuals</h3>
            <ul>
              <li>Custom 3D Intro/Outro - $95-$170</li>
              <li>Custom Merch Designs - $55-$230 (Non-NSFW designs)</li>
              <li>Advertisements - $40-$70</li>
              <li>Poster Advertisements - $60-$300</li>
              <li>Custom Designs (NFTs, Album Covers) - $40-$220</li>
              <li>Logo Designs - $75-$350</li>
              <li>Flyer Design - $40-$100</li>
              <li>Yard Signs Designs - $40-$150</li>
              <li>UI Website Design - $180 (Square Up or Wix)</li>
              <li>Thumbnail Designs - $15-$40</li>
              <li>Profile Avatars/Profile Photos - $5-$20</li>
              <li>Custom Sticker Designs - $10-$50</li>
              <li>Business Card Designs - $50-$200</li>
            </ul>
          </section>
          <section class="service-card client-notice">
            <h3>NSFW Graphics and Visuals</h3>
            <h4>Must be 18+ to order NSFW services.</h4>
            <ul>
              <li>Custom NSFW Merch Designs - $85-$350</li>
              <li>Custom NSFW Characters Hand Drawn Designs - $90-$500</li>
              <li>NSFW Thumbnails - $40-$150</li>
              <li>Custom NSFW Sticker Designs - $30-$80</li>
            </ul>
          </section>
          <section class="service-card">
            <h3>Tattoo Designs</h3>
            <ul>
              <li>Tattoo Designs - $50-$500</li>
            </ul>
          </section>
          <section class="service-card client-notice">
            <h3>3D Characters</h3>
            <h4>Must be 18+ to order NSFW services.</h4>
            <ul>
              <li>Custom 3D Characters (Rigged + Posed) - $90</li>
              <li>Custom 3D Characters (Rigged + Animated) - $250</li>
              <li>Custom/NSFW Merch Designs with 3D Render Concepts - $100</li>
              <li>Custom NSFW 3D Characters (Rigged + Posed) - $150</li>
              <li>Custom NSFW 3D Characters (Rigged + Animated) - $300</li>
            </ul>
          </section>
        </div>
      </div>
    </details>
    """

def form_html(status: str = "", error: bool = False) -> str:
    status_html = f'<div class="status{" error" if error else ""}">{html.escape(status)}</div>' if status else ""
    return f"""
    <form method="post" action="/submit">
      {status_html}
      <div class="grid">
        <label>Member Name
          <input name="member_name" autocomplete="name" required>
        </label>
        <label>Email
          <input name="email" type="email" autocomplete="email" required>
        </label>
        <label>Project Name
          <input name="project_name" required>
        </label>
        <label>Requested Deadline
          <input name="requested_deadline" type="date" required>
        </label>
        <label>Design Type
          <select id="design_type" name="design_type" required></select>
        </label>
        <label id="other_design_wrap" class="hidden">Tell Us What You Need
          <input id="other_design_type" name="other_design_type" placeholder="Describe the custom service request">
        </label>
        <label>Priority Level
          <select id="priority" name="priority" required></select>
        </label>
        <label>Expedited Option
          <select id="rush_option" name="rush_option" required></select>
        </label>
        <div id="rush_payment" class="rush-payment hidden">
          <strong>Rush Fee Payment Required</strong>
          <p id="rush_payment_text"></p>
          <a id="rush_pay_button" class="pay-button" href="#" target="_blank" rel="noopener">Pay Rush Fee</a>
          <label>
            <input id="rush_payment_confirmed" name="rush_payment_confirmed" type="checkbox" value="Client confirmed rush fee payment before submitting">
            I paid the selected rush fee before submitting.
          </label>
          <input id="rush_payment_link" name="rush_payment_link" type="hidden">
        </div>
        <div class="full upload-row">
          <label>File or URL Link
            <input name="uploaded_files_link" type="url" placeholder="Paste Google Drive, Dropbox, Canva, or website link">
          </label>
          <label>Upload Photos or Videos
            <input id="upload_files" name="upload_files" type="file" accept="image/*,video/*" multiple>
          </label>
          <input id="uploaded_files_json" name="uploaded_files_json" type="hidden">
        </div>
        <label class="full">Design Description
          <textarea name="description" required></textarea>
        </label>
        <label class="full">Notes for Designer
          <textarea name="notes"></textarea>
        </label>
      </div>
      <div class="actions">
        <button type="submit">Submit Request</button>
      </div>
    </form>
    {pricing_guide_html()}
    """


def option_tags(options: list[str], selected: str) -> str:
    return "".join(
        f'<option value="{html.escape(option, quote=True)}"{" selected" if option == selected else ""}>{html.escape(option)}</option>'
        for option in options
    )


def admin_session_token(config: dict[str, str]) -> str:
    seed = f"{config['admin_dashboard_password']}:{config['google_apps_script_secret']}"
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()


def admin_login_html(status: str = "") -> bytes:
    status_html = f'<div class="admin-alert">{html.escape(status)}</div>' if status else ""
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Forsaken Admin Dashboard</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      font-family: Arial, Helvetica, sans-serif;
      background: #111 url("/static/forsaken-background.png") center/cover fixed;
      color: #171717;
    }}
    body::before {{ content: ""; position: fixed; inset: 0; background: rgba(0,0,0,.72); }}
    form {{
      position: relative;
      width: min(420px, calc(100% - 28px));
      background: rgba(255,255,255,.97);
      border-radius: 8px;
      padding: 26px;
      box-shadow: 0 20px 60px rgba(0,0,0,.42);
    }}
    h1 {{ margin: 0 0 18px; text-transform: uppercase; font-size: 1.45rem; }}
    label {{ display: grid; gap: 8px; font-weight: 800; }}
    input {{
      min-height: 44px;
      border: 1px solid #d8d1c8;
      border-radius: 6px;
      padding: 10px 12px;
      font: inherit;
      background: #f8f5ef;
    }}
    button {{
      width: 100%;
      min-height: 44px;
      margin-top: 18px;
      border: 0;
      border-radius: 6px;
      background: #e74719;
      color: #fff;
      font-weight: 900;
      text-transform: uppercase;
      cursor: pointer;
    }}
    .admin-alert {{ margin-bottom: 14px; color: #a24112; font-weight: 700; }}
  </style>
</head>
<body>
  <form method="post" action="/admin/login">
    <h1>Admin Dashboard</h1>
    {status_html}
    <label>Password
      <input name="password" type="password" autocomplete="current-password" required autofocus>
    </label>
    <button type="submit">Open Dashboard</button>
  </form>
</body>
</html>""".encode("utf-8")


def admin_dashboard_html(requests: list[dict[str, Any]], status: str = "", error: bool = False) -> bytes:
    total = len(requests)
    unseen = sum(1 for item in requests if item.get("seen_status") != "Seen")
    pending = sum(1 for item in requests if item.get("approval_status") in ("", "Pending Review"))
    active = sum(1 for item in requests if item.get("request_status") in ("Submitted", "In Progress", "Revision"))
    status_class = " admin-alert-error" if error else ""
    status_html = f'<div class="admin-alert{status_class}">{html.escape(status)}</div>' if status else ""

    def text(item: dict[str, Any], key: str) -> str:
        return html.escape(cell_text(item.get(key)))

    rows = []
    for item in requests:
        request_id = cell_text(item.get("request_id"))
        asset_link = cell_text(item.get("uploaded_files_link"))
        final_link = cell_text(item.get("final_deliverables_link"))
        asset_html = f'<a href="{html.escape(asset_link, quote=True)}" target="_blank" rel="noopener">Open assets</a>' if asset_link else ""
        final_html = f'<a href="{html.escape(final_link, quote=True)}" target="_blank" rel="noopener">Open final</a>' if final_link else ""
        rows.append(f"""
        <tr>
          <td>
            <strong>{html.escape(request_id)}</strong>
            <span>{text(item, "submitted_at")}</span>
          </td>
          <td>
            <strong>{text(item, "project_name")}</strong>
            <span>{text(item, "member_name")} · {text(item, "email")}</span>
          </td>
          <td>
            <strong>{text(item, "design_type")}</strong>
            <span>{text(item, "rush_option")} · {text(item, "priority")}</span>
          </td>
          <td>
            <form class="request-form" method="post" action="/admin/update">
              <input type="hidden" name="request_id" value="{html.escape(request_id, quote=True)}">
              <div class="control-grid">
                <label>Seen
                  <select name="seen_status">{option_tags(SEEN_STATUS_OPTIONS, cell_text(item.get("seen_status")) or "Not Seen")}</select>
                </label>
                <label>Approval
                  <select name="approval_status">{option_tags(APPROVAL_STATUS_OPTIONS, cell_text(item.get("approval_status")) or "Pending Review")}</select>
                </label>
                <label>Request Status
                  <select name="request_status">{option_tags(ADMIN_STATUS_OPTIONS, cell_text(item.get("request_status")) or "Submitted")}</select>
                </label>
                <label>Assigned Designer
                  <input name="assigned_designer" value="{text(item, "assigned_designer")}" placeholder="Designer name">
                </label>
                <label>Seen By
                  <input name="seen_by" value="{text(item, "seen_by")}" placeholder="Admin name">
                </label>
                <label>Start Date
                  <input name="design_start_date" type="date" value="{text(item, "design_start_date")}">
                </label>
                <label>Due Date
                  <input name="design_due_date" type="date" value="{text(item, "design_due_date")}">
                </label>
                <label>Final Link
                  <input name="final_deliverables_link" type="url" value="{text(item, "final_deliverables_link")}" placeholder="https://...">
                </label>
                <label class="wide">Admin Notes
                  <textarea name="admin_notes">{text(item, "admin_notes")}</textarea>
                </label>
              </div>
              <div class="request-links">{asset_html}{final_html}</div>
              <details>
                <summary>Request details</summary>
                <p>{text(item, "description")}</p>
                <p>{text(item, "admin_notes")}</p>
              </details>
              <button type="submit">Save</button>
            </form>
            <form class="delete-form" method="post" action="/admin/delete" onsubmit="return confirm('Delete {html.escape(request_id, quote=True)} from the admin dashboard? This cannot be undone from the site.');">
              <input type="hidden" name="request_id" value="{html.escape(request_id, quote=True)}">
              <button type="submit">Delete</button>
            </form>
          </td>
        </tr>
        """)

    table_body = "\n".join(rows) if rows else '<tr><td colspan="4" class="empty">No requests found yet.</td></tr>'
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Forsaken Admin Dashboard</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      background: #111 url("/static/forsaken-background.png") center/cover fixed;
      color: #171717;
    }}
    body::before {{
      content: "";
      position: fixed;
      inset: 0;
      background: linear-gradient(90deg, rgba(0,0,0,.88), rgba(0,0,0,.62));
      pointer-events: none;
    }}
    main {{
      position: relative;
      z-index: 1;
      width: min(1440px, calc(100% - 28px));
      margin: 0 auto;
      padding: 28px 0 42px;
    }}
    header {{
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: end;
      color: #fff;
      margin-bottom: 18px;
    }}
    h1 {{ margin: 0; text-transform: uppercase; font-size: clamp(2rem, 5vw, 4rem); line-height: .92; }}
    .top-actions {{ display: flex; gap: 10px; align-items: center; }}
    .top-actions a, .top-actions button {{
      min-height: 40px;
      border: 1px solid rgba(255,255,255,.42);
      border-radius: 6px;
      background: rgba(255,255,255,.12);
      color: #fff;
      padding: 9px 12px;
      text-decoration: none;
      font-weight: 800;
      cursor: pointer;
    }}
    .cards {{
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }}
    .metric {{
      background: rgba(255,255,255,.96);
      border-radius: 8px;
      padding: 16px;
      border-left: 4px solid #e74719;
    }}
    .metric span {{ display: block; color: #6d6d6d; font-size: .8rem; font-weight: 800; text-transform: uppercase; }}
    .metric strong {{ display: block; margin-top: 6px; font-size: 2rem; }}
    .admin-alert {{
      background: #f0fdf4;
      border: 1px solid #bbf7d0;
      color: #166534;
      border-radius: 8px;
      padding: 12px 14px;
      margin-bottom: 16px;
      font-weight: 800;
    }}
    .admin-alert-error {{ background: #fff7ed; border-color: #fed7aa; color: #a24112; }}
    .table-wrap {{
      background: rgba(255,255,255,.97);
      border-radius: 8px;
      overflow: auto;
      box-shadow: 0 20px 70px rgba(0,0,0,.36);
    }}
    table {{ width: 100%; border-collapse: collapse; min-width: 1120px; }}
    th {{
      position: sticky;
      top: 0;
      background: #161616;
      color: #fff;
      text-align: left;
      padding: 12px;
      text-transform: uppercase;
      font-size: .78rem;
      z-index: 2;
    }}
    td {{ vertical-align: top; border-top: 1px solid #e4e0da; padding: 12px; }}
    td > span, td strong + span {{ display: block; color: #6d6d6d; margin-top: 4px; font-size: .86rem; }}
    .request-form {{ min-width: 640px; }}
    .control-grid {{ display: grid; grid-template-columns: repeat(4, minmax(130px, 1fr)); gap: 10px; }}
    label {{ display: grid; gap: 5px; color: #333; font-size: .78rem; font-weight: 900; text-transform: uppercase; }}
    input, select, textarea {{
      width: 100%;
      min-height: 38px;
      border: 1px solid #d4cec4;
      border-radius: 6px;
      background: #f8f5ef;
      padding: 8px 9px;
      font: inherit;
      text-transform: none;
    }}
    textarea {{ min-height: 70px; resize: vertical; }}
    .wide {{ grid-column: 1 / -1; }}
    .request-links {{ display: flex; gap: 12px; margin-top: 10px; flex-wrap: wrap; }}
    .request-links a {{ color: #b93412; font-weight: 800; }}
    details {{ margin-top: 10px; color: #4b4b4b; }}
    details p {{ white-space: pre-wrap; line-height: 1.45; }}
    .request-form button {{
      margin-top: 10px;
      min-height: 38px;
      border: 0;
      border-radius: 6px;
      background: #e74719;
      color: #fff;
      font-weight: 900;
      text-transform: uppercase;
      padding: 8px 14px;
      cursor: pointer;
    }}
    .delete-form {{ display: inline-block; margin-left: 8px; }}
    .delete-form button {{
      margin-top: 10px;
      min-height: 38px;
      border: 1px solid #991b1b;
      border-radius: 6px;
      background: #fff1f2;
      color: #991b1b;
      font-weight: 900;
      text-transform: uppercase;
      padding: 8px 14px;
      cursor: pointer;
    }}
    .empty {{ text-align: center; color: #6d6d6d; padding: 36px; }}
    @media (max-width: 900px) {{
      header {{ display: block; }}
      .top-actions {{ margin-top: 14px; }}
      .cards {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
    @media (max-width: 620px) {{
      .cards {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <span>NAXYSTUDIOS LLC</span>
        <h1>Request Admin</h1>
      </div>
      <div class="top-actions">
        <a href="/">Form</a>
        <form method="post" action="/admin/logout"><button type="submit">Logout</button></form>
      </div>
    </header>
    {status_html}
    <section class="cards">
      <div class="metric"><span>Total Requests</span><strong>{total}</strong></div>
      <div class="metric"><span>Unseen</span><strong>{unseen}</strong></div>
      <div class="metric"><span>Pending Approval</span><strong>{pending}</strong></div>
      <div class="metric"><span>Active</span><strong>{active}</strong></div>
    </section>
    <section class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Request</th>
            <th>Client</th>
            <th>Project</th>
            <th>Manage</th>
          </tr>
        </thead>
        <tbody>{table_body}</tbody>
      </table>
    </section>
  </main>
</body>
</html>""".encode("utf-8")


class RequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        print(f"{self.address_string()} - {format % args}")

    def is_admin_authorized(self) -> bool:
        config = get_config()
        password = config["admin_dashboard_password"]
        if not password:
            return False
        cookie = self.headers.get("Cookie", "")
        expected = admin_session_token(config)
        return any(part.strip() == f"{ADMIN_COOKIE_NAME}={expected}" for part in cookie.split(";"))

    def redirect(self, location: str, headers: dict[str, str] | None = None) -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()

    def send_static(self) -> None:
        raw_path = unquote(self.path.split("?", 1)[0])
        relative = raw_path.removeprefix("/static/").replace("/", os.sep)
        static_root = (BASE_DIR / "static").resolve()
        file_path = (static_root / relative).resolve()
        if not str(file_path).startswith(str(static_root)) or not file_path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = file_path.read_bytes()
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "public, max-age=3600")
        self.end_headers()
        self.wfile.write(body)

    def send_html(self, body: bytes, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, payload: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if self.path.startswith("/static/"):
            self.send_static()
            return
        if path == "/health":
            config = get_config()
            self.send_json({
                "ok": True,
                "google_sheet_id": config["google_sheet_id"],
                "google_sheet_url": config["google_sheet_url"],
                "google_sheet_configured": bool(config["google_service_account_json"]),
                "apps_script_webhook_configured": bool(config["google_apps_script_webhook_url"]),
                "email_configured": bool(config["admin_email"] and config["smtp_host"] and config["smtp_from"]),
                "admin_dashboard_configured": bool(config["admin_dashboard_password"]),
            })
            return
        if path == "/admin":
            config = get_config()
            if not config["admin_dashboard_password"]:
                self.send_html(admin_login_html("Admin password is not configured yet."), HTTPStatus.SERVICE_UNAVAILABLE)
                return
            if not self.is_admin_authorized():
                self.send_html(admin_login_html())
                return
            try:
                requests = fetch_admin_requests(config)
                self.send_html(admin_dashboard_html(requests))
            except Exception as exc:
                self.send_html(admin_dashboard_html([], str(exc), error=True), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        self.send_html(page_template(form_html()))

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/admin/login":
            length = int(self.headers.get("Content-Length", "0"))
            data = parse_form(self.rfile.read(length))
            config = get_config()
            if config["admin_dashboard_password"] and data.get("password") == config["admin_dashboard_password"]:
                cookie = f"{ADMIN_COOKIE_NAME}={admin_session_token(config)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=604800"
                self.redirect("/admin", {"Set-Cookie": cookie})
                return
            self.send_html(admin_login_html("Wrong password."), HTTPStatus.UNAUTHORIZED)
            return
        if path == "/admin/logout":
            self.redirect("/admin", {"Set-Cookie": f"{ADMIN_COOKIE_NAME}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0"})
            return
        if path == "/admin/update":
            if not self.is_admin_authorized():
                self.redirect("/admin")
                return
            length = int(self.headers.get("Content-Length", "0"))
            data = parse_form(self.rfile.read(length))
            config = get_config()
            try:
                update_admin_request(data, config)
                requests = fetch_admin_requests(config)
                self.send_html(admin_dashboard_html(requests, f"{data.get('request_id', 'Request')} updated."))
            except Exception as exc:
                try:
                    requests = fetch_admin_requests(config)
                except Exception:
                    requests = []
                self.send_html(admin_dashboard_html(requests, str(exc), error=True), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/admin/delete":
            if not self.is_admin_authorized():
                self.redirect("/admin")
                return
            length = int(self.headers.get("Content-Length", "0"))
            data = parse_form(self.rfile.read(length))
            config = get_config()
            request_id = data.get("request_id", "Request")
            try:
                delete_admin_request(data, config)
                requests = fetch_admin_requests(config)
                self.send_html(admin_dashboard_html(requests, f"{request_id} deleted."))
            except Exception as exc:
                try:
                    requests = fetch_admin_requests(config)
                except Exception:
                    requests = []
                self.send_html(admin_dashboard_html(requests, str(exc), error=True), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path != "/submit":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        length = int(self.headers.get("Content-Length", "0"))
        data = parse_form(self.rfile.read(length))
        errors = validate_form(data)
        if errors:
            self.send_html(page_template(form_html(" ".join(errors), error=True)), HTTPStatus.BAD_REQUEST)
            return
        try:
            result = save_submission(data)
            try:
                send_email(data, result)
            except Exception as email_error:
                print(f"Email could not be sent: {email_error}")
            self.send_html(page_template(form_html("I got your request")))
        except Exception as exc:
            self.send_html(page_template(form_html(str(exc), error=True)), HTTPStatus.INTERNAL_SERVER_ERROR)


def main() -> None:
    config = get_config()
    server = ThreadingHTTPServer((config["host"], int(config["port"])), RequestHandler)
    print(f"Creative request portal running at http://{config['host']}:{config['port']}")
    server.serve_forever()


handler = RequestHandler


if __name__ == "__main__":
    main()

