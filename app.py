from __future__ import annotations

import base64
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import hmac
import html
import hashlib
import json
import mimetypes
import os
import re
import smtplib
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from email.message import EmailMessage
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlencode, unquote, urlparse


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_GOOGLE_SHEET_ID = "1FRv0kfJc10hZLygUnApg5kbbFQ1Bnvxt2aE3-HQn5jQ"
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1FRv0kfJc10hZLygUnApg5kbbFQ1Bnvxt2aE3-HQn5jQ/edit?usp=drive_link"
ENV_PATH = BASE_DIR / ".env"

TRACKER_SHEET = "Request Tracker"
DEFAULT_UPLOAD_FOLDER = "https://drive.google.com/drive/folders/17Elym_RLgFLL2EPOOgS2FKhwNA3ikd-f?usp=sharing"
ADMIN_COOKIE_NAME = "forsaken_admin"
APPS_SCRIPT_NOT_CONFIGURED_MESSAGE = (
    "Google Apps Script webhook is not configured. Add GOOGLE_APPS_SCRIPT_WEBHOOK_URL "
    "to .env to sync Google Sheets data. Showing the local preview without synced requests."
)

ADMIN_STATUS_OPTIONS = ["Pending Payment", "Paid", "Submitted", "In Progress", "Revision", "Completed", "Cancelled", "Deleted"]
SEEN_STATUS_OPTIONS = ["Not Seen", "Seen"]
APPROVAL_STATUS_OPTIONS = ["Pending Review", "Approved", "Not Approved", "Needs Info"]
PROGRESS_STATES = {
    "not_started": ("Not Started", "red"),
    "in_process": ("In Process", "yellow"),
    "done": ("Done", "green"),
}

SERVICE_PRICES = {
    "AVIs": {"label": "AVI / Profile Picture", "display": "Free", "min": 0, "max": 0},
    "Twitter Headers": {"label": "Twitter / X Header", "display": "$25", "min": 25, "max": 25},
    "YouTube Headers": {"label": "YouTube Banner", "display": "$15", "min": 15, "max": 15},
    "TikTok Banners": {"label": "TikTok Cover Graphic", "display": "$15", "min": 15, "max": 15},
    "Social Media Packages": {"label": "Social Media Package", "display": "$100", "min": 100, "max": 100},
    "3D Animations": {"label": "Custom 3D Intro / Outro", "display": "$95-$170", "min": 95, "max": 170},
    "Intros or Outros": {"label": "Intro or Outro", "display": "$150", "min": 150, "max": 150},
    "Simple Editing": {"label": "Simple Editing", "display": "$3/clip", "min": 3, "max": 3, "unit": "clip"},
    "Boss Editing": {"label": "Boss Editing", "display": "$8/clip", "min": 8, "max": 8, "unit": "clip"},
    "Other": {"label": "Other", "display": "Free", "min": 0, "max": 0},
}

RUSH_PRICES = {
    "No Rush": {"label": "No Rush", "display": "$0", "min": 0, "max": 0},
    "Rush Order +$20": {"label": "Rush Order Fee", "display": "$20", "min": 20, "max": 20},
}

SQUARE_API_HOSTS = {
    "sandbox": "https://connect.squareupsandbox.com",
    "production": "https://connect.squareup.com",
}

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
    "design_type",
    "description",
    "priority",
    "rush_option",
]

SIMPLE_REQUIRED_FIELDS = [
    "member_name",
    "email",
    "design_type",
    "description",
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
        "base_url": env.get("BASE_URL", "https://forsakensociety.vercel.app").rstrip("/"),
        "square_env": env.get("SQUARE_ENV", "sandbox").lower(),
        "square_access_token": env.get("SQUARE_ACCESS_TOKEN", ""),
        "square_location_id": env.get("SQUARE_LOCATION_ID", ""),
        "square_webhook_signature_key": env.get("SQUARE_WEBHOOK_SIGNATURE_KEY", ""),
        "host": env.get("HOST", "127.0.0.1"),
        "port": env.get("PORT", "8000"),
    }


def cell_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def request_progress_state(item: dict[str, Any]) -> tuple[str, str]:
    request_status = cell_text(item.get("request_status"))
    approval_status = cell_text(item.get("approval_status"))
    final_link = cell_text(item.get("final_deliverables_link"))
    if request_status == "Completed" or final_link:
        return PROGRESS_STATES["done"]
    if request_status in ("In Progress", "Revision") or approval_status == "Approved":
        return PROGRESS_STATES["in_process"]
    return PROGRESS_STATES["not_started"]


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
            "Boss Editing",
            "Other",
        ],
        "priorities": ["Standard", "Expedited"],
        "rushOptions": list(RUSH_PRICES.keys()),
        "servicePrices": SERVICE_PRICES,
        "rushPrices": RUSH_PRICES,
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


def validate_simple_form(data: dict[str, str]) -> list[str]:
    errors = []
    for field in SIMPLE_REQUIRED_FIELDS:
        if not data.get(field):
            errors.append(f"{FIELD_LABELS[field]} is required.")
    if data.get("email") and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", data["email"]):
        errors.append("Enter a valid email address.")
    return errors


def normalize_simple_form(data: dict[str, str]) -> dict[str, str]:
    normalized = {key: value for key, value in data.items() if not key.startswith("_") and not key.startswith("square_")}
    design_type = normalized.get("design_type", "").strip()
    normalized["project_name"] = normalized.get("project_name", "").strip() or f"{design_type or 'Design'} Request"
    normalized["requested_deadline"] = normalized.get("requested_deadline", "").strip() or "Flexible"
    normalized["priority"] = normalized.get("priority", "").strip() or "Standard"
    normalized["rush_option"] = "Rush Order +$20" if normalized.get("rush_requested") else "No Rush"
    return normalized


def normalize_request_form(data: dict[str, str]) -> dict[str, str]:
    normalized = {key: value for key, value in data.items() if not key.startswith("_") and not key.startswith("square_")}
    design_type = normalized.get("design_type", "").strip()
    normalized["project_name"] = normalized.get("project_name", "").strip() or f"{design_type or 'Design'} Request"
    normalized["requested_deadline"] = normalized.get("requested_deadline", "").strip() or "Flexible"
    normalized["priority"] = normalized.get("priority", "").strip() or "Standard"
    if normalized.get("rush_requested"):
        normalized["rush_option"] = "Rush Order +$20"
    else:
        normalized["rush_option"] = normalized.get("rush_option", "").strip() or "No Rush"
    return normalized


def rush_fee_text(rush_option: str) -> str:
    if "Rush" in rush_option:
        return "$20 rush fee"
    if "Same Day" in rush_option:
        return "$35 rush fee"
    if "24 Hour" in rush_option:
        return "$15 rush fee"
    return "No additional fee"


def combined_admin_notes(data: dict[str, str]) -> str:
    parts = []
    if data.get("notes"):
        parts.append(f"Notes: {data['notes']}")
    elif data.get("admin_notes"):
        parts.append(cell_text(data.get("admin_notes")))
    if data.get("other_design_type"):
        parts.append(f"Other design type: {data['other_design_type']}")
    if data.get("custom_cart_items"):
        try:
            custom_items = json.loads(data["custom_cart_items"])
        except json.JSONDecodeError:
            custom_items = []
        lines = []
        for item in custom_items:
            name = cell_text(item.get("name"))
            price = cell_text(item.get("price"))
            if name:
                if price and not price.startswith("$"):
                    price = f"${price}"
                lines.append(f"- {name}: {price or '$0'}")
        if lines:
            parts.append("Custom cart items:\n" + "\n".join(lines))
    payment_status = data.get("_payment_status")
    if payment_status:
        parts.append(f"Payment status: {payment_status}")
    if data.get("_payment_total_display"):
        parts.append(f"Square checkout total: {data['_payment_total_display']}")
    if data.get("_payment_total_cents"):
        parts.append(f"Square checkout total cents: {data['_payment_total_cents']}")
    if data.get("_rush_fee_included"):
        parts.append(f"Square rush fee included: {data['_rush_fee_included']}")
    if data.get("_cart_summary"):
        parts.append(f"Cart summary:\n{data['_cart_summary']}")
    if data.get("_square_payment_link_id"):
        parts.append(f"Square payment link ID: {data['_square_payment_link_id']}")
    if data.get("_square_order_id"):
        parts.append(f"Square order ID: {data['_square_order_id']}")
    if data.get("_square_checkout_url"):
        parts.append(f"Square checkout URL: {data['_square_checkout_url']}")
    if data.get("_square_payment_id"):
        parts.append(f"Square payment ID: {data['_square_payment_id']}")
    if data.get("_square_payment_status"):
        parts.append(f"Square payment status: {data['_square_payment_status']}")
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
        data.get("_request_status", "Submitted"),
        "Not Seen",
        "",
        data.get("_approval_status", "Pending Review"),
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
    files = json.loads(data.get("uploaded_files_json") or "[]")
    payload = {
        "secret": config["google_apps_script_secret"],
        "submitted_at": now.strftime("%Y-%m-%d %I:%M %p"),
        "values": build_tracker_values(data, "", now, config, as_text=True),
        "fields": data,
        "files": files,
    }

    def send_payload(current_payload: dict[str, Any]) -> dict[str, Any]:
        body = json.dumps(current_payload).encode("utf-8")
        request = urllib.request.Request(
            config["google_apps_script_webhook_url"],
            data=body,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=55) as response:
                return json.loads(response.read().decode("utf-8") or "{}")
        except urllib.error.HTTPError as exc:
            message = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Google Apps Script rejected the request: {message}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Could not reach the Google Apps Script webhook: {exc.reason}") from exc

    response_payload = send_payload(payload)
    response_error = cell_text(response_payload.get("error"))
    is_drive_upload_error = (
        response_error
        and files
        and ("DriveApp.getFolderById" in response_error or "googleapis.com/auth/drive" in response_error)
    )

    if is_drive_upload_error:
        retry_payload = dict(payload)
        retry_values = list(payload["values"])
        upload_names = ", ".join(cell_text(file.get("name")) for file in files if cell_text(file.get("name")))
        fallback_note = "Upload files were not saved because Google Drive upload permission needs to be authorized in Apps Script."
        if upload_names:
            fallback_note += f" Attempted files: {upload_names}."
        retry_values[23] = "\n".join(value for value in [cell_text(retry_values[23]), fallback_note] if value)
        retry_payload["values"] = retry_values
        retry_payload["files"] = []
        response_payload = send_payload(retry_payload)

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
    if not config["google_apps_script_webhook_url"]:
        return []
    payload = get_apps_script({"action": "listRequests"}, config)
    if "requests" not in payload:
        raise RuntimeError("Redeploy the Google Apps Script web app so the admin dashboard can read requests.")
    requests = payload.get("requests", [])
    if not isinstance(requests, list):
        return []
    return [item for item in requests if cell_text(item.get("request_status")) != "Deleted"]


def fetch_designers(config: dict[str, str]) -> list[dict[str, str]]:
    if not config["google_apps_script_webhook_url"]:
        return []
    try:
        payload = get_apps_script({"action": "listDesigners"}, config)
    except Exception:
        return []
    designers = payload.get("designers", [])
    if not isinstance(designers, list):
        return []
    return [
        {
            "name": cell_text(item.get("name")),
            "email": cell_text(item.get("email")),
        }
        for item in designers
        if cell_text(item.get("name")) and cell_text(item.get("email"))
    ]


def add_designer(data: dict[str, str], config: dict[str, str]) -> None:
    version_payload = get_apps_script({"action": "version"}, config)
    if not version_payload.get("supports_designers"):
        raise RuntimeError("Redeploy the Google Apps Script web app so designer management can save.")
    call_apps_script(
        {
            "action": "addDesigner",
            "designer": {
                "name": data.get("designer_name", ""),
                "email": data.get("designer_email", ""),
            },
        },
        config,
    )


def delete_designer(data: dict[str, str], config: dict[str, str]) -> None:
    version_payload = get_apps_script({"action": "version"}, config)
    if not version_payload.get("supports_designers"):
        raise RuntimeError("Redeploy the Google Apps Script web app so designer management can save.")
    call_apps_script(
        {
            "action": "deleteDesigner",
            "email": data.get("designer_email", ""),
        },
        config,
    )


def update_admin_request(data: dict[str, str], config: dict[str, str]) -> None:
    version_payload = get_apps_script({"action": "version"}, config)
    if not version_payload.get("admin_dashboard"):
        raise RuntimeError("Redeploy the Google Apps Script web app so the admin dashboard can save updates.")
    updates = {
        "request_status": data.get("request_status", ""),
        "seen_status": data.get("seen_status", ""),
        "seen_by": data.get("seen_by", ""),
        "approval_status": data.get("approval_status", ""),
        "assigned_designer": data.get("assigned_designer", ""),
        "final_deliverables_link": data.get("final_deliverables_link", ""),
        "admin_notes": data.get("admin_notes", ""),
    }
    if "design_start_date" in data:
        updates["design_start_date"] = data.get("design_start_date", "")
    if "design_due_date" in data:
        updates["design_due_date"] = data.get("design_due_date", "")
    call_apps_script(
        {
            "action": "updateRequest",
            "request_id": data.get("request_id", ""),
            "updates": updates,
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


def decimal_to_cents(value: Any) -> int:
    try:
        amount = Decimal(str(value or "0")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return 0
    return max(0, int(amount * 100))


def cents_to_money(cents: int) -> str:
    dollars = Decimal(cents) / Decimal(100)
    return f"${dollars:,.2f}"


def money_label_from_cents(cents: int) -> str:
    dollars = Decimal(cents) / Decimal(100)
    if dollars == dollars.to_integral():
        return f"${int(dollars):,}"
    return f"${dollars:,.2f}"


def parse_custom_cart_items(value: str) -> list[dict[str, Any]]:
    if not value:
        return []
    try:
        items = json.loads(value)
    except json.JSONDecodeError:
        return []
    if not isinstance(items, list):
        return []
    parsed = []
    for item in items:
        if not isinstance(item, dict):
            continue
        name = cell_text(item.get("name"))
        if not name:
            continue
        cents = decimal_to_cents(item.get("price"))
        parsed.append({
            "name": name[:80],
            "price_cents": cents,
            "price": cents_to_money(cents),
        })
    return parsed


def parse_design_cart_items(value: str) -> list[dict[str, Any]]:
    if not value:
        return []
    try:
        items = json.loads(value)
    except json.JSONDecodeError:
        return []
    if not isinstance(items, list):
        return []
    parsed = []
    for item in items:
        if not isinstance(item, dict):
            continue
        design_type = cell_text(item.get("design_type"))
        if design_type not in SERVICE_PRICES:
            continue
        service = SERVICE_PRICES[design_type]
        quantity = 1
        if service.get("unit") == "clip":
            try:
                quantity = max(1, min(999, int(item.get("quantity", 1) or 1)))
            except (TypeError, ValueError):
                quantity = 1
        parsed.append({"design_type": design_type, "quantity": quantity})
    return parsed


def service_cart_line(design_type: str, quantity: int = 1) -> dict[str, Any]:
    service = SERVICE_PRICES.get(design_type, SERVICE_PRICES["Other"])
    quantity = max(1, min(999, quantity))
    if not service.get("unit"):
        quantity = 1
    service_cents = decimal_to_cents(service.get("max", service.get("min", 0)))
    total_cents = service_cents * quantity
    label = service["label"] if quantity == 1 else f"{service['label']} x {quantity}"
    design_label = design_type if quantity == 1 else f"{design_type} x {quantity}"
    return {
        "name": label,
        "design_label": design_label,
        "price_cents": total_cents,
        "price": money_label_from_cents(total_cents),
    }


def calculate_checkout_cart(data: dict[str, str]) -> dict[str, Any]:
    cart_items = parse_design_cart_items(data.get("cart_design_items", ""))
    if not cart_items:
        design_type = data.get("design_type") or "Other"
        quantity = 1
        service = SERVICE_PRICES.get(design_type, SERVICE_PRICES["Other"])
        if service.get("unit") == "clip":
            try:
                quantity = max(1, min(999, int(data.get("clip_count", "1") or "1")))
            except ValueError:
                quantity = 1
        cart_items = [{"design_type": design_type, "quantity": quantity}]
    lines = [service_cart_line(item["design_type"], item["quantity"]) for item in cart_items]
    rush_requested = bool(data.get("rush_requested")) or data.get("rush_option") == "Rush Order +$20"
    rush_cents = decimal_to_cents(RUSH_PRICES["Rush Order +$20"]["min"]) if rush_requested else 0
    if rush_cents:
        lines.append({
            "name": RUSH_PRICES["Rush Order +$20"]["label"],
            "price_cents": rush_cents,
            "price": money_label_from_cents(rush_cents),
        })
    total_cents = sum(line["price_cents"] for line in lines)
    return {
        "lines": lines,
        "total_cents": total_cents,
        "total_display": cents_to_money(total_cents),
        "rush_cents": rush_cents,
        "rush_included": "Yes" if rush_cents else "No",
        "design_type_summary": "; ".join(line["design_label"] for line in lines if line.get("design_label")),
        "summary": "\n".join(f"- {line['name']}: {line['price']}" for line in lines),
    }


def square_api_base(config: dict[str, str]) -> str:
    return SQUARE_API_HOSTS.get(config["square_env"], SQUARE_API_HOSTS["sandbox"])


def configured_secret(value: str) -> bool:
    value = cell_text(value)
    return bool(value) and value not in ("__SET_IN_VERCEL__", "SET_IN_VERCEL", "change-this-secret")


def require_square_checkout_config(config: dict[str, str]) -> None:
    missing = [
        key for key in ("square_access_token", "square_location_id")
        if not configured_secret(config.get(key, ""))
    ]
    if missing:
        raise RuntimeError("Square checkout is not configured yet. Add SQUARE_ACCESS_TOKEN and SQUARE_LOCATION_ID in Vercel.")


def square_success_url(config: dict[str, str], request_id: str) -> str:
    return f"{config['base_url']}/simple?payment=success&request_id={urlencode({'id': request_id})[3:]}"


def create_square_payment_link(data: dict[str, str], result: dict[str, str], cart: dict[str, Any], config: dict[str, str]) -> dict[str, str]:
    require_square_checkout_config(config)
    request_id = result["request_id"]
    body = {
        "idempotency_key": str(uuid.uuid4()),
        "quick_pay": {
            "name": f"Forsaken Society Creative Request {request_id}",
            "price_money": {
                "amount": cart["total_cents"],
                "currency": "USD",
            },
            "location_id": config["square_location_id"],
        },
        "checkout_options": {
            "redirect_url": square_success_url(config, request_id),
        },
    }
    request = urllib.request.Request(
        f"{square_api_base(config)}/v2/online-checkout/payment-links",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {config['square_access_token']}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Square could not start checkout: {message}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach Square checkout: {exc.reason}") from exc
    payment_link = payload.get("payment_link") or {}
    checkout_url = cell_text(payment_link.get("url"))
    if not checkout_url:
        raise RuntimeError("Square did not return a checkout URL.")
    return {
        "checkout_url": checkout_url,
        "payment_link_id": cell_text(payment_link.get("id")),
        "order_id": cell_text(payment_link.get("order_id")),
    }


def patch_request_fields(request_id: str, updates: dict[str, str], config: dict[str, str]) -> None:
    version_payload = get_apps_script({"action": "version"}, config)
    if not version_payload.get("admin_dashboard"):
        raise RuntimeError("Redeploy the Google Apps Script web app so payment status updates can save.")
    call_apps_script(
        {
            "action": "updateRequest",
            "request_id": request_id,
            "updates": updates,
        },
        config,
    )


def find_request_by_square_order(order_id: str, config: dict[str, str]) -> dict[str, Any] | None:
    if not order_id:
        return None
    needle = f"Square order ID: {order_id}"
    for request_item in fetch_admin_requests(config):
        if needle in cell_text(request_item.get("admin_notes")):
            return request_item
    return None


def extract_note_int(notes: str, label: str) -> int:
    match = re.search(rf"^{re.escape(label)}:\s*(\d+)\s*$", notes, re.MULTILINE)
    return int(match.group(1)) if match else 0


def square_signature_is_valid(raw_body: bytes, signature: str, config: dict[str, str]) -> bool:
    key = config.get("square_webhook_signature_key")
    if not configured_secret(key) or not signature:
        return False
    notification_url = f"{config['base_url']}/api/square/webhook"
    payload = notification_url.encode("utf-8") + raw_body
    digest = hmac.new(key.encode("utf-8"), payload, hashlib.sha256).digest()
    expected = base64.b64encode(digest).decode("utf-8")
    return hmac.compare_digest(expected, signature)


def handle_square_webhook_event(event: dict[str, Any], config: dict[str, str]) -> dict[str, Any]:
    event_type = cell_text(event.get("type"))
    if event_type not in ("payment.created", "payment.updated", "order.updated"):
        return {"ok": True, "ignored": True}
    if event_type == "order.updated":
        return {"ok": True, "placeholder": True}
    payment = (((event.get("data") or {}).get("object") or {}).get("payment") or {})
    payment_status = cell_text(payment.get("status"))
    if payment_status != "COMPLETED":
        return {"ok": True, "pending": True, "payment_status": payment_status}
    order_id = cell_text(payment.get("order_id"))
    request_item = find_request_by_square_order(order_id, config)
    if not request_item:
        return {"ok": True, "pending": True, "reason": "No request matched this Square order yet."}
    if cell_text(request_item.get("request_status")) == "Paid":
        return {"ok": True, "request_id": cell_text(request_item.get("request_id")), "paid": True, "already_processed": True}
    notes = cell_text(request_item.get("admin_notes"))
    expected_total = extract_note_int(notes, "Square checkout total cents")
    paid_money = payment.get("total_money") or payment.get("amount_money") or {}
    paid_total = int(paid_money.get("amount") or 0)
    rush_requested = "Rush" in cell_text(request_item.get("rush_option"))
    rush_included = "Square rush fee included: Yes" in notes
    if expected_total and paid_total < expected_total:
        return {"ok": True, "pending": True, "reason": "Square payment amount is lower than the checkout total."}
    if rush_requested and not rush_included:
        return {"ok": True, "pending": True, "reason": "Rush fee was requested but not included in the checkout total."}
    paid_data = {
        "_payment_status": "Paid",
        "_payment_total_display": cents_to_money(paid_total or expected_total),
        "_payment_total_cents": str(paid_total or expected_total),
        "_rush_fee_included": "Yes" if rush_included else "No",
        "_cart_summary": re.search(r"Cart summary:\n([\s\S]*?)(?:\nSquare payment link ID:|\Z)", notes).group(1).strip()
        if "Cart summary:" in notes else "",
        "_square_payment_link_id": re.search(r"Square payment link ID:\s*(.+)", notes).group(1).strip()
        if "Square payment link ID:" in notes else "",
        "_square_order_id": order_id,
        "_square_checkout_url": re.search(r"Square checkout URL:\s*(.+)", notes).group(1).strip()
        if "Square checkout URL:" in notes else "",
        "_square_payment_id": cell_text(payment.get("id")),
        "_square_payment_status": payment_status,
    }
    updated_notes = combined_admin_notes({**request_item, **paid_data})
    patch_request_fields(
        cell_text(request_item.get("request_id")),
        {
            "request_status": "Paid",
            "admin_notes": updated_notes,
        },
        config,
    )
    return {"ok": True, "request_id": cell_text(request_item.get("request_id")), "paid": True}


def send_smtp_message(msg: EmailMessage, config: dict[str, str]) -> None:
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

    send_smtp_message(msg, config)
    return True


def send_client_confirmation_email(data: dict[str, str], result: dict[str, str]) -> bool:
    config = get_config()
    required = ["smtp_host", "smtp_port", "smtp_from"]
    if any(not config[key] for key in required) or not data.get("email"):
        return False

    msg = EmailMessage()
    msg["Subject"] = f"We got your request: {result['request_id']}"
    msg["From"] = config["smtp_from"]
    msg["To"] = data["email"]
    if config["admin_email"]:
        msg["Reply-To"] = config["admin_email"]

    plain = [
        "I got your request.",
        "",
        f"Request ID: {result['request_id']}",
        f"Project: {data.get('project_name', '')}",
        f"Design Type: {data.get('design_type', '')}",
        f"Requested Deadline: {data.get('requested_deadline', '')}",
        "",
        "NAXYSTUDIOS LLC will review it and follow up soon.",
    ]
    msg.set_content("\n".join(plain))
    msg.add_alternative(
        f"""
        <html>
          <body>
            <h2>I got your request.</h2>
            <p>NAXYSTUDIOS LLC will review it and follow up soon.</p>
            <table cellpadding="8" cellspacing="0" border="1">
              <tr><th>Request ID</th><td>{html.escape(result['request_id'])}</td></tr>
              <tr><th>Project</th><td>{html.escape(data.get('project_name', ''))}</td></tr>
              <tr><th>Design Type</th><td>{html.escape(data.get('design_type', ''))}</td></tr>
              <tr><th>Requested Deadline</th><td>{html.escape(data.get('requested_deadline', ''))}</td></tr>
            </table>
          </body>
        </html>
        """,
        subtype="html",
    )

    send_smtp_message(msg, config)
    return True


def send_designer_assignment_email(designer: dict[str, str], request_data: dict[str, str]) -> bool:
    config = get_config()
    required = ["smtp_host", "smtp_port", "smtp_from"]
    if any(not config[key] for key in required) or not designer.get("email"):
        return False

    msg = EmailMessage()
    msg["Subject"] = f"Project assigned: {request_data.get('request_id', '')}"
    msg["From"] = config["smtp_from"]
    msg["To"] = designer["email"]
    if config["admin_email"]:
        msg["Reply-To"] = config["admin_email"]

    lines = [
        "A design request has been assigned to you.",
        "",
        f"Request ID: {request_data.get('request_id', '')}",
        f"Project: {request_data.get('project_name', '')}",
        f"Design Type: {request_data.get('design_type', '')}",
        f"Requested Deadline: {request_data.get('requested_deadline', '')}",
    ]
    msg.set_content("\n".join(lines))
    send_smtp_message(msg, config)
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
    .background-video {{
      position: fixed;
      inset: 0;
      z-index: 0;
      width: 100%;
      height: 100%;
      object-fit: cover;
      pointer-events: none;
    }}
    body::before {{
      content: "";
      position: fixed;
      inset: -7%;
      z-index: 1;
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
    .client-progress-link {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 36px;
      border: 1px solid rgba(247, 183, 51, .78);
      border-radius: 6px;
      padding: 0 12px;
      color: var(--gold);
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
    .rush-toggle span {{
      display: flex;
      align-items: center;
      gap: 10px;
      min-height: 44px;
      border: 1px solid #d2ccc2;
      border-radius: 6px;
      background: var(--field);
      padding: 10px 12px;
      font-weight: 700;
    }}
    .rush-toggle input[type="checkbox"] {{
      width: 18px;
      min-height: 18px;
      accent-color: var(--accent);
    }}
    .order-cart {{
      grid-column: 1 / -1;
      border: 1px solid #d8d1c8;
      border-radius: 8px;
      background: #fffdf8;
      padding: 16px;
    }}
    .order-cart h3 {{
      margin: 0 0 12px;
      color: #161616;
      font-size: 1rem;
      text-transform: uppercase;
    }}
    .cart-list {{
      display: grid;
      gap: 8px;
      margin-bottom: 12px;
    }}
    .cart-line, .cart-total {{
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 14px;
    }}
    .cart-line > span:first-child {{
      min-width: 0;
      overflow-wrap: anywhere;
    }}
    .cart-line span:first-child {{ color: #3f3f3f; font-weight: 800; }}
    .cart-line span:last-child {{ color: var(--accent-dark); font-weight: 900; white-space: nowrap; }}
    .cart-line button {{
      min-height: 30px;
      padding: 5px 8px;
      border: 1px solid #991b1b;
      border-radius: 6px;
      background: #fff1f2;
      color: #991b1b;
      font-size: .72rem;
    }}
    .cart-price {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      white-space: nowrap;
    }}
    .clip-count {{
      grid-column: 1 / -1;
    }}
    .cart-builder {{
      grid-column: 1 / -1;
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }}
    .cart-builder button {{
      min-height: 40px;
      padding: 9px 14px;
    }}
    .cart-total {{
      border-top: 1px solid #e4e0da;
      padding-top: 12px;
      color: #161616;
      font-size: 1.1rem;
      font-weight: 900;
    }}
    .cart-note {{
      display: block;
      margin-top: 8px;
      color: var(--muted);
      font-size: .82rem;
      font-weight: 400;
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
    @media (prefers-reduced-motion: reduce) {{
      .background-video {{ display: none; }}
      body::before {{ animation: none; }}
    }}
  </style>
</head>
<body>
  <video class="background-video" autoplay muted loop playsinline poster="/static/forsaken-background.png" aria-hidden="true">
    <source src="/static/forsaken-background.mp4" type="video/mp4">
  </video>
  <script>
    (() => {{
      const video = document.querySelector(".background-video");
      if (!video) return;
      const setSlowPlayback = () => {{ video.playbackRate = 0.85; }};
      video.addEventListener("loadedmetadata", setSlowPlayback);
      video.addEventListener("play", setSlowPlayback);
      setSlowPlayback();
    }})();
  </script>
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
          <a class="client-progress-link" href="/work-in-process">Work</a>
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
      if (!select || select.tagName !== "SELECT") return;
      select.innerHTML = "";
      values.forEach((value) => {{
        const option = document.createElement("option");
        option.value = value;
        const price = id === "design_type" ? choices.servicePrices[value] : null;
        option.textContent = price && value !== "AVIs" ? `${{value}} - ${{price.display}}` : value;
        if (value === selected) option.selected = true;
        select.appendChild(option);
      }});
    }}
    fillSelect("design_type", choices.designTypes, "AVIs");
    fillSelect("priority", choices.priorities, "Standard");

    const form = document.querySelector("form");
    const usesSquareCheckout = form && form.getAttribute("action") === "/simple-submit";
    const designTypeSelect = document.getElementById("design_type");
    const otherDesignWrap = document.getElementById("other_design_wrap");
    const otherDesignInput = document.getElementById("other_design_type");
    const clipCountWrap = document.getElementById("clip_count_wrap");
    const clipCountInput = document.getElementById("clip_count");
    const addDesignToCartButton = document.getElementById("add_design_to_cart");
    const rushInput = document.getElementById("rush_option");
    const rushCheckbox = document.getElementById("rush_requested");
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

    function syncClipCount() {{
      if (!clipCountWrap || !clipCountInput) return;
      const service = choices.servicePrices[designTypeSelect.value] || {{}};
      const needsClipCount = service.unit === "clip";
      clipCountWrap.classList.toggle("hidden", !needsClipCount);
      clipCountInput.required = needsClipCount;
      if (!needsClipCount) clipCountInput.value = "1";
    }}
    designTypeSelect.addEventListener("change", syncClipCount);
    syncClipCount();

    function syncRushPayment() {{
      const needsPayment = Boolean(rushCheckbox && rushCheckbox.checked);
      const value = needsPayment ? "Rush Order +$20" : "No Rush";
      if (rushInput) rushInput.value = value;
      rushPayment.classList.add("hidden");
      rushPayButton.href = "#";
      rushPaymentText.textContent = "";
      rushPaymentLink.value = "";
      rushPaymentConfirmed.required = false;
      rushPaymentConfirmed.checked = false;
    }}
    if (rushCheckbox) rushCheckbox.addEventListener("change", syncRushPayment);
    syncRushPayment();

    const cartItems = document.getElementById("cart_items");
    const cartTotal = document.getElementById("cart_total");
    const cartNote = document.getElementById("cart_note");
    const cartDesignItemsInput = document.getElementById("cart_design_items");
    const cartTotalCentsInput = document.getElementById("cart_total_cents");
    const rushFeeCentsInput = document.getElementById("rush_fee_cents");
    const cartLinesInput = document.getElementById("cart_lines_json");
    let designCartLines = [];
    function money(value) {{
      const amount = Number(value || 0);
      return amount % 1 === 0
        ? `$${{amount.toLocaleString()}}`
        : `$${{amount.toLocaleString(undefined, {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }})}}`;
    }}
    function centsToMoney(cents) {{
      return money(Number(cents || 0) / 100);
    }}
    function dollarsToCents(value) {{
      return Math.max(0, Math.round(Number(value || 0) * 100));
    }}
    function currentClipCount() {{
      if (!clipCountInput) return 1;
      const value = Number.parseInt(clipCountInput.value || "1", 10);
      return Number.isFinite(value) && value > 0 ? value : 1;
    }}
    function escapeText(value) {{
      return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }}
    function selectedDesignCartItem() {{
      const designType = designTypeSelect.value || "Other";
      const service = choices.servicePrices[designType] || choices.servicePrices.Other;
      return {{
        design_type: designType,
        quantity: service.unit === "clip" ? currentClipCount() : 1,
      }};
    }}
    function syncDesignCartInput() {{
      if (!cartDesignItemsInput) return;
      cartDesignItemsInput.value = JSON.stringify(designCartLines.map((line) => ({{
        design_type: line.design_type,
        quantity: line.quantity,
      }})));
    }}
    function addSelectedDesignToCart() {{
      designCartLines.push(selectedDesignCartItem());
      syncDesignCartInput();
      syncCart();
    }}
    function syncCart() {{
      if (!cartItems || !cartTotal) return;
      const lines = designCartLines.map((line, index) => {{
        const service = choices.servicePrices[line.design_type] || choices.servicePrices.Other;
        const quantity = service.unit === "clip" ? Math.max(1, Number(line.quantity || 1)) : 1;
        const cents = dollarsToCents(service.max || service.min || 0) * quantity;
        return {{
          name: quantity > 1 ? `${{service.label}} x ${{quantity}}` : service.label,
          price: centsToMoney(cents),
          cents,
          designIndex: index,
        }};
      }});
      const hasDesignItems = lines.length > 0;
      if (hasDesignItems) {{
        const rushValue = rushCheckbox && rushCheckbox.checked ? "Rush Order +$20" : "No Rush";
        const rush = choices.rushPrices[rushValue] || choices.rushPrices["No Rush"];
        if ((rush.min || 0) > 0) {{
          lines.push({{ name: rush.label, price: rush.display, cents: dollarsToCents(rush.min || 0), rush: true }});
        }}
      }}
      cartItems.innerHTML = lines.map((line) => `
        <div class="cart-line">
          <span>${{escapeText(line.name)}}</span>
          <span class="cart-price">${{line.price}}${{line.designIndex !== undefined ? `<button type="button" data-remove-design="${{line.designIndex}}">Remove</button>` : ""}}</span>
        </div>
      `).join("") || `
        <div class="cart-line">
          <span>Pick a design type and add it to the cart.</span>
          <span class="cart-price">$0</span>
        </div>
      `;
      cartItems.querySelectorAll("[data-remove-design]").forEach((button) => {{
        button.addEventListener("click", () => {{
          designCartLines.splice(Number(button.dataset.removeDesign), 1);
          syncDesignCartInput();
          syncCart();
        }});
      }});
      const totalCents = lines.reduce((sum, line) => sum + (line.cents || 0), 0);
      const rushCents = lines.reduce((sum, line) => sum + (line.rush ? line.cents || 0 : 0), 0);
      cartTotal.textContent = centsToMoney(totalCents);
      if (cartTotalCentsInput) cartTotalCentsInput.value = String(totalCents);
      if (rushFeeCentsInput) rushFeeCentsInput.value = String(rushCents);
      if (cartLinesInput) {{
        cartLinesInput.value = JSON.stringify(lines.map((line) => ({{
          name: line.name,
          price: line.price,
          cents: line.cents || 0,
          rush: Boolean(line.rush),
        }})));
      }}
      if (cartNote) {{
        cartNote.textContent = hasDesignItems
          ? "Total due based on selected cart items."
          : "Add at least one design type to the cart summary.";
      }}
      syncDesignCartInput();
    }}
    if (addDesignToCartButton) addDesignToCartButton.addEventListener("click", addSelectedDesignToCart);
    designTypeSelect.addEventListener("change", syncCart);
    if (clipCountInput) clipCountInput.addEventListener("input", syncCart);
    if (rushCheckbox) rushCheckbox.addEventListener("change", syncCart);
    syncCart();

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

    function setFormStatus(message, isError = false) {{
      let box = form.querySelector(".status");
      if (!box) {{
        box = document.createElement("div");
        box.className = "status";
        form.prepend(box);
      }}
      box.textContent = message;
      box.classList.toggle("error", Boolean(isError));
    }}

    async function prepareUploads() {{
      const files = Array.from((fileInput && fileInput.files) || []);
      const totalBytes = files.reduce((sum, file) => sum + file.size, 0);
      if (totalBytes > maxUploadBytes) {{
        throw new Error("Uploads must be 4 MB total or less. For larger videos, paste a share link instead.");
      }}
      if (files.length && filePayload) {{
        filePayload.value = JSON.stringify(await Promise.all(files.map(readFileAsPayload)));
      }}
    }}

    function formValues() {{
      if (!designCartLines.length) addSelectedDesignToCart();
      syncCart();
      const values = Object.fromEntries(new FormData(form).entries());
      values.custom_cart_items = "[]";
      values.cart_design_items = cartDesignItemsInput ? cartDesignItemsInput.value : "[]";
      values.cart_total_cents = cartTotalCentsInput ? cartTotalCentsInput.value : "0";
      values.rush_fee_cents = rushFeeCentsInput ? rushFeeCentsInput.value : "0";
      values.cart_lines_json = cartLinesInput ? cartLinesInput.value : "[]";
      return values;
    }}

    form.addEventListener("submit", async (event) => {{
      if (!usesSquareCheckout) {{
        try {{
          await prepareUploads();
          if (filePayload && filePayload.value) {{
            event.preventDefault();
            submitButton.disabled = true;
            submitButton.textContent = "Uploading...";
            form.submit();
          }}
        }} catch (error) {{
          event.preventDefault();
          alert(error.message || "Could not prepare the upload. Please try again or paste a share link.");
          submitButton.disabled = false;
          submitButton.textContent = "Submit Request";
        }}
        return;
      }}
      event.preventDefault();
      submitButton.disabled = true;
      submitButton.textContent = "Creating checkout...";
      setFormStatus("Creating checkout...");
      try {{
        await prepareUploads();
        const response = await fetch("/api/square/create-checkout", {{
          method: "POST",
          headers: {{ "Content-Type": "application/json" }},
          body: JSON.stringify({{ request: formValues() }}),
        }});
        const payload = await response.json().catch(() => ({{}}));
        if (!response.ok || !payload.checkoutUrl) {{
          throw new Error(payload.error || "Payment failed to start");
        }}
        if (payload.noPaymentRequired) {{
          submitButton.textContent = "Payment received / request submitted";
          setFormStatus("Payment received / request submitted");
          window.location.assign(payload.checkoutUrl);
          return;
        }}
        submitButton.textContent = "Redirecting to Square...";
        setFormStatus("Redirecting to Square...");
        window.location.assign(payload.checkoutUrl);
      }} catch (error) {{
        setFormStatus("Payment failed to start", true);
        if (error && error.message && error.message !== "Payment failed to start") {{
          console.error(error);
        }}
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
              <li>Rush Orders add a flat $20 fee.</li>
              <li>Pricing may vary depending on complexity, revisions, licensing, and turnaround time.</li>
            </ul>
          </section>
          <section class="service-card">
            <h3>Social Media Revamp</h3>
            <ul>
              <li>Social Media Packages - $100</li>
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
              <li>Custom 3D Intro/Outro - $150</li>
              <li>Custom Merch Designs - $55-$230 (Non-NSFW designs)</li>
              <li>Advertisements - $40-$70</li>
              <li>Poster Advertisements - $60-$300</li>
              <li>Custom Designs (NFTs, Album Covers) - $40-$220</li>
              <li>Logo Designs - $75-$350</li>
              <li>Flyer Design - $40-$100</li>
              <li>Yard Signs Designs - $40-$150</li>
              <li>UI Website Design - $180 (Square Up or Wix)</li>
              <li>Thumbnail Designs - $15-$40</li>
              <li>Profile Avatars/Profile Photos - Free for AVIs</li>
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
        <label>Design Type
          <select id="design_type" name="design_type" required></select>
        </label>
        <label id="other_design_wrap" class="hidden">Tell Us What You Need
          <input id="other_design_type" name="other_design_type" placeholder="Describe the custom service request">
        </label>
        <label id="clip_count_wrap" class="clip-count hidden">Clip Count
          <input id="clip_count" name="clip_count" type="number" min="1" step="1" value="1">
        </label>
        <label>Priority Level
          <select id="priority" name="priority" required></select>
        </label>
        <label class="rush-toggle">Rush Order
          <span>
            <input id="rush_requested" name="rush_requested" type="checkbox" value="Rush Order +$20">
            Add rush fee (+$20)
          </span>
          <input id="rush_option" name="rush_option" type="hidden" value="No Rush">
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
        <section class="order-cart" aria-live="polite">
          <h3>Cart Summary</h3>
          <div id="cart_items" class="cart-list"></div>
          <div class="cart-total">
            <span>Total Due</span>
            <span id="cart_total">$0</span>
          </div>
          <span id="cart_note" class="cart-note"></span>
          <input id="cart_total_cents" name="cart_total_cents" type="hidden">
          <input id="rush_fee_cents" name="rush_fee_cents" type="hidden">
          <input id="cart_lines_json" name="cart_lines_json" type="hidden">
        </section>
      </div>
      <div class="actions">
        <button type="submit">Submit Request</button>
      </div>
    </form>
    {pricing_guide_html()}
    """


def simple_form_html(status: str = "", error: bool = False) -> str:
    status_html = f'<div class="status{" error" if error else ""}">{html.escape(status)}</div>' if status else ""
    return f"""
    <form method="post" action="/simple-submit">
      {status_html}
      <div class="grid">
        <label>Member Name
          <input name="member_name" autocomplete="name" required>
        </label>
        <label>Email
          <input name="email" type="email" autocomplete="email" required>
        </label>
        <label class="full">Design Type
          <select id="design_type" name="design_type" required></select>
        </label>
        <label id="other_design_wrap" class="hidden">Tell Us What You Need
          <input id="other_design_type" name="other_design_type" placeholder="Describe the custom service request">
        </label>
        <label id="clip_count_wrap" class="clip-count hidden">Clip Count
          <input id="clip_count" name="clip_count" type="number" min="1" step="1" value="1">
        </label>
        <div class="cart-builder">
          <button id="add_design_to_cart" type="button">Add Selected Design to Cart</button>
          <span class="hint">Pick a design type, then add it to the cart summary.</span>
        </div>
        <label class="full">Describe What You Need
          <textarea name="description" required placeholder="Include style, colors, text, platform, and any reference links."></textarea>
        </label>
        <label class="full">File or URL Link
          <input name="uploaded_files_link" type="url" placeholder="Optional Google Drive, Canva, Dropbox, or reference link">
        </label>
        <label class="full">Notes for Designer
          <textarea name="notes" placeholder="Optional"></textarea>
        </label>
        <label class="hidden">Priority Level
          <select id="priority" name="priority"></select>
        </label>
        <label class="rush-toggle full">Rush Order
          <span>
            <input id="rush_requested" name="rush_requested" type="checkbox" value="Rush Order +$20">
            Add rush fee (+$20)
          </span>
          <input id="rush_option" name="rush_option" type="hidden" value="No Rush">
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
        <label class="hidden">Upload Photos or Videos
          <input id="upload_files" name="upload_files" type="file" accept="image/*,video/*" multiple>
        </label>
        <input id="uploaded_files_json" name="uploaded_files_json" type="hidden">
        <section class="order-cart" aria-live="polite">
          <h3>Cart Summary</h3>
          <div id="cart_items" class="cart-list"></div>
          <div class="cart-total">
            <span>Total Due</span>
            <span id="cart_total">$0</span>
          </div>
          <span id="cart_note" class="cart-note"></span>
          <input id="cart_design_items" name="cart_design_items" type="hidden">
          <input id="cart_total_cents" name="cart_total_cents" type="hidden">
          <input id="rush_fee_cents" name="rush_fee_cents" type="hidden">
          <input id="cart_lines_json" name="cart_lines_json" type="hidden">
        </section>
      </div>
      <div class="actions">
        <button type="submit">Submit Request</button>
      </div>
    </form>
    """


def process_page_html(request_item: dict[str, Any] | None = None, status: str = "", error: bool = False) -> bytes:
    status_html = f'<div class="notice{" error" if error else ""}">{html.escape(status)}</div>' if status else ""
    result_html = ""
    if request_item:
        label, color = request_progress_state(request_item)
        result_html = f"""
        <section class="result">
          <div class="progress-line">
            <span class="dot {html.escape(color)}"></span>
            <strong>{html.escape(label)}</strong>
          </div>
          <dl>
            <div><dt>Request ID</dt><dd>{html.escape(cell_text(request_item.get("request_id")))}</dd></div>
            <div><dt>Project</dt><dd>{html.escape(cell_text(request_item.get("project_name")))}</dd></div>
            <div><dt>Design Type</dt><dd>{html.escape(cell_text(request_item.get("design_type")))}</dd></div>
            <div><dt>Approval</dt><dd>{html.escape(cell_text(request_item.get("approval_status")) or "Pending Review")}</dd></div>
            <div><dt>Status</dt><dd>{html.escape(cell_text(request_item.get("request_status")) or "Submitted")}</dd></div>
            <div><dt>Assigned Designer</dt><dd>{html.escape(cell_text(request_item.get("assigned_designer")) or "Pending")}</dd></div>
          </dl>
        </section>
        """
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Request Process</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      font-family: Arial, Helvetica, sans-serif;
      color: #171717;
      background:
        linear-gradient(90deg, rgba(10,10,10,.92), rgba(10,10,10,.62)),
        url("/static/forsaken-background.png") center/cover fixed,
        #111;
    }}
    main {{
      width: min(760px, calc(100% - 28px));
      background: rgba(255,255,255,.97);
      border-radius: 8px;
      box-shadow: 0 22px 70px rgba(0,0,0,.36);
      overflow: hidden;
    }}
    header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      background: #161616;
      color: #fff;
      border-bottom: 4px solid #e74719;
      padding: 20px 24px;
    }}
    h1 {{ margin: 0; text-transform: uppercase; font-size: clamp(1.4rem, 4vw, 2.2rem); }}
    header a {{ color: #f7b733; font-weight: 900; text-transform: uppercase; text-decoration: none; }}
    form, .result {{ padding: 24px; }}
    .grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }}
    label {{ display: grid; gap: 7px; font-weight: 900; text-transform: uppercase; font-size: .82rem; }}
    input {{
      width: 100%;
      min-height: 44px;
      border: 1px solid #d4cec4;
      border-radius: 6px;
      background: #f8f5ef;
      padding: 10px 12px;
      font: inherit;
      text-transform: none;
    }}
    button {{
      min-height: 44px;
      border: 0;
      border-radius: 6px;
      background: #e74719;
      color: #fff;
      padding: 10px 16px;
      font-weight: 900;
      text-transform: uppercase;
      cursor: pointer;
    }}
    .actions {{ margin-top: 16px; }}
    .notice {{
      margin: 18px 24px 0;
      padding: 12px 14px;
      border-radius: 6px;
      background: #f0fdf4;
      color: #166534;
      font-weight: 800;
    }}
    .notice.error {{ background: #fff7ed; color: #a24112; }}
    .result {{ border-top: 1px solid #e4e0da; }}
    .progress-line {{ display: flex; align-items: center; gap: 10px; font-size: 1.25rem; text-transform: uppercase; }}
    .dot {{ width: 14px; height: 14px; border-radius: 999px; display: inline-block; box-shadow: 0 0 0 4px rgba(0,0,0,.08); }}
    .dot.red {{ background: #dc2626; }}
    .dot.yellow {{ background: #facc15; }}
    .dot.green {{ background: #16a34a; }}
    dl {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; margin: 18px 0 0; }}
    dt {{ color: #6d6d6d; font-size: .78rem; font-weight: 900; text-transform: uppercase; }}
    dd {{ margin: 4px 0 0; font-weight: 800; }}
    @media (max-width: 640px) {{
      header {{ display: block; }}
      header a {{ display: inline-block; margin-top: 8px; }}
      .grid, dl {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <h1>Request Process</h1>
      <a href="/">Back to form</a>
    </header>
    {status_html}
    <form method="post" action="/process">
      <div class="grid">
        <label>Request ID
          <input name="request_id" placeholder="CRP-0011" required>
        </label>
        <label>Email
          <input name="email" type="email" placeholder="you@example.com" required>
        </label>
      </div>
      <div class="actions">
        <button type="submit">Check Process</button>
      </div>
    </form>
    {result_html}
  </main>
</body>
</html>""".encode("utf-8")


def public_work_page_html(requests: list[dict[str, Any]], status: str = "", error: bool = False) -> bytes:
    visible = [
        item for item in requests
        if cell_text(item.get("request_status")) != "Deleted"
        and request_progress_state(item)[0] in ("Not Started", "In Process", "Done")
    ]
    rows = []
    for item in visible:
        label, color = request_progress_state(item)
        rows.append(f"""
        <tr>
          <td>{html.escape(cell_text(item.get("member_name")) or "Client")}</td>
          <td>{html.escape(cell_text(item.get("design_type")) or "Design")}</td>
          <td><span class="stage"><span class="dot {html.escape(color)}"></span>{html.escape(label)}</span></td>
        </tr>
        """)
    table_body = "\n".join(rows) if rows else '<tr><td colspan="3" class="empty">No public work items yet.</td></tr>'
    status_html = f'<div class="notice{" error" if error else ""}">{html.escape(status)}</div>' if status else ""
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Work In Process</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      font-family: Arial, Helvetica, sans-serif;
      color: #171717;
      background:
        linear-gradient(90deg, rgba(10,10,10,.92), rgba(10,10,10,.62)),
        url("/static/forsaken-background.png") center/cover fixed,
        #111;
    }}
    main {{ width: min(980px, calc(100% - 28px)); margin: 0 auto; padding: 32px 0; }}
    header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      color: #fff;
      margin-bottom: 18px;
    }}
    h1 {{ margin: 0; text-transform: uppercase; font-size: clamp(2rem, 6vw, 4rem); line-height: .92; }}
    header a {{ color: #f7b733; font-weight: 900; text-transform: uppercase; text-decoration: none; }}
    .legend {{
      display: flex;
      gap: 14px;
      flex-wrap: wrap;
      background: rgba(255,255,255,.96);
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 14px;
      font-weight: 900;
      text-transform: uppercase;
    }}
    .legend span, .stage {{ display: inline-flex; align-items: center; gap: 8px; }}
    .dot {{
      width: 12px;
      height: 12px;
      border-radius: 999px;
      display: inline-block;
      box-shadow: 0 0 0 3px rgba(0,0,0,.08);
    }}
    .dot.red {{ background: #dc2626; }}
    .dot.yellow {{ background: #facc15; }}
    .dot.green {{ background: #16a34a; }}
    .notice {{
      margin-bottom: 14px;
      padding: 12px 14px;
      border-radius: 6px;
      background: #f0fdf4;
      color: #166534;
      font-weight: 800;
    }}
    .notice.error {{ background: #fff7ed; color: #a24112; }}
    .table-wrap {{
      background: rgba(255,255,255,.97);
      border-radius: 8px;
      overflow: auto;
      box-shadow: 0 20px 70px rgba(0,0,0,.36);
    }}
    table {{ width: 100%; border-collapse: collapse; min-width: 620px; }}
    th {{
      background: #161616;
      color: #fff;
      text-align: left;
      padding: 14px;
      text-transform: uppercase;
      font-size: .82rem;
    }}
    td {{ border-top: 1px solid #e4e0da; padding: 14px; font-weight: 800; }}
    .empty {{ text-align: center; color: #6d6d6d; padding: 34px; }}
    @media (max-width: 700px) {{
      header {{ display: block; }}
      header a {{ display: inline-block; margin-top: 10px; }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <h1>Work In Process</h1>
      <a href="/">Back to form</a>
    </header>
    {status_html}
    <section class="legend">
      <span><span class="dot red"></span>Not Started</span>
      <span><span class="dot yellow"></span>In Process</span>
      <span><span class="dot green"></span>Done</span>
    </section>
    <section class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Member Name</th>
            <th>Design Type</th>
            <th>Stage</th>
          </tr>
        </thead>
        <tbody>{table_body}</tbody>
      </table>
    </section>
  </main>
</body>
</html>""".encode("utf-8")


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


def admin_dashboard_html(
    requests: list[dict[str, Any]],
    status: str = "",
    error: bool = False,
    designers: list[dict[str, str]] | None = None,
) -> bytes:
    designers = designers or []
    total = len(requests)
    unseen = sum(1 for item in requests if item.get("seen_status") != "Seen")
    pending = sum(1 for item in requests if item.get("approval_status") in ("", "Pending Review"))
    approved = sum(1 for item in requests if item.get("approval_status") == "Approved")
    disapproved = sum(1 for item in requests if item.get("approval_status") in ("Not Approved", "Needs Info"))
    active = sum(1 for item in requests if item.get("request_status") in ("Submitted", "In Progress", "Revision"))
    status_class = " admin-alert-error" if error else ""
    status_html = f'<div class="admin-alert{status_class}">{html.escape(status)}</div>' if status else ""

    def text(item: dict[str, Any], key: str) -> str:
        return html.escape(cell_text(item.get(key)))

    def designer_options(selected: str) -> str:
        names = [designer["name"] for designer in designers if designer.get("name")]
        if selected and selected not in names:
            names.insert(0, selected)
        options = ['<option value="">Unassigned</option>']
        options.extend(
            f'<option value="{html.escape(name, quote=True)}"{" selected" if name == selected else ""}>{html.escape(name)}</option>'
            for name in names
        )
        return "".join(options)

    designer_rows = "".join(
        f"""
        <tr>
          <td>{html.escape(designer.get("name", ""))}</td>
          <td>{html.escape(designer.get("email", ""))}</td>
          <td>
            <form method="post" action="/admin/designers/delete" onsubmit="return confirm('Remove {html.escape(designer.get("name", ""), quote=True)} from designer options?');">
              <input type="hidden" name="designer_email" value="{html.escape(designer.get("email", ""), quote=True)}">
              <button type="submit">Remove</button>
            </form>
          </td>
        </tr>
        """
        for designer in designers
    ) or '<tr><td colspan="3" class="empty">No designers added yet.</td></tr>'

    rows = []
    for item in requests:
        request_id = cell_text(item.get("request_id"))
        source_tab = cell_text(item.get("source_tab")) or "Request Tracker"
        approval_status = cell_text(item.get("approval_status")) or "Pending Review"
        seen_status = cell_text(item.get("seen_status")) or "Not Seen"
        progress_label, progress_color = request_progress_state(item)
        asset_link = cell_text(item.get("uploaded_files_link"))
        final_link = cell_text(item.get("final_deliverables_link"))
        asset_html = f'<a href="{html.escape(asset_link, quote=True)}" target="_blank" rel="noopener">Open assets</a>' if asset_link else ""
        final_html = f'<a href="{html.escape(final_link, quote=True)}" target="_blank" rel="noopener">Open final</a>' if final_link else ""
        rows.append(f"""
        <tr data-source-tab="{html.escape(source_tab, quote=True)}" data-approval-status="{html.escape(approval_status, quote=True)}" data-seen-status="{html.escape(seen_status, quote=True)}" data-progress-state="{html.escape(progress_label, quote=True)}">
          <td class="select-cell">
            <input class="request-check" form="bulk-delete-form" type="checkbox" name="request_id" value="{html.escape(request_id, quote=True)}" aria-label="Select {html.escape(request_id, quote=True)}">
          </td>
          <td>
            <strong>{html.escape(request_id)}</strong>
            <span>{text(item, "submitted_at")}</span>
            <span>{html.escape(source_tab)}</span>
            <span class="progress-pill"><span class="status-dot {html.escape(progress_color)}"></span>{html.escape(progress_label)}</span>
          </td>
          <td class="client-cell">
            <strong>{text(item, "member_name") or "Unnamed Client"}</strong>
            <span>{text(item, "email")}</span>
          </td>
          <td class="project-cell">
            <strong>{text(item, "design_type")}</strong>
            <span>{text(item, "rush_option")} · {text(item, "priority")}</span>
          </td>
          <td>
            <form class="request-form" method="post" action="/admin/update">
              <input type="hidden" name="request_id" value="{html.escape(request_id, quote=True)}">
              <div class="quick-controls">
                <label>Seen
                  <select name="seen_status">{option_tags(SEEN_STATUS_OPTIONS, cell_text(item.get("seen_status")) or "Not Seen")}</select>
                </label>
                <label>Status
                  <select name="request_status">{option_tags(ADMIN_STATUS_OPTIONS, cell_text(item.get("request_status")) or "Submitted")}</select>
                </label>
                <label>Approval
                  <select name="approval_status">{option_tags(APPROVAL_STATUS_OPTIONS, cell_text(item.get("approval_status")) or "Pending Review")}</select>
                </label>
                <label>Designer
                  <select name="assigned_designer">{designer_options(cell_text(item.get("assigned_designer")))}</select>
                </label>
              </div>
              <details class="admin-more">
                <summary>Details and notes</summary>
              <div class="control-grid">
                <label>Seen By
                  <input name="seen_by" value="{text(item, "seen_by")}" placeholder="Admin name">
                </label>
                <label>Final Link
                  <input name="final_deliverables_link" type="url" value="{text(item, "final_deliverables_link")}" placeholder="https://...">
                </label>
                <label class="wide">Admin Notes
                  <textarea name="admin_notes">{text(item, "admin_notes")}</textarea>
                </label>
              </div>
              </details>
              <details class="request-more">
                <summary>Client request</summary>
                <p>{text(item, "description")}</p>
                <p>{text(item, "admin_notes")}</p>
              </details>
              <div class="request-links">{asset_html}{final_html}</div>
              <div class="row-actions">
                <button type="submit">Save</button>
              </div>
            </form>
            <form class="delete-form" method="post" action="/admin/delete" onsubmit="return confirm('Delete {html.escape(request_id, quote=True)} from the admin dashboard? This cannot be undone from the site.');">
              <input type="hidden" name="request_id" value="{html.escape(request_id, quote=True)}">
              <button type="submit">Delete</button>
            </form>
          </td>
        </tr>
        """)

    table_body = "\n".join(rows) if rows else '<tr><td colspan="5" class="empty">No requests found yet.</td></tr>'
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
    .bulk-bar {{
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
      background: rgba(255,255,255,.97);
      border-radius: 8px;
      margin-bottom: 12px;
      padding: 12px;
    }}
    .bulk-select {{
      display: flex;
      align-items: center;
      gap: 8px;
      color: #171717;
      font-size: .82rem;
      font-weight: 900;
      text-transform: uppercase;
    }}
    .bulk-select input, .select-cell input {{
      width: 18px;
      min-height: 18px;
      accent-color: #e74719;
    }}
    .bulk-bar button {{
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
    .bulk-count {{ color: #6d6d6d; font-weight: 800; }}
    .designer-panel {{
      display: none;
      background: rgba(255,255,255,.97);
      border-radius: 8px;
      padding: 16px;
      box-shadow: 0 20px 70px rgba(0,0,0,.36);
    }}
    .designer-panel.active {{ display: block; }}
    .designer-form {{
      display: grid;
      grid-template-columns: minmax(180px, 1fr) minmax(220px, 1fr) auto;
      gap: 10px;
      align-items: end;
      margin-bottom: 16px;
    }}
    .designer-form button, .designer-panel table button {{
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
    .progress-pill {{
      display: inline-flex;
      align-items: center;
      gap: 7px;
      color: #171717;
      font-weight: 900;
    }}
    .status-dot {{
      width: 10px;
      height: 10px;
      border-radius: 999px;
      display: inline-block;
      box-shadow: 0 0 0 3px rgba(0,0,0,.08);
    }}
    .status-dot.red {{ background: #dc2626; }}
    .status-dot.yellow {{ background: #facc15; }}
    .status-dot.green {{ background: #16a34a; }}
    .admin-tabs {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 12px;
    }}
    .admin-tab {{
      min-height: 38px;
      border: 1px solid rgba(255,255,255,.28);
      border-radius: 6px;
      background: rgba(255,255,255,.14);
      color: #fff;
      padding: 8px 12px;
      font-weight: 900;
      text-transform: uppercase;
      cursor: pointer;
    }}
    .admin-tab.active {{
      background: #e74719;
      border-color: #e74719;
    }}
    table {{ width: 100%; border-collapse: collapse; min-width: 1040px; }}
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
    tbody tr {{ transition: background .18s ease; }}
    tbody tr:hover {{ background: #fffaf4; }}
    .select-cell {{ width: 44px; text-align: center; }}
    td > span, td strong + span {{ display: block; color: #6d6d6d; margin-top: 4px; font-size: .86rem; }}
    .client-cell {{ min-width: 180px; }}
    .project-cell {{ min-width: 180px; }}
    .request-form {{ min-width: 520px; }}
    .quick-controls {{ display: grid; grid-template-columns: repeat(4, minmax(118px, 1fr)); gap: 8px; align-items: end; }}
    .control-grid {{ display: grid; grid-template-columns: repeat(2, minmax(170px, 1fr)); gap: 10px; margin-top: 10px; }}
    label {{ display: grid; gap: 5px; color: #333; font-size: .78rem; font-weight: 900; text-transform: uppercase; }}
    input, select, textarea {{
      width: 100%;
      min-height: 36px;
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
    details {{ margin-top: 9px; color: #4b4b4b; }}
    summary {{
      width: max-content;
      max-width: 100%;
      color: #333;
      cursor: pointer;
      font-size: .78rem;
      font-weight: 900;
      text-transform: uppercase;
    }}
    .admin-more {{
      border: 1px solid #eadfd4;
      border-radius: 8px;
      padding: 9px;
      background: #fffdf9;
    }}
    .request-more p {{ white-space: pre-wrap; line-height: 1.45; }}
    .row-actions {{ display: flex; gap: 8px; align-items: center; }}
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
      <div class="metric"><span>Approved</span><strong>{approved}</strong></div>
      <div class="metric"><span>Disapproved</span><strong>{disapproved}</strong></div>
      <div class="metric"><span>Active</span><strong>{active}</strong></div>
    </section>
    <form id="bulk-delete-form" class="bulk-bar" method="post" action="/admin/bulk-delete" onsubmit="return confirmBulkDelete();">
      <label class="bulk-select">
        <input id="select-all-requests" type="checkbox">
        Select all visible
      </label>
      <button type="submit">Delete Selected</button>
      <span id="bulk-count" class="bulk-count">0 selected</span>
    </form>
    <nav class="admin-tabs" aria-label="Request tabs">
      <button class="admin-tab active" type="button" data-tab-filter="All">All</button>
      <button class="admin-tab" type="button" data-tab-filter="Request Tracker">Request Tracker</button>
      <button class="admin-tab" type="button" data-tab-filter="Work In Process">Work In Process</button>
      <button class="admin-tab" type="button" data-tab-filter="Approved">Approved</button>
      <button class="admin-tab" type="button" data-tab-filter="Disapproved">Disapproved</button>
      <button class="admin-tab" type="button" data-tab-filter="Designers">Designers</button>
    </nav>
    <section id="designer-panel" class="designer-panel">
      <form class="designer-form" method="post" action="/admin/designers/add">
        <label>Designer Name
          <input name="designer_name" required>
        </label>
        <label>Designer Email
          <input name="designer_email" type="email" required>
        </label>
        <button type="submit">Add Designer</button>
      </form>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Designer</th>
              <th>Email</th>
              <th>Manage</th>
            </tr>
          </thead>
          <tbody>{designer_rows}</tbody>
        </table>
      </div>
    </section>
    <section id="request-table-wrap" class="table-wrap">
      <table>
        <thead>
          <tr>
            <th class="select-cell"></th>
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
  <script>
    const selectAll = document.getElementById("select-all-requests");
    const checks = Array.from(document.querySelectorAll(".request-check"));
    const bulkCount = document.getElementById("bulk-count");
    const tabButtons = Array.from(document.querySelectorAll(".admin-tab"));
    const rows = Array.from(document.querySelectorAll("tbody tr[data-source-tab]"));

    function syncBulkCount() {{
      const visibleChecks = checks.filter((check) => !check.closest("tr").hidden);
      const selected = visibleChecks.filter((check) => check.checked).length;
      bulkCount.textContent = `${{selected}} selected`;
      selectAll.checked = selected > 0 && selected === visibleChecks.length;
      selectAll.indeterminate = selected > 0 && selected < visibleChecks.length;
    }}

    function applyTabFilter(tabName) {{
      rows.forEach((row) => {{
        const approval = row.dataset.approvalStatus || "";
        const seen = row.dataset.seenStatus || "";
        const isApproved = row.dataset.sourceTab === "Approved" || approval === "Approved";
        const isDisapproved = row.dataset.sourceTab === "Disapproved" || approval === "Not Approved" || approval === "Needs Info";
        const isUnseenPending = row.dataset.sourceTab === "Request Tracker" && seen !== "Seen" && !isApproved && !isDisapproved;
      const isVisible =
          tabName === "All" ||
          (tabName === "Request Tracker" && isUnseenPending) ||
          (tabName !== "Request Tracker" && row.dataset.sourceTab === tabName) ||
          (tabName === "Approved" && isApproved) ||
          (tabName === "Disapproved" && isDisapproved) ||
          (tabName === "Work In Process" && row.dataset.progressState === "In Process");
        row.hidden = !isVisible;
        if (!isVisible) {{
          const checkbox = row.querySelector(".request-check");
          if (checkbox) checkbox.checked = false;
        }}
      }});
      tabButtons.forEach((button) => {{
        button.classList.toggle("active", button.dataset.tabFilter === tabName);
      }});
      document.getElementById("designer-panel").classList.toggle("active", tabName === "Designers");
      document.getElementById("request-table-wrap").hidden = tabName === "Designers";
      document.getElementById("bulk-delete-form").hidden = tabName === "Designers";
      syncBulkCount();
    }}

    function confirmBulkDelete() {{
      const selected = checks.filter((check) => check.checked).length;
      if (!selected) {{
        alert("Select at least one request first.");
        return false;
      }}
      return confirm(`Delete ${{selected}} selected request${{selected === 1 ? "" : "s"}} from the admin dashboard?`);
    }}

    selectAll.addEventListener("change", () => {{
      checks.forEach((check) => {{
        if (!check.closest("tr").hidden) check.checked = selectAll.checked;
      }});
      syncBulkCount();
    }});
    checks.forEach((check) => check.addEventListener("change", syncBulkCount));
    tabButtons.forEach((button) => {{
      button.addEventListener("click", () => applyTabFilter(button.dataset.tabFilter));
    }});
    syncBulkCount();
  </script>
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
        if file_path.suffix.lower() == ".mp4":
            content_type = "video/mp4"
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
                designers = fetch_designers(config)
                status = "" if config["google_apps_script_webhook_url"] else APPS_SCRIPT_NOT_CONFIGURED_MESSAGE
                self.send_html(admin_dashboard_html(requests, status, error=not config["google_apps_script_webhook_url"], designers=designers))
            except Exception as exc:
                self.send_html(admin_dashboard_html([], str(exc), error=True), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/process":
            self.send_html(process_page_html())
            return
        if path == "/work-in-process":
            try:
                requests = fetch_admin_requests(get_config())
                self.send_html(public_work_page_html(requests))
            except Exception as exc:
                self.send_html(public_work_page_html([], str(exc), error=True), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/simple":
            query = parse_qs(urlparse(self.path).query)
            if query.get("payment", [""])[0] == "success":
                request_id = query.get("request_id", [""])[0]
                message = "Payment received / request submitted"
                if request_id:
                    message += f". Request ID: {request_id}"
                self.send_html(page_template(simple_form_html(message)))
                return
            self.send_html(page_template(simple_form_html()))
            return
        self.redirect("/simple")

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/square/create-checkout":
            length = int(self.headers.get("Content-Length", "0"))
            raw_body = self.rfile.read(length)
            try:
                payload = json.loads(raw_body.decode("utf-8") or "{}")
            except json.JSONDecodeError:
                self.send_json({"ok": False, "error": "Invalid checkout request."}, HTTPStatus.BAD_REQUEST)
                return
            submitted = payload.get("request") if isinstance(payload, dict) else {}
            if not isinstance(submitted, dict):
                self.send_json({"ok": False, "error": "Invalid checkout request."}, HTTPStatus.BAD_REQUEST)
                return
            data = normalize_simple_form({key: cell_text(value) for key, value in submitted.items()})
            errors = validate_simple_form(data)
            if errors:
                self.send_json({"ok": False, "error": " ".join(errors)}, HTTPStatus.BAD_REQUEST)
                return
            config = get_config()
            cart = calculate_checkout_cart(data)
            if cart.get("design_type_summary"):
                data["design_type"] = cart["design_type_summary"]
                data["project_name"] = f"{cart['design_type_summary']} Request"
            data["custom_cart_items"] = "[]"
            data["_payment_total_display"] = cart["total_display"]
            data["_payment_total_cents"] = str(cart["total_cents"])
            data["_rush_fee_included"] = cart["rush_included"]
            data["_cart_summary"] = cart["summary"]
            if cart["total_cents"] <= 0:
                data["_request_status"] = "Submitted"
                data["_payment_status"] = "No payment due"
                try:
                    result = save_submission(data)
                    try:
                        send_email(data, result)
                    except Exception as email_error:
                        print(f"Admin email could not be sent: {email_error}")
                    try:
                        send_client_confirmation_email(data, result)
                    except Exception as email_error:
                        print(f"Client confirmation email could not be sent: {email_error}")
                    self.send_json({
                        "ok": True,
                        "requestId": result["request_id"],
                        "checkoutUrl": square_success_url(config, result["request_id"]),
                        "noPaymentRequired": True,
                    })
                except Exception as exc:
                    self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
                return
            try:
                require_square_checkout_config(config)
                data["_request_status"] = "Pending Payment"
                data["_payment_status"] = "Pending"
                result = save_submission(data)
                try:
                    square_link = create_square_payment_link(data, result, cart, config)
                except Exception as square_error:
                    try:
                        data["_payment_status"] = f"Payment failed to start: {square_error}"
                        patch_request_fields(
                            result["request_id"],
                            {
                                "request_status": "Cancelled",
                                "admin_notes": combined_admin_notes(data),
                            },
                            config,
                        )
                    except Exception as patch_error:
                        print(f"Could not mark failed Square checkout request as cancelled: {patch_error}")
                    raise
                data["_square_payment_link_id"] = square_link["payment_link_id"]
                data["_square_order_id"] = square_link["order_id"]
                data["_square_checkout_url"] = square_link["checkout_url"]
                patch_request_fields(
                    result["request_id"],
                    {
                        "request_status": "Pending Payment",
                        "admin_notes": combined_admin_notes(data),
                    },
                    config,
                )
                self.send_json({
                    "ok": True,
                    "requestId": result["request_id"],
                    "checkoutUrl": square_link["checkout_url"],
                })
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/api/square/webhook":
            length = int(self.headers.get("Content-Length", "0"))
            raw_body = self.rfile.read(length)
            config = get_config()
            signature = self.headers.get("x-square-hmacsha256-signature", "")
            if not square_signature_is_valid(raw_body, signature, config):
                self.send_json({"ok": False, "error": "Invalid Square webhook signature."}, HTTPStatus.UNAUTHORIZED)
                return
            try:
                event = json.loads(raw_body.decode("utf-8") or "{}")
            except json.JSONDecodeError:
                self.send_json({"ok": False, "error": "Invalid webhook payload."}, HTTPStatus.BAD_REQUEST)
                return
            try:
                self.send_json(handle_square_webhook_event(event, config))
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
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
                before_designer = data.get("assigned_designer", "")
                update_admin_request(data, config)
                requests = fetch_admin_requests(config)
                designers = fetch_designers(config)
                assigned = next((designer for designer in designers if designer.get("name") == before_designer), None)
                request_item = next((item for item in requests if cell_text(item.get("request_id")) == data.get("request_id")), {})
                if assigned and request_item:
                    try:
                        send_designer_assignment_email(assigned, request_item)
                    except Exception as email_error:
                        print(f"Designer assignment email could not be sent: {email_error}")
                self.send_html(admin_dashboard_html(requests, f"{data.get('request_id', 'Request')} updated.", designers=designers))
            except Exception as exc:
                try:
                    requests = fetch_admin_requests(config)
                except Exception:
                    requests = []
                self.send_html(admin_dashboard_html(requests, str(exc), error=True, designers=fetch_designers(config)), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/admin/designers/add":
            if not self.is_admin_authorized():
                self.redirect("/admin")
                return
            length = int(self.headers.get("Content-Length", "0"))
            data = parse_form(self.rfile.read(length))
            config = get_config()
            try:
                add_designer(data, config)
                requests = fetch_admin_requests(config)
                designers = fetch_designers(config)
                self.send_html(admin_dashboard_html(requests, f"{data.get('designer_name', 'Designer')} added.", designers=designers))
            except Exception as exc:
                requests = fetch_admin_requests(config)
                self.send_html(admin_dashboard_html(requests, str(exc), error=True, designers=fetch_designers(config)), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/admin/designers/delete":
            if not self.is_admin_authorized():
                self.redirect("/admin")
                return
            length = int(self.headers.get("Content-Length", "0"))
            data = parse_form(self.rfile.read(length))
            config = get_config()
            try:
                delete_designer(data, config)
                requests = fetch_admin_requests(config)
                designers = fetch_designers(config)
                self.send_html(admin_dashboard_html(requests, "Designer removed.", designers=designers))
            except Exception as exc:
                requests = fetch_admin_requests(config)
                self.send_html(admin_dashboard_html(requests, str(exc), error=True, designers=fetch_designers(config)), HTTPStatus.INTERNAL_SERVER_ERROR)
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
                self.send_html(admin_dashboard_html(requests, f"{request_id} deleted.", designers=fetch_designers(config)))
            except Exception as exc:
                try:
                    requests = fetch_admin_requests(config)
                except Exception:
                    requests = []
                self.send_html(admin_dashboard_html(requests, str(exc), error=True, designers=fetch_designers(config)), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/admin/bulk-delete":
            if not self.is_admin_authorized():
                self.redirect("/admin")
                return
            length = int(self.headers.get("Content-Length", "0"))
            parsed = parse_qs(self.rfile.read(length).decode("utf-8"), keep_blank_values=True)
            request_ids = [value.strip() for value in parsed.get("request_id", []) if value.strip()]
            config = get_config()
            deleted = 0
            errors = []
            for request_id in request_ids:
                try:
                    delete_admin_request({"request_id": request_id}, config)
                    deleted += 1
                except Exception as exc:
                    errors.append(f"{request_id}: {exc}")
            try:
                requests = fetch_admin_requests(config)
            except Exception:
                requests = []
            if errors:
                message = f"Deleted {deleted} request(s). " + " ".join(errors)
                self.send_html(admin_dashboard_html(requests, message, error=True, designers=fetch_designers(config)), HTTPStatus.INTERNAL_SERVER_ERROR)
            else:
                self.send_html(admin_dashboard_html(requests, f"Deleted {deleted} selected request(s).", designers=fetch_designers(config)))
            return
        if path == "/process":
            length = int(self.headers.get("Content-Length", "0"))
            data = parse_form(self.rfile.read(length))
            request_id = data.get("request_id", "").strip().upper()
            email = data.get("email", "").strip().lower()
            if not request_id or not email:
                self.send_html(process_page_html(status="Request ID and email are required.", error=True), HTTPStatus.BAD_REQUEST)
                return
            try:
                requests = fetch_admin_requests(get_config())
                match = next(
                    (
                        item for item in requests
                        if cell_text(item.get("request_id")).upper() == request_id
                        and cell_text(item.get("email")).lower() == email
                    ),
                    None,
                )
                if not match:
                    self.send_html(process_page_html(status="No matching request found.", error=True), HTTPStatus.NOT_FOUND)
                    return
                self.send_html(process_page_html(match))
            except Exception as exc:
                self.send_html(process_page_html(status=str(exc), error=True), HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/simple-submit":
            self.send_html(
                page_template(simple_form_html("Square checkout is required. Please submit from the form to start checkout.", error=True)),
                HTTPStatus.BAD_REQUEST,
            )
            return
        if path == "/submit":
            self.send_html(
                page_template(simple_form_html("The full form is disabled. Please use the simple form checkout.", error=True)),
                HTTPStatus.GONE,
            )
            return
        self.send_error(HTTPStatus.NOT_FOUND)


def main() -> None:
    config = get_config()
    server = ThreadingHTTPServer((config["host"], int(config["port"])), RequestHandler)
    print(f"Creative request portal running at http://{config['host']}:{config['port']}")
    server.serve_forever()


handler = RequestHandler


if __name__ == "__main__":
    main()

