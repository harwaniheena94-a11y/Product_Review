import io
import os
from uuid import uuid4
from datetime import datetime, timedelta, timezone
from html import escape
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs

import pandas as pd
import requests


def load_environment_file():
    """Load local configuration without requiring an additional package."""
    base_directory = Path(__file__).resolve().parent
    for filename in (".env", ".env.example"):
        path = base_directory / filename
        if not path.exists():
            continue
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))
        return


load_environment_file()

USER_EMAIL = os.getenv("PRODUCT_REVIEW_USER_EMAIL", "leads@sunboost.com.au")
SENDER = os.getenv("PRODUCT_REVIEW_SENDER", "dev@productreview.com.au")
EXPORT_COLUMNS = (
    "Title", "First Name", "Last Name", "Mobile Number", "Home Phone", "Email",
    "Lead Source", "Fax", "Notes", "Unit Type", "Unit Number", "Address Type",
    "Address", "Street Number", "Street Name", "Street Type", "Suburb", "Post Code",
    "State", "Lead Date",
)
PREVIEWS = {}


def get_credentials():
    credentials = {
        "tenant_id": os.getenv("AZURE_TENANT_ID"),
        "client_id": os.getenv("AZURE_CLIENT_ID"),
        "client_secret": os.getenv("AZURE_CLIENT_SECRET"),
    }
    if not all(credentials.values()):
        raise ValueError(
            "Missing Azure credentials. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET."
        )
    return credentials


def get_access_token(credentials):
    response = requests.post(
        f"https://login.microsoftonline.com/{credentials['tenant_id']}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": credentials["client_id"],
            "client_secret": credentials["client_secret"],
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]


def get_messages(start_date, end_date):
    credentials = get_credentials()
    token = get_access_token(credentials)
    start = datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc)
    end = datetime.combine(end_date + timedelta(days=1), datetime.min.time(), tzinfo=timezone.utc)
    start_str = start.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end.strftime("%Y-%m-%dT%H:%M:%SZ")

    url = (
        f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/mailFolders/inbox/messages"
        f"?$filter=receivedDateTime ge {start_str} and receivedDateTime lt {end_str} "
        f"and from/emailAddress/address eq '{SENDER}' and isRead eq false"
        "&$select=subject,receivedDateTime,body,from"
        "&$orderby=receivedDateTime desc"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": 'outlook.body-content-type="text"',
    }
    messages = []
    while url:
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()
        payload = response.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return messages


def state_from_postcode(postcode):
    try:
        value = int(str(postcode).strip())
    except (TypeError, ValueError):
        return ""

    if 200 <= value <= 299 or 2600 <= value <= 2619 or 2900 <= value <= 2920:
        return "ACT"
    if 800 <= value <= 999:
        return "NT"
    if 1000 <= value <= 1999 or 2000 <= value <= 2599 or 2620 <= value <= 2899:
        return "NSW"
    if 3000 <= value <= 3999 or 8000 <= value <= 8999:
        return "VIC"
    if 4000 <= value <= 4999 or 9000 <= value <= 9999:
        return "QLD"
    if 5000 <= value <= 5999:
        return "SA"
    if 6000 <= value <= 6999:
        return "WA"
    if 7000 <= value <= 7999:
        return "TAS"
    return ""


def message_to_record(message):
    body = message.get("body", {}).get("content", "")
    lines = [line.strip() for line in body.splitlines() if line.strip()]
    answers = {lines[i]: lines[i + 1] for i in range(len(lines) - 1)}
    name_parts = answers.get("What is your name?", "").split(maxsplit=1)
    mobile_number = answers.get("What is your mobile number?", "").replace(" ", "")
    if mobile_number.startswith("+61"):
        mobile_number = "" + mobile_number[3:]
    lead_date = message.get("receivedDateTime", "")
    if lead_date:
        lead_date = datetime.fromisoformat(lead_date.replace("Z", "+00:00")).strftime("%m/%d/%Y")
    postcode = answers.get("What is your postcode?", "")

    return {
        "Title": "",
        "First Name": name_parts[0] if name_parts else "",
        "Last Name": name_parts[1] if len(name_parts) > 1 else "",
        "Mobile Number": mobile_number,
        "Home Phone": "",
        "Email": answers.get("What is your email address?", ""),
        "Lead Source": "ProductReview",
        "Fax": "",
        "Notes": "",
        "Unit Type": "",
        "Unit Number": "",
        "Address Type": "",
        "Address": answers.get("What is the street address where the system will be installed?", ""),
        "Street Number": "",
        "Street Name": "",
        "Street Type": "",
        "Suburb": "",
        "Post Code": postcode,
        "State": state_from_postcode(postcode),
        "Lead Date": lead_date,
    }


def create_workbook(records):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(records, columns=EXPORT_COLUMNS).to_excel(writer, index=False, sheet_name="ProductReview Leads")
    output.seek(0)
    return output


PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ProductReview Leads</title><style>
:root { color-scheme: light; --ink:#13212c; --muted:#61717c; --line:#d7e0e5; --accent:#007c78; --accent-dark:#005e5b; --surface:#fff; --bg:#f3f6f6; }
* { box-sizing:border-box } body { margin:0; min-height:100vh; background:var(--bg); color:var(--ink); font:16px/1.5 Arial,sans-serif; }
main { width:min(700px,calc(100% - 32px)); margin:0 auto; padding:clamp(48px,10vh,100px) 0; }
.brand { display:flex; gap:12px; align-items:center; color:var(--accent); font-weight:700; letter-spacing:0; }
.brand-mark { display:grid; place-items:center; width:32px; height:32px; background:var(--accent); color:#fff; font-size:20px; }
h1 { margin:32px 0 8px; font-size:32px; line-height:1.15; letter-spacing:0; } p { color:var(--muted); margin:0; }
form { margin-top:32px; padding:28px; border:1px solid var(--line); border-radius:8px; background:var(--surface); box-shadow:0 10px 30px rgba(19,33,44,.06); }
.dates { display:grid; grid-template-columns:1fr 1fr; gap:18px; margin:24px 0 28px; } label { font-weight:700; font-size:14px; } input { width:100%; margin-top:8px; border:1px solid #aebcc4; border-radius:4px; padding:12px; color:var(--ink); font:inherit; } input:focus { outline:3px solid #bce6e2; border-color:var(--accent); }
button { width:100%; border:0; border-radius:4px; background:var(--accent); color:#fff; padding:13px 18px; font:700 16px Arial,sans-serif; cursor:pointer; } button:hover { background:var(--accent-dark); } button:disabled { opacity:.7; cursor:wait; }
.notice { margin-top:20px; padding:12px; border-left:3px solid var(--accent); background:#edf8f7; color:#285653; font-size:14px; } .error { border-color:#b42318; background:#fff2f0; color:#8a1c15; }
@media (max-width:520px) { main { padding-top:48px; } h1 { font-size:28px; } form { padding:22px; } .dates { grid-template-columns:1fr; } }
</style></head><body><main><div class="brand"><span class="brand-mark">P</span>ProductReview Leads</div><h1>Preview lead emails</h1><p>Select the inclusive date range for the unread ProductReview emails you want to review.</p><form method="post" action="/preview"><div class="dates"><label>From date<input type="date" name="start_date" max="__MAX_DATE__" required></label><label>To date<input type="date" name="end_date" max="__MAX_DATE__" required></label></div><button type="submit">Generate preview</button></form><div class="notice">The file includes only unread leads received from the configured ProductReview sender.</div></main><script>document.querySelector('form').addEventListener('submit', function () { const b=this.querySelector('button'); b.disabled=true; b.textContent='Generating preview...'; });</script></body></html>"""


def render_page(error_message=""):
    page = PAGE.replace("__MAX_DATE__", datetime.now().date().isoformat())
    if error_message:
        error_html = f'<div class="notice error">{escape(error_message)}</div>'
        return page.replace('<div class="notice">', f"{error_html}<div class=\"notice\">")
    return page


def render_preview(records, start_date, end_date, preview_id):
    headers = "".join(f"<th>{escape(column)}</th>" for column in EXPORT_COLUMNS)
    rows = "".join(
        "<tr>" + "".join(f"<td>{escape(str(record.get(column, '')))}</td>" for column in EXPORT_COLUMNS) + "</tr>"
        for record in records
    )
    table = f"<div class=\"table-wrap\"><table><thead><tr>{headers}</tr></thead><tbody>{rows}</tbody></table></div>" if records else "<div class=\"notice\">No unread ProductReview leads were found for this date range.</div>"
    download = f'<form method="post" action="/download"><input type="hidden" name="preview_id" value="{preview_id}"><button type="submit">Download Excel</button></form>' if records else ""
    return f"""<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>Lead Preview</title><style>
    :root {{ --ink:#13212c; --muted:#61717c; --line:#d7e0e5; --accent:#007c78; --accent-dark:#005e5b; --surface:#fff; --bg:#f3f6f6; }} * {{ box-sizing:border-box }} body {{ margin:0; background:var(--bg); color:var(--ink); font:16px/1.5 Arial,sans-serif }} main {{ width:min(1440px,calc(100% - 32px)); margin:0 auto; padding:48px 0 }} .brand {{ color:var(--accent); font-weight:700 }} h1 {{ margin:22px 0 4px; font-size:30px }} p {{ margin:0; color:var(--muted) }} .actions {{ display:flex; gap:12px; align-items:center; margin:26px 0 18px }} a {{ color:var(--accent); font-weight:700; text-decoration:none }} form {{ margin:0 }} button {{ border:0; border-radius:4px; background:var(--accent); color:#fff; padding:12px 20px; font:700 15px Arial,sans-serif; cursor:pointer }} button:hover {{ background:var(--accent-dark) }} .table-wrap {{ overflow:auto; border:1px solid var(--line); border-radius:8px; background:var(--surface) }} table {{ width:100%; min-width:1300px; border-collapse:collapse; font-size:14px }} th {{ background:#e8f4f2; text-align:left; white-space:nowrap; font-weight:700 }} th,td {{ padding:11px 12px; border-bottom:1px solid var(--line); vertical-align:top }} td {{ max-width:280px; overflow-wrap:anywhere }} tr:last-child td {{ border-bottom:0 }} .notice {{ padding:14px; border-left:3px solid var(--accent); background:#edf8f7; color:#285653 }} @media (max-width:520px) {{ main {{ width:calc(100% - 24px); padding-top:32px }} .actions {{ align-items:stretch; flex-direction:column }} button {{ width:100% }} }}</style></head><body><main><div class=\"brand\">ProductReview Leads</div><h1>Lead preview</h1><p>{len(records)} unread lead{'s' if len(records) != 1 else ''} from {start_date:%m/%d/%Y} to {end_date:%m/%d/%Y}.</p><div class=\"actions\"><a href=\"/\">Change dates</a>{download}</div>{table}</main></body></html>"""


def parse_date_range(data):
    start_date = datetime.strptime(data.get("start_date", [""])[0], "%Y-%m-%d").date()
    end_date = datetime.strptime(data.get("end_date", [""])[0], "%Y-%m-%d").date()
    if end_date < start_date:
        raise ValueError("The end date must be the same as or later than the start date.")
    today = datetime.now().date()
    if start_date > today or end_date > today:
        raise ValueError("Invalid date. Please select today or an earlier date.")
    return start_date, end_date


class LeadExportHandler(BaseHTTPRequestHandler):
    def send_html(self, content, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content.encode())))
        self.end_headers()
        self.wfile.write(content.encode())

    def do_GET(self):
        self.send_html(render_page() if self.path == "/" else "Not found", 200 if self.path == "/" else 404)

    def do_POST(self):
        if self.path not in {"/preview", "/download"}:
            self.send_html("Not found", 404)
            return
        try:
            size = int(self.headers.get("Content-Length", 0))
            data = parse_qs(self.rfile.read(size).decode())
            if self.path == "/preview":
                start_date, end_date = parse_date_range(data)
                records = [message_to_record(message) for message in get_messages(start_date, end_date)]
                preview_id = uuid4().hex
                PREVIEWS[preview_id] = {"records": records, "start_date": start_date, "end_date": end_date}
                self.send_html(render_preview(records, start_date, end_date, preview_id))
                return

            preview_id = data.get("preview_id", [""])[0]
            preview = PREVIEWS.get(preview_id)
            if not preview:
                raise ValueError("This preview has expired. Generate a new preview before downloading.")
            workbook = create_workbook(preview["records"])
            start_date = preview["start_date"]
            end_date = preview["end_date"]
            filename = f"ProductReview_Leads_{start_date:%Y%m%d}_{end_date:%Y%m%d}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(workbook.getvalue())))
            self.end_headers()
            self.wfile.write(workbook.getvalue())
        except (ValueError, requests.RequestException, KeyError) as error:
            self.send_html(render_page(str(error)), 400)
        except Exception:
            self.log_error("Unexpected error while generating lead export")
            self.send_html(render_page("Unable to generate the Excel file. Check the server terminal and try again."), 500)


if __name__ == "__main__":
    host = "0.0.0.0"
    port = int(os.environ.get("PORT", 8000))

    server = HTTPServer((host, port), LeadExportHandler)

    print(f"Server running on http://{host}:{port}")

    server.serve_forever()
