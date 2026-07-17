import io
import os
import html
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests
import uvicorn
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, Response


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
SOLAR_CRM_IMPORT_URL = "https://app.solarcrm.com.au/backend/api/leads/import/excel/internal"
SOLAR_CRM_ORGANIZATION_ID = os.getenv("SOLAR_CRM_ORGANIZATION_ID", "2")


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
    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.body-content-type="text"',
        },
        timeout=60,
    )
    response.raise_for_status()
    return response.json().get("value", [])


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
    mobile_number = int(mobile_number) if mobile_number.isdigit() else None
    lead_date = message.get("receivedDateTime", "")
    if lead_date:
        lead_date = datetime.fromisoformat(
    lead_date.replace("Z", "+00:00")
).date()
    postcode = answers.get("What is your postcode?", "").strip()

    postcode = int(postcode) if postcode.isdigit() else None

    return {
        "Title": None,
        "First Name": name_parts[0] if name_parts else "",
        "Last Name": name_parts[1] if len(name_parts) > 1 else "",
        "Mobile Number": mobile_number,
        "Home Phone": None,
        "Email": answers.get("What is your email address?", ""),
        "Lead Source": "ProductReview",
        "Fax": None,
        "Notes": None,
        "Unit Type": None,
        "Unit Number": None,
        "Address Type": None,
        "Address": answers.get("What is the street address where the system will be installed?", ""),
        "Street Number": None,
        "Street Name": None,
        "Street Type": None,
        "Suburb": None,
        "Post Code": postcode,
        "State": state_from_postcode(postcode),
        "Lead Date": lead_date,
    }


def create_workbook(records):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(records).to_excel(writer, index=False, sheet_name="ProductReview Leads")
    output.seek(0)
    return output


def upload_to_solar_crm(filename, workbook):
    """Upload one generated Excel workbook to SolarCRM."""
    internal_secret = os.getenv("SOLAR_CRM_INTERNAL_SECRET")
    if not internal_secret:
        raise ValueError("SolarCRM is not configured. Set SOLAR_CRM_INTERNAL_SECRET in .env.")

    response = requests.post(
        SOLAR_CRM_IMPORT_URL,
        headers={
            "X-Internal-Secret": internal_secret,
            "OrganizationId": SOLAR_CRM_ORGANIZATION_ID,
        },
        files={
            "files[]": (
                filename,
                workbook,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
        timeout=60,
    )
    response.raise_for_status()
    return response


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
.preview { margin-top:24px; overflow-x:auto; border:1px solid var(--line); border-radius:4px; } table { width:100%; border-collapse:collapse; font-size:14px; } th,td { padding:10px 12px; text-align:left; border-bottom:1px solid var(--line); white-space:nowrap; } th { background:#f3f6f6; } tr:last-child td { border-bottom:0; } .export-actions { display:flex; gap:12px; margin-top:20px; } .export-actions form { flex:1; margin:0; padding:0; border:0; box-shadow:none; background:transparent; } .export-actions button { width:100%; }
@media (max-width:520px) { main { padding-top:48px; } h1 { font-size:28px; } form { padding:22px; } .dates { grid-template-columns:1fr; } }
</style></head><body><main><div class="brand"><span class="brand-mark">P</span>ProductReview Leads</div><h1>Export lead emails</h1><p>Select the inclusive date range for the ProductReview emails you want in the Excel file.</p><form method="post" action="/generate"><div class="dates"><label>From date<input type="date" name="start_date" required></label><label>To date<input type="date" name="end_date" required></label></div><button type="submit">Generate preview</button></form>{{CONTENT}}<div class="notice">The file includes leads received from the configured ProductReview sender.</div></main><script>document.querySelector('form').addEventListener('submit', function () { const b=this.querySelector('button'); b.disabled=true; b.textContent='Generating preview...'; });</script></body></html>"""

EXPORTS = {}
app = FastAPI(title="ProductReview Leads")


def render_page(content=""):
    return PAGE.replace("{{CONTENT}}", content)


def preview_html(records, token):
    rows = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(record.get(column, '')))}</td>" for column in ("First Name", "Last Name", "Email", "Mobile Number", "Post Code", "State", "Lead Date")) + "</tr>"
        for record in records
    )
    if not rows:
        rows = '<tr><td colspan="7">No leads found for this date range.</td></tr>'
    return f'''<div class="notice"><strong>Preview ready:</strong> {len(records)} lead(s) found. Review the results, then download when ready.</div>
<div class="preview"><table><thead><tr><th>First name</th><th>Last name</th><th>Email</th><th>Mobile</th><th>Postcode</th><th>State</th><th>Lead date</th></tr></thead><tbody>{rows}</tbody></table></div>
<div class="export-actions"><form method="post" action="/download"><input type="hidden" name="token" value="{token}"><button type="submit">Download Excel</button></form><form method="post" action="/upload"><input type="hidden" name="token" value="{token}"><button type="submit">Upload to SolarCRM</button></form></div>'''


@app.get("/", response_class=HTMLResponse)
def home():
    return render_page()


@app.head("/")
def home_head():
    """Respond successfully to health checks and other HEAD requests."""
    return Response(status_code=200)


@app.post("/generate", response_class=HTMLResponse)
def generate_preview(start_date: str = Form(...), end_date: str = Form(...)):
    try:
        start = datetime.strptime(start_date, "%Y-%m-%d").date()
        end = datetime.strptime(end_date, "%Y-%m-%d").date()
        if end < start:
            raise ValueError("The end date must be the same as or later than the start date.")

        records = [message_to_record(message) for message in get_messages(start, end)]
        workbook = create_workbook(records).getvalue()
        filename = f"ProductReview_Leads_{start:%Y%m%d}_{end:%Y%m%d}.xlsx"
        token = uuid.uuid4().hex
        EXPORTS[token] = (filename, workbook)
        return render_page(preview_html(records, token))
    except (ValueError, requests.RequestException, KeyError) as error:
        return HTMLResponse(
            render_page(f'<div class="notice error">{html.escape(str(error))}</div>'),
            status_code=400,
        )


@app.post("/download")
def download_excel(token: str = Form(...)):
    export = EXPORTS.get(token)
    if not export:
        return HTMLResponse(
            render_page('<div class="notice error">This preview has expired. Please generate it again.</div>'),
            status_code=400,
        )

    filename, workbook = export
    return Response(
        content=workbook,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/upload", response_class=HTMLResponse)
def upload_excel(token: str = Form(...)):
    export = EXPORTS.get(token)
    if not export:
        return HTMLResponse(
            render_page('<div class="notice error">This preview has expired. Please generate it again.</div>'),
            status_code=400,
        )

    filename, workbook = export
    try:
        upload_to_solar_crm(filename, workbook)
    except ValueError as error:
        return HTMLResponse(render_page(f'<div class="notice error">{html.escape(str(error))}</div>'), status_code=400)
    except requests.RequestException:
        return HTMLResponse(
            render_page('<div class="notice error">SolarCRM could not accept the file. Please try again.</div>'),
            status_code=502,
        )

    EXPORTS.pop(token, None)
    return render_page('<div class="notice"><strong>Upload complete.</strong> The Excel file has been sent to SolarCRM.</div>')


if __name__ == "__main__":
    # Render can route traffic only to services listening on all interfaces.
    host = "0.0.0.0"
    port = int(os.environ.get("PORT", 8000))
    print(f"Server listening on http://{host}:{port}")
    uvicorn.run(app, host=host, port=port)
