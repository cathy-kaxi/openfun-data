import json, os, base64, re
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]

ROWS = [
    {"date": "2026-11-21", "B": "CQJH", "C": "TT", "H": 97.26, "K": 4.98, "L": 1392, "M": 0.07, "N": 0.0713, "O": 0},
    {"date": "2026-11-22", "B": "CQJH", "C": "TT", "H": 101.01, "K": 8.97, "L": 885, "M": 0.11, "N": 0.0786, "O": 0},
]


def parse_date_cell(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        dt = datetime(1899, 12, 30) + timedelta(days=float(v))
        return (dt.year, dt.month, dt.day)
    s = str(v).strip()
    m = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m:
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


raw = os.environ["GOOGLE_CREDENTIALS"]
try:
    creds_info = json.loads(raw)
except Exception:
    creds_info = json.loads(base64.b64decode(raw).decode())

creds = Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
sheet = next(s for s in meta["sheets"] if s["properties"]["sheetId"] == 0)
name = sheet["properties"]["title"]
print("TARGET", name)

dv = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A1:AB90").execute()
vals = dv.get("values", [])

updates = []
for rec in ROWS:
    y, m, d = [int(x) for x in rec["date"].split("-")]
    target = (y, m, d)
    row_num = None
    for i, r in enumerate(vals):
        if parse_date_cell(r[0] if r else "") == target:
            row_num = i + 1
            break
    if row_num is None:
        last = 2
        for i, r in enumerate(vals):
            if parse_date_cell(r[0] if r else ""):
                last = i + 1
        row_num = last + 1
        updates.append({"range": f"'{name}'!A{row_num}", "values": [[rec["date"].replace("-", "/")]]})
        print(f"  new row {row_num} for {rec['date']}")
    else:
        print(f"  existing row {row_num} for {rec['date']}")

    for col, val in rec.items():
        if col == "date":
            continue
        updates.append({"range": f"'{name}'!{col}{row_num}", "values": [[val]]})
    print(f"  wrote {rec['date']} -> row {row_num}")

if updates:
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": updates},
    ).execute()
    print(f"BATCH OK, {len(updates)} cells")

dv2 = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A1:AB45").execute()
for i, r in enumerate(dv2.get("values", [])):
    dt = parse_date_cell(r[0] if r else "")
    if dt and dt in [(2026, 11, 21), (2026, 11, 22)]:
        print(f"VERIFY R{i+1}\t" + json.dumps(r, ensure_ascii=False))
