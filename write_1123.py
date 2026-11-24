import json, os, base64, re
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]

REC = {"date": "2026-11-23", "B": "CQJH", "C": "TT", "H": 87.43, "O": 0}


def parse_date_cell(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        dt = datetime(1899, 12, 30) + timedelta(days=float(v))
        return (dt.year, dt.month, dt.day)
    m = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(v).strip())
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
name, sid = sheet["properties"]["title"], sheet["properties"]["sheetId"]
print("TARGET", name, sid)

vals = svc.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A1:AB90").execute().get("values", [])

y, m, d = [int(x) for x in REC["date"].split("-")]
row_num = None
for i, r in enumerate(vals):
    if parse_date_cell(r[0] if r else "") == (y, m, d):
        row_num = i + 1
        break

new_row = False
if row_num is None:
    last = 2
    for i, r in enumerate(vals):
        if parse_date_cell(r[0] if r else ""):
            last = i + 1
    row_num = last + 1
    new_row = True

print("row", row_num, "new" if new_row else "existing")

if new_row:
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "copyPaste": {
                "source": {"sheetId": sid, "startRowIndex": row_num - 2, "endRowIndex": row_num - 1,
                           "startColumnIndex": 0, "endColumnIndex": 28},
                "destination": {"sheetId": sid, "startRowIndex": row_num - 1, "endRowIndex": row_num,
                                "startColumnIndex": 0, "endColumnIndex": 28},
                "pasteType": "PASTE_FORMAT",
            }
        }]},
    ).execute()
    print("format copied from row", row_num - 1)

updates = [{"range": f"'{name}'!A{row_num}", "values": [[REC["date"].replace("-", "/")]]}]
for col, val in REC.items():
    if col == "date":
        continue
    updates.append({"range": f"'{name}'!{col}{row_num}", "values": [[val]]})
updates.append({"range": f"'{name}'!Q{row_num}", "values": [[f"=IFERROR(O{row_num}/L{row_num},0)"]]})

svc.spreadsheets().values().batchUpdate(
    spreadsheetId=SPREADSHEET_ID,
    body={"valueInputOption": "USER_ENTERED", "data": updates},
).execute()
print("cells written", len(updates))

dv = svc.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A32:AB{row_num + 2}").execute()
for i, r in enumerate(dv.get("values", [])):
    print(f"R{i+32}\t" + json.dumps(r, ensure_ascii=False))
