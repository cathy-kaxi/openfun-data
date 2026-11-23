import json, os, base64
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
raw = os.environ["GOOGLE_CREDENTIALS"]
try:
    creds_info = json.loads(raw)
except Exception:
    creds_info = json.loads(base64.b64decode(raw).decode())

creds = Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
for s in meta["sheets"]:
    p = s["properties"]
    print("SHEET", p["sheetId"], p["title"], p.get("gridProperties"))

sheet = None
for s in meta["sheets"]:
    if s["properties"]["sheetId"] == 0:
        sheet = s
        break
name = sheet["properties"]["title"]
print("TARGET", name)

dv = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A1:AB90").execute()
vals = dv.get("values", [])
print("ROWS", len(vals))
for i, r in enumerate(vals):
    print(f"R{i+1}\t" + json.dumps(r, ensure_ascii=False))
