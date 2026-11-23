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
sheet = next(s for s in meta["sheets"] if s["properties"]["sheetId"] == 0)
name = sheet["properties"]["title"]
sid = sheet["properties"]["sheetId"]
print("TARGET", name, sid)

svc.spreadsheets().batchUpdate(
    spreadsheetId=SPREADSHEET_ID,
    body={"requests": [{
        "copyPaste": {
            "source": {"sheetId": sid, "startRowIndex": 32, "endRowIndex": 33,
                       "startColumnIndex": 0, "endColumnIndex": 28},
            "destination": {"sheetId": sid, "startRowIndex": 33, "endRowIndex": 34,
                            "startColumnIndex": 0, "endColumnIndex": 28},
            "pasteType": "PASTE_FORMAT",
        }
    }]},
).execute()
print("format copied")

upd = []
for r in (33, 34):
    upd.append({"range": f"'{name}'!Q{r}", "values": [[f"=IFERROR(O{r}/L{r},0)"]]})
svc.spreadsheets().values().batchUpdate(
    spreadsheetId=SPREADSHEET_ID,
    body={"valueInputOption": "USER_ENTERED", "data": upd},
).execute()
print("Q formulas restored")

dv = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A30:AB35").execute()
for i, r in enumerate(dv.get("values", [])):
    print(f"R{i+30}\t" + json.dumps(r, ensure_ascii=False))
