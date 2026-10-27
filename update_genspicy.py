import json, os, re, sys, base64
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = "1p0VIGOh016x599S3X6s_8t77pk8rsKGLbyN5B7lxhss"
SHEET_NAME = "genspicv"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def parse_date_cell(v):
    if v is None or v == "": return None
    if isinstance(v, (int, float)):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=float(v))
            return (dt.year, dt.month, dt.day)
        except: return None
    s = str(v).strip()
    m = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m: return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

def main():
    raw = os.environ["GOOGLE_CREDENTIALS"]
    try:
        creds_info = json.loads(raw)
    except Exception:
        decoded = base64.b64decode(raw).decode()
        creds_info = json.loads(decoded)

    with open("genspicy_data.json") as f:
        data = json.load(f)
    records = data.get("records", [])
    print(f"Loaded {len(records)} genspicy records")

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{SHEET_NAME}'!A:A"
    ).execute()
    existing = [(r[0] if r else "") for r in dv.get("values", [])]

    updates = []
    for rec in records:
        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]
        target_ymd = (y, m, d)

        target_row = None
        for i, v in enumerate(existing):
            if parse_date_cell(v) == target_ymd:
                target_row = i + 1
                break
        if target_row is None:
            print(f"  {date_str}: row not found, skip")
            continue

        for col, val in rec.items():
            if col == "date" or val is None:
                continue
            updates.append({
                "range": f"'{SHEET_NAME}'!{col}{target_row}",
                "values": [[val]]
            })
        print(f"  {date_str} -> row {target_row}: D={rec.get('D')}, E={rec.get('E')}, F={rec.get('F')}")

    if updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": updates}
        ).execute()
        print(f"Done: {len(updates)} cells updated")
    else:
        print("Nothing to update")

if __name__ == "__main__":
    main()
