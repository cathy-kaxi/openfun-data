import json, os, re, sys, base64
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = "1p0VIGOh016x599S3X6s_8t77pk8rsKGLbyN5B7lxhss"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

debug = {}

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
    debug["loaded_records"] = len(records)

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    # 获取所有 sheet 名字
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = [s["properties"]["title"] for s in meta["sheets"]]
    debug["all_sheets"] = sheets
    print(f"All sheets: {sheets}")

    # 查找目标 sheet：优先找包含 spicy 的，找不到就用第一个
    target_sheet = None
    for s in sheets:
        if "spicy" in s.lower():
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = sheets[0]
    debug["target_sheet"] = target_sheet
    print(f"Using sheet: {target_sheet}")

    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{target_sheet}'!A:A"
    ).execute()
    existing = [(r[0] if r else "") for r in dv.get("values", [])]
    debug["existing_dates_count"] = len(existing)
    debug["existing_dates_sample"] = existing[:15]
    print(f"Existing dates: {len(existing)} rows")

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
            debug[f"row_not_found_{date_str}"] = True
            print(f"  {date_str}: row not found, skip")
            continue

        for col, val in rec.items():
            if col == "date" or val is None:
                continue
            updates.append({
                "range": f"'{target_sheet}'!{col}{target_row}",
                "values": [[val]]
            })
        debug[f"matched_{date_str}"] = f"row {target_row}"
        print(f"  {date_str} -> row {target_row}: D={rec.get('D')}, E={rec.get('E')}, F={rec.get('F')}")

    if updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": updates}
        ).execute()
        debug["updated_cells"] = len(updates)
        print(f"Done: {len(updates)} cells updated")
    else:
        debug["updated_cells"] = 0
        print("Nothing to update")

    # 写调试文件
    with open("genspicy_debug.json", "w") as f:
        json.dump(debug, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
