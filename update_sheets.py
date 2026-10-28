import json, os, re, sys, base64
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
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
    print(f"Length: {len(raw)}")
    try:
        creds_info = json.loads(raw)
    except Exception as e:
        print(f"Direct json.loads failed: {e}")
        try:
            decoded = base64.b64decode(raw).decode()
            creds_info = json.loads(decoded)
        except Exception as e2:
            print(f"Base64 also failed: {e2}")
            sys.exit(1)

    with open("data.json") as f:
        data = json.load(f)
    records = data.get("records", [])
    print(f"Loaded {len(records)} records")

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet = meta["sheets"][0]
    sheet_name = sheet["properties"]["title"]
    sheet_id = sheet["properties"]["sheetId"]
    print(f"Sheet: {sheet_name} (id={sheet_id})")

    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:A").execute()
    existing = [(r[0] if r else "") for r in dv.get("values", [])]

    # Track which rows have been used per date
    date_used_rows = {}  # date_ymd -> list of row numbers already assigned

    # First pass: determine target rows for each record
    # If we need to insert a new row, collect all insert requests first
    inserts_needed = []  # list of row positions to insert (1-based)

    record_target_rows = []

    for rec in records:
        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]
        target_ymd = (y, m, d)

        # Find first matching row not yet used for this date
        target_row = None
        used_rows = date_used_rows.get(target_ymd, [])
        for i, v in enumerate(existing):
            row_num = i + 1
            if row_num in used_rows:
                continue
            if parse_date_cell(v) == target_ymd:
                target_row = row_num
                break

        if target_row is None:
            # Need a new row: find last row with this date, insert after it
            last_row = None
            for i, v in enumerate(existing):
                if parse_date_cell(v) == target_ymd:
                    last_row = i + 1
            if last_row is None:
                print(f"  {date_str}: row not found, skip")
                record_target_rows.append(None)
                continue
            # We need to insert a row at position last_row + 1 (1-based)
            insert_at = last_row + 1
            inserts_needed.append(insert_at)
            target_row = insert_at
            # Add a placeholder in existing list so subsequent inserts adjust correctly
            existing.insert(last_row, "")  # empty string as placeholder

        if target_ymd not in date_used_rows:
            date_used_rows[target_ymd] = []
        date_used_rows[target_ymd].append(target_row)
        record_target_rows.append(target_row)

    print(f"Inserts needed at rows: {inserts_needed}")

    # Perform row inserts first (in reverse order so positions don't shift)
    if inserts_needed:
        insert_requests = []
        # Sort descending so we insert from bottom up
        for row_pos in sorted(set(inserts_needed), reverse=True):
            # row_pos is 1-based, insertDimension uses 0-based index
            insert_requests.append({
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row_pos - 1,  # 0-based
                        "endIndex": row_pos  # insert 1 row
                    },
                    "inheritFromBefore": False
                }
            })
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": insert_requests}
        ).execute()
        print(f"Inserted {len(insert_requests)} rows")

    # Now write all the data
    updates = []
    for idx, rec in enumerate(records):
        target_row = record_target_rows[idx]
        if target_row is None:
            continue

        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]

        # If this record is in a newly inserted row, write the date first
        # Check if this row was an insert
        # Actually, we need to know if this row is new. Let's track it.
        # For simplicity: if row's A cell is empty (placeholder), write date
        # We'll just always write date for records that are inserts

        for col, val in rec.items():
            if col in ("date", "line") or val is None:
                continue
            updates.append({
                "range": f"'{sheet_name}'!{col}{target_row}",
                "values": [[val]]
            })
        line_label = rec.get("line", "")
        print(f"  {date_str} ({line_label}) -> row {target_row}")

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
