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

    # Step 1: Read existing A column
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:A").execute()
    existing = [(r[0] if r else "") for r in dv.get("values", [])]
    print(f"Existing rows: {len(existing)}")

    # Step 2: Determine which dates need extra rows
    # Count records per date
    from collections import defaultdict
    date_record_count = defaultdict(int)
    for rec in records:
        date_record_count[rec["date"]] += 1

    # Count existing rows per date
    date_existing_count = defaultdict(int)
    for v in existing:
        ymd = parse_date_cell(v)
        if ymd:
            date_existing_count[ymd] += 1

    # Determine how many extra rows we need per date
    # extra_needed: date_ymd -> number of rows to insert
    extra_needed = {}
    for date_str, count in date_record_count.items():
        y, m, d = [int(x) for x in date_str.split("-")]
        ymd = (y, m, d)
        existing_count = date_existing_count.get(ymd, 0)
        if count > existing_count:
            extra_needed[ymd] = count - existing_count
            print(f"  Need {extra_needed[ymd]} extra rows for {date_str} (have {existing_count}, need {count})")

    # Step 3: Insert rows
    # For each date that needs extra rows, find the last row of that date,
    # and insert after it
    if extra_needed:
        insert_requests = []
        # We need to insert from bottom to top so positions don't shift
        # Collect all insert positions first
        insert_positions = []  # list of (row_pos_1based, num_rows)
        for ymd, num_extra in extra_needed.items():
            # Find last row with this date
            last_row = None
            for i, v in enumerate(existing):
                if parse_date_cell(v) == ymd:
                    last_row = i + 1  # 1-based
            if last_row is None:
                print(f"  Warning: date {ymd} not found in sheet, skip insert")
                continue
            insert_at = last_row + 1  # insert after this row
            insert_positions.append((insert_at, num_extra, ymd))

        # Sort by insert position descending (bottom to top)
        insert_positions.sort(key=lambda x: x[0], reverse=True)

        for insert_at, num_extra, ymd in insert_positions:
            # insertDimension uses 0-based index
            # We want to insert num_extra rows starting at insert_at (1-based)
            # which is index insert_at - 1 (0-based)
            insert_requests.append({
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": insert_at - 1,  # 0-based
                        "endIndex": insert_at - 1 + num_extra  # insert num_extra rows
                    },
                    "inheritFromBefore": False
                }
            })
            y, m, d = ymd
            print(f"  Inserting {num_extra} row(s) at position {insert_at} (after {y}-{m}-{d})")

        if insert_requests:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={"requests": insert_requests}
            ).execute()
            print(f"Done: inserted {len(insert_requests)} row groups")

    # Step 4: Re-read A column (after inserts)
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:A").execute()
    existing = [(r[0] if r else "") for r in dv.get("values", [])]

    # Step 5: Now match each record to a row
    # Track which rows have been used per date
    date_used_rows = defaultdict(list)
    updates = []

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
            print(f"  {date_str}: row not found after insert, skip")
            continue

        # Mark this row as used
        date_used_rows[target_ymd].append(target_row)

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
