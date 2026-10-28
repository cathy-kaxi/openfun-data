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
    from collections import defaultdict
    date_record_count = defaultdict(int)
    for rec in records:
        date_record_count[rec["date"]] += 1

    date_existing_count = defaultdict(int)
    for v in existing:
        ymd = parse_date_cell(v)
        if ymd:
            date_existing_count[ymd] += 1

    extra_needed = {}
    for date_str, count in date_record_count.items():
        y, m, d = [int(x) for x in date_str.split("-")]
        ymd = (y, m, d)
        existing_count = date_existing_count.get(ymd, 0)
        if count > existing_count:
            extra_needed[ymd] = count - existing_count
            print(f"  Need {extra_needed[ymd]} extra rows for {date_str} (have {existing_count}, need {count})")

    # Step 3: Insert rows AND write dates to new rows
    # We'll collect all updates first, then do them in order
    row_updates = []  # list of (row_num_1based, col_letter, value)

    if extra_needed:
        # Collect insert positions
        insert_positions = []  # list of (insert_at_1based, num_rows, ymd)
        for ymd, num_extra in extra_needed.items():
            last_row = None
            for i, v in enumerate(existing):
                if parse_date_cell(v) == ymd:
                    last_row = i + 1
            if last_row is None:
                print(f"  Warning: date {ymd} not found in sheet, skip insert")
                continue
            insert_at = last_row + 1
            insert_positions.append((insert_at, num_extra, ymd))

        # Sort by insert position ascending (top to bottom) so we can track new rows
        insert_positions.sort(key=lambda x: x[0])

        # First, do all insertions
        insert_requests = []
        # We need to adjust positions because earlier inserts shift later ones
        # Let's calculate actual insert positions after all shifts
        # Simpler: insert from bottom to top (reverse order)
        for insert_at, num_extra, ymd in reversed(insert_positions):
            insert_requests.append({
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": insert_at - 1,  # 0-based
                        "endIndex": insert_at - 1 + num_extra
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

        # Now we need to write dates to the new rows
        # Re-read A column first to get accurate row numbers
        dv = svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:A").execute()
        existing = [(r[0] if r else "") for r in dv.get("values", [])]

        # For each date that needed extra rows, find the empty rows after it and fill dates
        for ymd, num_extra in extra_needed.items():
            y, m, d = ymd
            date_str = f"{y}/{m}/{d}"
            # Find all rows with this date
            date_rows = []
            for i, v in enumerate(existing):
                if parse_date_cell(v) == ymd:
                    date_rows.append(i + 1)
            # We have len(date_rows) rows with this date now
            # We need num_extra more rows filled with this date
            # Find empty rows right after the last date row
            last_date_row = max(date_rows) if date_rows else None
            if last_date_row is None:
                continue
            # Fill empty rows after last_date_row
            filled = 0
            for row_num in range(last_date_row + 1, len(existing) + 1 + num_extra):
                if filled >= num_extra:
                    break
                # Check if this row is empty
                cell_val = existing[row_num - 1] if row_num - 1 < len(existing) else ""
                if cell_val == "" or parse_date_cell(cell_val) is None:
                    row_updates.append((row_num, "A", date_str))
                    filled += 1
                    print(f"  Writing date {date_str} to row {row_num}")

    # Step 4: Re-read A column (after inserting and filling dates)
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:A").execute()
    existing = [(r[0] if r else "") for r in dv.get("values", [])]

    # Step 5: Match each record to a row and collect value updates
    date_used_rows = defaultdict(list)
    value_updates = []  # list of (row_num, col, value)

    for rec in records:
        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]
        target_ymd = (y, m, d)

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
            print(f"  {date_str}: row not found, skip")
            continue

        date_used_rows[target_ymd].append(target_row)

        for col, val in rec.items():
            if col in ("date", "line") or val is None:
                continue
            value_updates.append((target_row, col, val))
        line_label = rec.get("line", "")
        print(f"  {date_str} ({line_label}) -> row {target_row}")

    # Combine all updates (date fills + value updates)
    all_updates = row_updates + value_updates

    if all_updates:
        # Convert to the format expected by batchUpdate
        data_body = []
        for row_num, col, val in all_updates:
            data_body.append({
                "range": f"'{sheet_name}'!{col}{row_num}",
                "values": [[val]]
            })
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": data_body}
        ).execute()
        print(f"Done: {len(data_body)} cells updated")
    else:
        print("Nothing to update")

if __name__ == "__main__":
    main()
