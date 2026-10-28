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
    existing_values = dv.get("values", [])
    existing = [(r[0] if r else "") for r in existing_values]
    print(f"Existing rows with values: {len(existing)}")

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

    extra_needed = {}  # ymd -> num_extra_rows
    for date_str, count in date_record_count.items():
        y, m, d = [int(x) for x in date_str.split("-")]
        ymd = (y, m, d)
        existing_count = date_existing_count.get(ymd, 0)
        if count > existing_count:
            extra_needed[ymd] = count - existing_count
            print(f"  Need {extra_needed[ymd]} extra rows for {date_str} (have {existing_count}, need {count})")

    # Step 3: Calculate insert positions
    # For each date needing extra rows, find the last row of that date
    # We insert num_extra rows right after the last row of that date
    insert_ops = []  # list of (insert_at_1based, num_rows, ymd)

    for ymd, num_extra in extra_needed.items():
        last_row = None
        for i, v in enumerate(existing):
            if parse_date_cell(v) == ymd:
                last_row = i + 1  # 1-based
        if last_row is None:
            print(f"  Warning: date {ymd} not found in sheet, skip")
            continue
        insert_at = last_row + 1
        insert_ops.append((insert_at, num_extra, ymd))
        y, m, d = ymd
        print(f"  Plan: insert {num_extra} row(s) at position {insert_at} (after {y}-{m}-{d})")

    # Step 4: Perform row inserts (bottom to top so positions don't shift)
    if insert_ops:
        insert_requests = []
        for insert_at, num_rows, ymd in reversed(insert_ops):
            insert_requests.append({
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": insert_at - 1,  # 0-based
                        "endIndex": insert_at - 1 + num_rows
                    },
                    "inheritFromBefore": False
                }
            })

        if insert_requests:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={"requests": insert_requests}
            ).execute()
            print(f"Done: inserted {len(insert_requests)} row groups")

    # Step 5: Now read A column again (after inserts)
    # We need to read more rows to capture the newly inserted empty rows
    # Let's read a large range (e.g., 100 rows)
    max_rows_to_read = 100
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A1:A{max_rows_to_read}").execute()
    existing_values = dv.get("values", [])
    # Pad to max_rows_to_read
    existing = []
    for i in range(max_rows_to_read):
        if i < len(existing_values) and existing_values[i]:
            existing.append(existing_values[i][0] if existing_values[i] else "")
        else:
            existing.append("")
    print(f"Read {len(existing)} rows from A column")

    # Step 6: Fill empty rows with the date from the row above
    # Scan top to bottom, track the last seen date
    date_fills = []  # list of (row_num, date_str)
    last_date = None
    for i, v in enumerate(existing):
        row_num = i + 1  # 1-based
        current_date = parse_date_cell(v)
        if current_date:
            last_date = current_date
        else:
            # This row is empty, fill with last_date
            if last_date:
                y, m, d = last_date
                date_str = f"{y}/{m}/{d}"
                date_fills.append({
                    "range": f"'{sheet_name}'!A{row_num}",
                    "values": [[date_str]]
                })
                print(f"  Filling row {row_num} with date {date_str}")

    if date_fills:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": date_fills}
        ).execute()
        print(f"Done: filled {len(date_fills)} empty rows with dates")

    # Step 7: Re-read A column (now all rows have dates)
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A1:A{max_rows_to_read}").execute()
    existing_values = dv.get("values", [])
    existing = []
    for i in range(max_rows_to_read):
        if i < len(existing_values) and existing_values[i]:
            existing.append(existing_values[i][0] if existing_values[i] else "")
        else:
            existing.append("")

    # Step 8: Match each record to a row and write values
    date_used_rows = defaultdict(list)
    value_updates = []

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
            value_updates.append({
                "range": f"'{sheet_name}'!{col}{target_row}",
                "values": [[val]]
            })
        line_label = rec.get("line", "")
        print(f"  {date_str} ({line_label}) -> row {target_row}")

    if value_updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": value_updates}
        ).execute()
        print(f"Done: {len(value_updates)} cells updated")
    else:
        print("No values to update")

if __name__ == "__main__":
    main()
