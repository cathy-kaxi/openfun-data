import json, os, re, sys, base64
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TARGET_SHEET_ID = 0  # Al-CAI-11 sheet has gid=0

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

def parse_line(v):
    """Read G column line value, return '11' or '08'"""
    if v is None or v == "": return ""
    s = str(v).strip()
    if s == "11": return "11"
    if s in ("8", "08"): return "08"
    return s

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
    # Find target sheet by sheetId=0
    sheet = None
    for s in meta["sheets"]:
        if s["properties"]["sheetId"] == TARGET_SHEET_ID:
            sheet = s
            break
    if not sheet:
        print("ERROR: target sheet id=0 not found!")
        sys.exit(1)
    sheet_name = sheet["properties"]["title"]
    sheet_id = sheet["properties"]["sheetId"]
    print(f"Sheet: {sheet_name} (id={sheet_id})")

    # Step 1: Read existing A:G columns (need G for line matching)
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:G").execute()
    existing_values = dv.get("values", [])
    print(f"Existing rows in A:G: {len(existing_values)}")

    # DEBUG: print rows around 10/28
    print("=== DEBUG: rows 15-20 (date | G line) ===")
    for i in range(14, min(20, len(existing_values))):
        row = existing_values[i]
        date_v = row[0] if len(row) > 0 else ""
        g_v = row[6] if len(row) > 6 else ""
        print(f"  row {i+1}: date={date_v!r}, G={g_v!r} -> parsed_date={parse_date_cell(date_v)}, line={parse_line(g_v)}")
    print("=== END DEBUG ===")

    # Step 2: Determine which dates need extra rows
    from collections import defaultdict
    date_record_count = defaultdict(int)
    for rec in records:
        date_record_count[rec["date"]] += 1

    date_existing_count = defaultdict(int)
    for v in existing_values:
        ymd = parse_date_cell(v[0] if v else "")
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

    # Step 3: For each date needing extra rows, find the LAST row of that date
    insert_ops = []  # list of (original_insert_at_row_1based, num_rows, ymd)

    for ymd, num_extra in extra_needed.items():
        last_row = None
        for i, v in enumerate(existing_values):
            cell_val = v[0] if v else ""
            if parse_date_cell(cell_val) == ymd:
                last_row = i + 1  # 1-based
        if last_row is None:
            print(f"  Warning: date {ymd} not found in sheet, skip")
            continue
        original_insert_at = last_row + 1  # insert after this row
        insert_ops.append((original_insert_at, num_extra, ymd))
        y, m, d = ymd
        print(f"  Plan: insert {num_extra} row(s) after row {last_row} ({y}-{m}-{d}), original insert position = {original_insert_at}")

    # Step 4: Perform row inserts (BOTTOM to TOP so positions don't shift for upper inserts)
    insert_ops.sort(key=lambda x: x[0], reverse=True)
    total_inserts = len(insert_ops)

    insert_ops_with_actual_pos = []
    for i, (orig_at, num_rows, ymd) in enumerate(insert_ops):
        pushes_from_above = total_inserts - i - 1
        actual_start_row = orig_at + pushes_from_above
        insert_ops_with_actual_pos.append((actual_start_row, num_rows, ymd))
        y, m, d = ymd
        print(f"  Insert {i}: original={orig_at}, pushed by={pushes_from_above}, actual_start={actual_start_row} ({y}-{m}-{d})")

    if insert_ops:
        insert_requests = []
        for orig_at, num_rows, ymd in insert_ops:
            insert_requests.append({
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": orig_at - 1,
                        "endIndex": orig_at - 1 + num_rows
                    },
                    "inheritFromBefore": False
                }
            })

        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": insert_requests}
        ).execute()
        print(f"Done: inserted {len(insert_requests)} row groups")

    # Step 5: Write dates to newly inserted rows
    date_updates = []
    for actual_start_row, num_rows, ymd in insert_ops_with_actual_pos:
        y, m, d = ymd
        date_str = f"{y}/{m}/{d}"
        for i in range(num_rows):
            row_num = actual_start_row + i
            date_updates.append({
                "range": f"'{sheet_name}'!A{row_num}",
                "values": [[date_str]]
            })
            print(f"  Writing date {date_str} to row {row_num}")

    if date_updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": date_updates}
        ).execute()
        print(f"Done: wrote dates to {len(date_updates)} new rows")

    # Step 6: Re-read A:G after inserts
    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:G").execute()
    existing_values = dv.get("values", [])
    print(f"Existing rows after inserts: {len(existing_values)}")

    # Step 7: Match each record to a row using G column line value
    value_updates = []
    matched_rows = set()

    for rec in records:
        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]
        target_ymd = (y, m, d)
        rec_line = rec.get("line", "11")  # default 11 for old records without line field

        target_row = None
        for i, row in enumerate(existing_values):
            row_num = i + 1
            if row_num in matched_rows:
                continue
            cell_val = row[0] if len(row) > 0 else ""
            if parse_date_cell(cell_val) != target_ymd:
                continue
            g_val = row[6] if len(row) > 6 else ""
            row_line = parse_line(g_val)
            # If row has no G value yet (new empty row), match by order
            if row_line == "" or row_line is None:
                # fallback: if this row is unused and date matches, take it
                target_row = row_num
                break
            if row_line == rec_line:
                target_row = row_num
                break

        if target_row is None:
            # Fallback: any unused row with matching date
            for i, row in enumerate(existing_values):
                row_num = i + 1
                if row_num in matched_rows:
                    continue
                cell_val = row[0] if len(row) > 0 else ""
                if parse_date_cell(cell_val) == target_ymd:
                    target_row = row_num
                    break

        if target_row is None:
            print(f"  {date_str} ({rec_line}): row not found, skip")
            continue

        matched_rows.add(target_row)

        for col, val in rec.items():
            if col in ("date", "line") or val is None:
                continue
            value_updates.append({
                "range": f"'{sheet_name}'!{col}{target_row}",
                "values": [[val]]
            })
        print(f"  {date_str} ({rec_line}) -> row {target_row}")

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
