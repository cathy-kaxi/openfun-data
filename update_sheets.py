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

    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:G").execute()
    existing_values = dv.get("values", [])
    print(f"Existing rows in A:G: {len(existing_values)}")

    from collections import defaultdict
    date_record_count = defaultdict(int)
    for rec in records:
        date_record_count[rec["date"]] += 1

    date_existing_count = defaultdict(int)
    for v in existing_values:
        ymd = parse_date_cell(v[0] if v else "")
        if ymd:
            date_existing_count[ymd] += 1

    extra_needed = {}
    for date_str, count in date_record_count.items():
        y, m, d = [int(x) for x in date_str.split("-")]
        ymd = (y, m, d)
        existing_count = date_existing_count.get(ymd, 0)
        if count > existing_count:
            extra_needed[ymd] = count - existing_count

    insert_ops = []
    for ymd, num_extra in extra_needed.items():
        last_row = None
        for i, v in enumerate(existing_values):
            cell_val = v[0] if v else ""
            if parse_date_cell(cell_val) == ymd:
                last_row = i + 1
        if last_row is None:
            continue
        original_insert_at = last_row + 1
        insert_ops.append((original_insert_at, num_extra, ymd))

    insert_ops.sort(key=lambda x: x[0], reverse=True)
    total_inserts = len(insert_ops)

    insert_ops_with_actual_pos = []
    for i, (orig_at, num_rows, ymd) in enumerate(insert_ops):
        pushes_from_above = total_inserts - i - 1
        actual_start_row = orig_at + pushes_from_above
        insert_ops_with_actual_pos.append((actual_start_row, num_rows, ymd))

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

    if date_updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": date_updates}
        ).execute()

    dv = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A:G").execute()
    existing_values = dv.get("values", [])

    value_updates = []
    matched_rows = set()

    for rec in records:
        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]
        target_ymd = (y, m, d)
        rec_line = rec.get("line", "11")

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
            if row_line == "" or row_line is None:
                target_row = row_num
                break
            if row_line == rec_line:
                target_row = row_num
                break

        if target_row is None:
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

    # Fix Q column: 注册安装转化率 = O/L, write formula for all data rows
    formula_updates = []
    for row_num in range(3, len(existing_values) + 1):
        # Skip empty rows
        if row_num - 1 >= len(existing_values):
            break
        cell_val = existing_values[row_num - 1][0] if existing_values[row_num - 1] else ""
        if parse_date_cell(cell_val) is None:
            continue
        formula_updates.append({
            "range": f"'{sheet_name}'!Q{row_num}",
            "values": [[f"=IFERROR(O{row_num}/L{row_num},0)"]]
        })
    print(f"Writing Q column formulas for {len(formula_updates)} rows")

    if formula_updates:
        value_updates.extend(formula_updates)

    if value_updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": value_updates}
        ).execute()
        print(f"Done: {len(value_updates)} cells updated")
    else:
        print("No values to update")

if __name__ == "__main__":
    main()
