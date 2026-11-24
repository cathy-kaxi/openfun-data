import json, os, re, sys, base64
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TARGET_SHEET_ID = 0  # gid=0 => openfun日报

# 当前表结构（2026-11 新版）
# A日期 B投放人 C广告平台 D账户时区 E线路名 F(空) G回流后消耗(公式) HSpends
# I回流消耗 JImpression KCPM LClicks MCPC NCTR点击率 O注册量(面板) PFB注册单价
# Q注册安装转化率(公式) R购物次数 S购物单价 T注册用户 U注册成本 V付费用户 W付费成本
# X首购订单数 Y复购订单数 Z首购金额 AA复购金额 ABSuccess总金额
LINE_COL_IDX = 4  # E列 = 线路名
WRITABLE = {"B", "C", "D", "E", "H", "J", "K", "L", "M", "N", "O", "P",
            "R", "S", "T", "V", "X", "Y", "Z", "AA", "AB"}


def parse_date_cell(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=float(v))
            return (dt.year, dt.month, dt.day)
        except Exception:
            return None
    m = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(v).strip())
    if m:
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def parse_line(v):
    if v is None or v == "":
        return ""
    s = str(v).strip()
    if s == "11":
        return "11"
    if s in ("8", "08"):
        return "08"
    return s


def main():
    raw = os.environ["GOOGLE_CREDENTIALS"]
    try:
        creds_info = json.loads(raw)
    except Exception:
        creds_info = json.loads(base64.b64decode(raw).decode())

    with open("data.json") as f:
        data = json.load(f)
    records = data.get("records", [])
    print(f"Loaded {len(records)} records")

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet = next((s for s in meta["sheets"] if s["properties"]["sheetId"] == TARGET_SHEET_ID), None)
    if not sheet:
        print("ERROR: target sheet id=0 not found")
        sys.exit(1)
    sheet_name = sheet["properties"]["title"]
    sheet_id = sheet["properties"]["sheetId"]
    print(f"Sheet: {sheet_name} (id={sheet_id})")

    def read_all():
        dv = svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A1:AB").execute()
        return dv.get("values", [])

    values = read_all()
    print(f"Existing rows: {len(values)}")

    # ---- 1) 为表里还不存在的日期补行（沿用上一行格式）----
    rec_dates = sorted({rec["date"] for rec in records})
    existing_dates = set()
    last_dated = 2
    for i, r in enumerate(values):
        ymd = parse_date_cell(r[0] if r else "")
        if ymd:
            existing_dates.add(ymd)
            last_dated = i + 1

    created = []
    for date_str in rec_dates:
        y, m, d = [int(x) for x in date_str.split("-")]
        if (y, m, d) in existing_dates:
            continue
        row = last_dated + 1
        # 复制上一行格式（只复制格式，不动内容）
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{
                "copyPaste": {
                    "source": {"sheetId": sheet_id, "startRowIndex": row - 2, "endRowIndex": row - 1,
                               "startColumnIndex": 0, "endColumnIndex": 28},
                    "destination": {"sheetId": sheet_id, "startRowIndex": row - 1, "endRowIndex": row,
                                    "startColumnIndex": 0, "endColumnIndex": 28},
                    "pasteType": "PASTE_FORMAT",
                }
            }]},
        ).execute()
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW",
                  "data": [{"range": f"'{sheet_name}'!A{row}",
                            "values": [[f"{y}/{m}/{d}"]]}]},
        ).execute()
        existing_dates.add((y, m, d))
        last_dated = row
        created.append((row, date_str))
        print(f"  created row {row} for {date_str}")

    if created:
        values = read_all()
        print(f"Rows after append: {len(values)}")

    # ---- 2) 按日期 + 线路匹配行，写入数据 ----
    matched_rows = set()
    value_updates = []
    plan = []

    for rec in records:
        date_str = rec["date"]
        y, m, d = [int(x) for x in date_str.split("-")]
        target_ymd = (y, m, d)
        rec_line = str(rec.get("line", "11"))

        target_row = None
        # 优先：同日期 + 同线路
        for i, row in enumerate(values):
            row_num = i + 1
            if row_num in matched_rows:
                continue
            if parse_date_cell(row[0] if row else "") != target_ymd:
                continue
            g_val = row[LINE_COL_IDX] if len(row) > LINE_COL_IDX else ""
            row_line = parse_line(g_val)
            if row_line == "":
                target_row = row_num
                break
            if row_line == rec_line or (rec_line == "08" and row_line == "08"):
                target_row = row_num
                break
        # 兜底：同日期任意未占用行
        if target_row is None:
            for i, row in enumerate(values):
                row_num = i + 1
                if row_num in matched_rows:
                    continue
                if parse_date_cell(row[0] if row else "") == target_ymd:
                    target_row = row_num
                    break
            if target_row is not None:
                print(f"  WARN {date_str} ({rec_line}): line not matched, using row {target_row}")

        if target_row is None:
            print(f"  {date_str} ({rec_line}): row not found, skip")
            continue

        matched_rows.add(target_row)
        n_cols = 0
        for col, val in rec.items():
            if col in ("date", "line") or val is None or col not in WRITABLE:
                continue
            value_updates.append({"range": f"'{sheet_name}'!{col}{target_row}", "values": [[val]]})
            n_cols += 1
        # Q列：注册安装转化率 = 注册量/点击量
        value_updates.append({"range": f"'{sheet_name}'!Q{target_row}",
                              "values": [[f"=IFERROR(O{target_row}/L{target_row},0)"]]})
        plan.append((target_row, date_str, rec_line, n_cols))
        print(f"  {date_str} ({rec_line}) -> row {target_row}, {n_cols} cols")

    if value_updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": value_updates},
        ).execute()
        print(f"Done: {len(value_updates)} cells updated")
    else:
        print("No values to update")

    # ---- 3) 回读校验 ----
    values2 = read_all()
    for row_num, date_str, rec_line, _ in plan:
        if row_num - 1 < len(values2):
            print(f"VERIFY row {row_num}\t" + json.dumps(values2[row_num - 1][:28], ensure_ascii=False))


if __name__ == "__main__":
    main()
