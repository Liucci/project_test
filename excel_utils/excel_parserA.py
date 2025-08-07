import pandas as pd
from datetime import datetime, timedelta
from pytz import timezone

def parse_schedule_from_sheet_a(filepath, selected_name):
    df = pd.read_excel(filepath, sheet_name=0, header=None)
    c1, j1 = df.iloc[0, 2], df.iloc[0, 9]

    year = c1.year if isinstance(c1, datetime) else int(c1)
    month = j1.month if isinstance(j1, datetime) else int(j1)
    print(f"C1セル: {c1}, 型: {type(c1)}")
    print(f"J1セル: {j1}, 型: {type(j1)}")
    name_rows = df.iloc[8:, :]
    target_row = name_rows[name_rows.iloc[:, 2] == selected_name]

    if target_row.empty:
        return [], year, month, f"{selected_name} の勤務情報が見つかりませんでした"

    start_col = 9
    max_days = 31
    work_cells = target_row.iloc[0, start_col:start_col + max_days].tolist()
    days = list(range(1, max_days + 1))
    events = []
    jst = timezone("Asia/Tokyo")

    for i, cell in enumerate(work_cells):
        if pd.isna(cell):
            continue

        day = days[i]
        try:
            date = datetime(year, month, day)
        except ValueError:
            continue

        cell = str(cell).strip()

        if cell == "1":
            summary = "1st on call"
        elif cell == "2":
            summary = "2nd on call"
        elif cell == "⑯":
            summary = "当直"
        elif cell == "年休":
            summary = "有給休暇"
        elif cell in ["振休", "代休"]:
            summary = "振替休日"
        elif cell == "ＡＭ休":
            events.append({
                "summary": "午後勤務（午前休）",
                "start": {"dateTime": date.strftime("%Y-%m-%dT13:22:30"), "timeZone": "Asia/Tokyo"},
                "end": {"dateTime": date.strftime("%Y-%m-%dT17:15:00"), "timeZone": "Asia/Tokyo"},
                "description": f"集中管理部勤務表由来\nstaff={selected_name}"
            })
            continue
        elif cell == "ＰＭ休":
            events.append({
                "summary": "午前勤務（午後休）",
                "start": {"dateTime": date.strftime("%Y-%m-%dT08:30:00"), "timeZone": "Asia/Tokyo"},
                "end": {"dateTime": date.strftime("%Y-%m-%dT12:22:30"), "timeZone": "Asia/Tokyo"},
                "description": f"集中管理部勤務表由来\nstaff={selected_name}"
            })
            continue
        elif cell == "早HD":
            events.append({
                "summary": "HD早番",
                "start": {"dateTime": date.strftime("%Y-%m-%dT07:30:00"), "timeZone": "Asia/Tokyo"},
                "end": {"dateTime": date.strftime("%Y-%m-%dT16:15:00"), "timeZone": "Asia/Tokyo"},
                "description": f"集中管理部勤務表由来\nstaff={selected_name}"
            })
            continue
        else:
            continue

        # 終日イベント
        events.append({
            "summary": summary,
            "start": {"date": date.strftime("%Y-%m-%d")},
            "end": {"date": (date + timedelta(days=1)).strftime("%Y-%m-%d")},
            "description": f"集中管理部勤務表由来\nstaff={selected_name}"
        })
    print(f"[DEBUG] Parsed {len(events)} events from file A for {selected_name}")
    print(f"[DEBUG] 勤務表A抽出結果: {events}")
    return events, year, month, None

def extract_names_from_sheet_a(file_path):
    df = pd.read_excel(file_path, sheet_name=0, header=None)

    # 8行目（インデックス7）以降のデータを対象にする
    data_start_row = 8
    name_column = 2  # 列C：Unnamed: 2

    names = df.iloc[data_start_row:, name_column]
    names = names.dropna().unique().tolist()
    #print(f"抽出された名前: {names}")
    return names


def get_schedule_month_from_sheet_a(filepath):
    df = pd.read_excel(filepath, sheet_name=0, header=None)
    c1, j1 = df.iloc[0, 2], df.iloc[0, 9]

    year = c1.year if isinstance(c1, datetime) else int(c1)
    month = j1.month if isinstance(j1, datetime) else int(j1)

    return year, month