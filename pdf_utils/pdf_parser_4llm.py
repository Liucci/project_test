#pymupdf4llmは制度は高いが、速度は遅いので注意

import pymupdf4llm
import re
from datetime import datetime, timedelta

def pdf_to_markdown(pdf_path):
    md_chunks = pymupdf4llm.to_markdown(pdf_path, page_chunks=True)
    return "\n".join(chunk["text"] for chunk in md_chunks)



def extract_schedule_from_markdown(pdf_path, staff_name):
    md_text=pdf_to_markdown(pdf_path)
    lines = md_text.splitlines()
    
    # 年月抽出（例："2025年8月"）
    year = month = None
    for line in lines:
        m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", line)
        
        if m:
            year, month = int(m.group(1)), int(m.group(2))
            print(f"[DEBUG] Found year: {year}, month: {month}")  # デバッグ用
            break
    if not year:
        raise ValueError("年月情報が見つかりませんでした。")
    
    # 日付行の探索
    date_row = None
    for line in lines:
        nums = re.findall(r"\b\d{1,2}\b", line)
        if len(nums) >= 20 and all(1 <= int(n) <= 31 for n in nums):
            date_row = nums
            print(f"[DEBUG] Found date row: \n{date_row}")  # デバッグ用
            break
    if not date_row:
        raise ValueError("日付の行が見つかりませんでした。")
    
    # 対象職員の行を抽出
    
    staff_name_norm = staff_name.replace("　", " ").replace("\u3000", " ").strip()

    target_line = next((line for line in lines
                        if staff_name_norm in line.replace("　", " ").replace("\u3000", " ").replace("\u200b", "")
                        ),
                        None)  
    print(f"[DEBUG] Target line for {staff_name}: {target_line}")  # デバッグ用
    target_line = [re.sub(r"<[^>]+>|~+", "", c).strip() for c in target_line.split("|")]
    
    print(f"[DEBUG] Target line for {staff_name}: \n{target_line}")  # デバッグ用
    
    for i, cell in enumerate(target_line):
        if re.search(r"(日|夜|日夜)", cell):
            target_line = target_line[i + 1:]
            print(f"[DEBUG] Found work start position: \n{target_line}")  # デバッグ用
            break
    else:
        raise ValueError("勤務マーク開始位置（日夜など）が見つかりませんでした。")

    if not target_line:
        raise ValueError(f"{staff_name} の勤務行が見つかりませんでした。")
    

    #date_rowとtarget_lineをdictにする
    
    work_day_dict = {
                    day: mark
                    for day, mark in zip(date_row, target_line)
                    if day is not None
                    }
    print(f"[DEBUG] Work day dictionary: \n{list(work_day_dict.items())[:5]}")

    
    # Step 1: 最初に year, month, day を持つ dict に変換
    for day_str in list(work_day_dict.keys()):
        raw_summary = work_day_dict[day_str]
        if not day_str.strip().isdigit():
            continue  # 無効な日付はスキップ
        day = int(day_str.strip())
        work_day_dict[day_str] = {
                            "year": year,
                            "month": month,
                            "day": day,
                            "summary": raw_summary  # まだ翻訳前
                        }
    print(f"[DEBUG] work day dictionary　add year and month: \n{list(work_day_dict.items())[:5]}")  # デバッグ用                    
    # 勤務略称を書き換え
    work_translation = {
                            "1": ("1st on call", None),                    
                            "2": ("2nd on call", None),
                            "代休": ("代替休日", None),
                            "年休": ("年次休暇", None),
                            "×": ("業務対応不可", None),
                            "AM休": ("午後勤務（午前休）", ("13:22:30", "17:15:00")),  
                            "PM休": ("午前勤務（午後休）", ("09:00:00", "12:00:00")),
                            "明": ("明け", ("00:00:00", "08:30:00")),
                            "出": ("出勤", None),
                            "振休": ("振替休日", None),
                            "⑯": ("当直", None),
                            # 必要に応じて追加
                                }
    
    for day_str, info in work_day_dict.items():
        raw = info.get("summary", "").strip()
        year = info["year"]
        month = info["month"]
        day = info["day"]

        date_obj = datetime(year, month, day)

        # 変換辞書から取得
        if raw in work_translation:
            summary_text, time_range = work_translation[raw]
        else:
            summary_text = raw  # 未定義ならそのまま
            time_range = None

        # 終日イベント（start/endは日付のみ）
        if time_range is None:
            start = date_obj.date().isoformat()
            end = (date_obj + timedelta(days=1)).date().isoformat()
            info["start"] = {"date": start}
            info["end"] = {"date": end}
    

        # 時間指定イベント（start/endは日時）
        else:
            start = datetime.combine(date_obj.date(), datetime.strptime(time_range[0], "%H:%M:%S").time()).isoformat()
            end = datetime.combine(date_obj.date(), datetime.strptime(time_range[1], "%H:%M:%S").time()).isoformat()
    # summaryが空欄ならstart/endは空文字
        if summary_text == '':
            start = ''
            end = ''

        # descriptionの設定
        description = f"[勤務表:MAIN] [職員名:{staff_name}] "

        # 上書き
        info["summary"] = summary_text
        info["start"] = start
        info["end"] = end
        info["description"] = description
    
    print(f"[DEBUG] add start end:\n {list(work_day_dict.items())[:5]}")  # デバッグ用
    print(f"work_day_dict:{type(work_day_dict)}")#辞書型
    # 空のsummaryを持つエントリを削除
    work_day_dict = {
                        k: v for k, v in work_day_dict.items()
                        if v.get("summary", "").strip() != ""}
    #start:dict, end:dict, summary:str, description:strに変換
    converted_events = []
    for day_str, info in work_day_dict.items():
        start = info["start"]
        end= info["end"]
        if "T" in start and "T" in end:
            converted_events.append({
            "start": {"dateTime": start, "timeZone": "Asia/Tokyo"},
            "end": {"dateTime": end, "timeZone": "Asia/Tokyo"},
            "summary": info["summary"],
            "description": info.get("description", "")
        })
        else:
            converted_events.append({
            "start": {"date": start},
            "end": {"date": end},
            "summary": info["summary"],
            "description": info.get("description", "")
            })


   
    
    for i, ev in enumerate(converted_events[:3]):
        print(f"[DEBUG] converted_events[{i}] type: {type(ev)}")
        for key in [ 'start', 'end', 'summary', 'description']:
            print(f"    {key}: {ev.get(key)} (type: {type(ev.get(key))})")
    
    #google calendar API用にそのままわたせる形式
    return converted_events



def extract_names_from_pdf_with_4llm(pdf_path):
    md_text = pdf_to_markdown(pdf_path)
    lines = md_text.splitlines()
    pattern = re.compile(r"([\u4E00-\u9FFF]{1,5})[ 　]+([\u4E00-\u9FFF\u3040-\u309F]{1,5})")
    names = set()
    for line in lines:
        matches = pattern.findall(line)
        for last, first in matches:
            full_name = f"{last} {first}"
            if not re.match(r"^(主|副|助|代\d?|振\d?)", full_name):
                names.add(full_name)
    return sorted(names)

def get_schedule_month_from_pdf_with_4llm(pdf_path):
    md_text = pdf_to_markdown(pdf_path)
    print(f"[DEBUG] Extracted markdown text: {md_text[:100]}...")  # デバッグ用
    lines = md_text.splitlines()
    print(f"[DEBUG] Total lines extracted: {len(lines)}")
    print(f"[DEBUG] First 10 lines: {lines[:10]}")

    for line in lines:
        m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", line)
        if m:
            return int(m.group(1)), int(m.group(2))
    raise ValueError("年月情報が見つかりませんでした。")

#test用
if __name__ == "__main__":
    import os

    # テスト対象のファイルパス
    upload_dir = "uploads"
    test_filename = "勤務表2025.8ver4.pdf"
    test_path = os.path.join(upload_dir, test_filename)

    

    print("📄 [TEST] ファイル:", test_path)

    try:
        # 年月抽出テスト
        year, month = get_schedule_month_from_pdf_with_4llm(test_path)
        print(f"✅ 年月抽出: {year}年 {month}月")

        # 職員名一覧抽出テスト
        names = extract_names_from_pdf_with_4llm(test_path)
        print("✅ 抽出された職員名一覧:")
        for name in names:
            print("・", name)

        # 勤務予定抽出テスト（最初の職員で）
        if names:
            test_name = names[1]
            print(f"\n📆 {test_name} の勤務予定を抽出中...")
            work_days = extract_schedule_from_markdown(test_path, test_name)
        else:
            print("⚠ 職員名が1人も見つかりませんでした。")

    except Exception as e:
        print("❌ エラー:", e)
