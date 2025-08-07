# pdf_utils/pdf_parser.py

import fitz  # PyMuPDF
import re
from datetime import datetime, timedelta
import unicodedata


def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_names_from_pdf(pdf_path):
    text = extract_text(pdf_path)

    # 漢字・ひらがな・長音符号（ー）を許容
    # 苗字：漢字1～5文字
    # 名：漢字・ひらがな・長音符号 1～5文字
    # スペースは全角・半角どちらもOK
    pattern = re.compile(
        r"([\u4E00-\u9FFF]{1,5})[ 　]{1}([\u4E00-\u9FFF\u3040-\u309Fー]{1,5})"
    )

    matches = pattern.findall(text)

    full_names = []
    for last, first in matches:
        full_name = f"{last} {first}"
        # 役職などの前置き除外
        #if re.match(r"^(主|副|助|代\d?|振\d?)", full_name):
            #continue 
        full_names.append(full_name)

    return sorted(set(full_names))


def extract_text_top_area(pdf_path, height_ratio=0.1):
    """
    PDFの1ページ目の上部（ページ高さのheight_ratio分）だけテキストを抽出する。
    """
    doc = fitz.open(pdf_path)
    page = doc[0]  # 1ページ目を対象

    rect = page.rect
    top_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * height_ratio)

    # 指定矩形のテキスト抽出（ブロック単位で取得して結合）
    blocks = page.get_text("blocks")  # 各テキストブロック: (x0, y0, x1, y1, "text", block_no, block_type)
    texts = []
    for b in blocks:
        b_rect = fitz.Rect(b[:4])
        if b_rect.intersects(top_rect):
            texts.append(b[4])

    return "\n".join(texts)

def get_schedule_month_from_pdf(pdf_path):
    text = extract_text_top_area(pdf_path, height_ratio=0.15)  # 上部15%を抽出

    # 例：「2025 年 8 月」や「2025年8月」の形式を想定
    match = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        print(f"抽出された年月: {year}年{month}月")
        return year, month
    raise ValueError("年月情報が見つかりませんでした。")



def find_date_row_only(lines):
    for line in lines:
        normalized = re.sub(r"[　]+", " ", line.strip())  # 全角スペース除去
        numbers = re.findall(r"\b\d{1,2}\b", normalized)
        if len(numbers) >= 20 and all(1 <= int(n) <= 31 for n in numbers):
            return numbers
    raise ValueError("日付の行が見つかりませんでした。")



def extract_schedule_from_pdf(pdf_path, staff_name):
    text = extract_text(pdf_path)
    lines = text.splitlines()
    year, month = get_schedule_month_from_pdf(pdf_path)

    # 日付行の取得（曜日行は使わない）
    date_row = find_date_row_only(lines)

    # 職員の行を探す
    target_line = next((line for line in lines if staff_name in line), None)
    if not target_line:
        raise ValueError(f"{staff_name} の勤務行が見つかりませんでした。")

    # 勤務内容を抽出（全角スペース→半角へ変換）
    target_line_cleaned = re.sub(r"[　]+", " ", target_line)
    parts = target_line_cleaned.split()

    work_columns = parts[1:] if staff_name.replace("　", " ") in parts[0] else parts

    events = []
    work_marks = {"年休", "当", "明", "出", "振休", "⑯", "2", "1", "×"}

    for i, mark in enumerate(work_columns):
        if i >= len(date_row):
            break
        if mark in work_marks:
            try:
                day = int(date_row[i])
                start_date = datetime(year, month, day).date()
                end_date = start_date + timedelta(days=1)
                events.append({
                    "summary": f"{staff_name}：{mark}",
                    "start": {"date": start_date.isoformat()},
                    "end": {"date": end_date.isoformat()},
                    "description": f"staff={staff_name}",
                })
            except ValueError:
                continue

    return events, year, month

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
        year, month = get_schedule_month_from_pdf(test_path)
        print(f"✅ 年月抽出: {year}年 {month}月")

        # 職員名一覧抽出テスト
        names = extract_names_from_pdf(test_path)
        print("✅ 抽出された職員名一覧:")
        for name in names:
            print("・", name)

        # 勤務予定抽出テスト（最初の職員で）
        if names:
            test_name = names[1]
            print(f"\n📆 {test_name} の勤務予定を抽出中...")
            work_days = extract_schedule_from_pdf(test_path, test_name)
        else:
            print("⚠ 職員名が1人も見つかりませんでした。")

    except Exception as e:
        print("❌ エラー:", e)