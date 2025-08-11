import fitz  # PyMuPDF
import re
from datetime import datetime, timedelta
import unicodedata

def extract_text(PDF_path):
    doc = fitz.open(PDF_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

import fitz  # PyMuPDF

def find_word_positions(PDF_path, keyword,search_height=200):
    """
    PDF内の指定文字の座標(x0, y0, x1, y1)をすべて返す。
    戻り値はページごとのリスト: [(page_num, x0, y0, x1, y1), ...]
    """
    # PDFを開いてテキスト化する
    doc = fitz.open(PDF_path)
    positions = []

    for page_num, page in enumerate(doc, start=1):
        words = page.get_text("words")
        for x0, y0, x1, y1, text, *_ in words:
            if y1 <= search_height and text == keyword:   #keywordが完全一致
                positions.append((page_num, x0, y0, x1, y1))
                
    print(f"[DEBUG] '{keyword}' の位置: {positions}")  # デバッグ用
    return positions


def extract_text_in_xrange(PDF_path,  x_min, x_max,page_num=1,):
    """
    指定ページの指定x範囲にあるテキストを、上から順に返す。
    """
    doc = fitz.open(PDF_path)
    page = doc[page_num - 1]

    words = page.get_text("words")
    # X座標範囲でフィルタ
    filtered = [
        (y0, text) for x0, y0, x1, y1, text, *_ in words
        if x0 >= x_min and x1 <= x_max
    ]
    # Y座標順に並べてテキストだけ返す
    filtered.sort(key=lambda w: w[0])
    print(f"[DEBUG] keywordと同列のテキスト: {filtered}")  # デバッグ用
    return [{"text": text, "y": y} for y, text in filtered]


# 指定キーワードの列を抽出し、個々のy座標を取得
def extract_column_and_yrange_from_PDF_B(PDF_path,keyword,range=30):
    
    find_column= find_word_positions(PDF_path,keyword, search_height=200)
    print(f"[DEBUG] {keyword} 文字の座標: {find_column}")  # デバッグ用
    x = find_column[0][1]  # keywordのx座標を取得
    x_min=x- range
    x_max=x+ range
    target_column=extract_text_in_xrange(PDF_path, x_min, x_max, page_num=1)
    target_column = [item for item in target_column if item['text'] != keyword]
    print(f"[DEBUG] {keyword} 列のテキスト: {target_column}")  # デバッグ用
    return target_column

#HD早出勤務表の抽出用関数
def extract_HD_schedule_from_PDF_B(PDF_path,year, selected_name,y_tolerance=5):

    date_column = extract_column_and_yrange_from_PDF_B(PDF_path, "日付", range=40)
    name_column = extract_column_and_yrange_from_PDF_B(PDF_path, "早出")

    print(f"[DEBUG] 日付列: {date_column}")  # デバッグ用
    print(f"[DEBUG] 早出列: {name_column}")  # デバッグ用
    


    merged = []
    
    for d in date_column:
        d_y = d["y"]
        # y座標が近い name_column 要素を探す（差が最小のもの）
        candidates = [(abs(n["y"] - d_y), n) for n in name_column]
        candidates = [c for c in candidates if c[0] <= y_tolerance]

        if candidates:
            # 差が最小のものを選択
            candidates.sort(key=lambda x: x[0])
            closest = candidates[0][1]
            merged.append({"year":year,
                           "date_text": d["text"], 
                           "name_text": closest["text"]})
        else:
            # 見つからなければ None
            merged.append({"year":year,
                           "date_text": d["text"], 
                           "name_text": None})
    

           
    convert_for_google = []
    timezone = "Asia/Tokyo"  # 日本時間

    def format_date(date_text):
        m = re.match(r"(\d{1,2})月(\d{1,2})日", date_text)
        if m:
            month = int(m.group(1))
            day = int(m.group(2))
            return f"{month:02d}-{day:02d}"
        else:
            raise ValueError(f"日付形式が不正です: {date_text}")
    
    for m in merged:
        if m["name_text"]:
            start = f"{m['year']}-{format_date(m['date_text'])}T07:30:00"
            end = f"{m['year']}-{format_date(m['date_text'])}T16:15:00"
            convert_for_google.append({
                "start": {
                    "dateTime": start,
                    "timeZone": timezone
                },
                "end": {
                    "dateTime": end,
                    "timeZone": timezone
                },
                "summary": "HD早出",
                "description": f"勤務表:HD 職員:{m['name_text']}"
            })
    # selected_nameの苗字だけ抽出
    last_name = selected_name.split()[0]  # スペースで分割して先頭を取得
    HD_schedule = [n for n in convert_for_google if last_name in n.get("description", "")]
    if not HD_schedule:
        print(f"[DEBUG] {selected_name} のHD早出勤務は無し")
        HD_schedule = []
    else:
        print(f"[DEBUG] {selected_name} のHD早出勤務イベント数: {len(HD_schedule)}")

    for a in HD_schedule[:3]:
        print("・", a)

    return HD_schedule

#file_PDF_Bから名前だけ取り出す関数
def extract_names_from_PDF_B(PDF_path): 

    name_column = extract_column_and_yrange_from_PDF_B(PDF_path, "早出")
    name_list=[]
    seen = set()
    for n in name_column:
        text = n.get("text")
        if text is not None and text not in seen:
            name_list.append(text)
            seen.add(text)

    for a in name_list:
        print("・", a)
    return name_list

def extract_month_from_PDF_B(PDF_path):
    date_column = extract_column_and_yrange_from_PDF_B(PDF_path, "日付", range=40)
    date = date_column[0].get("text") 
    month=date[0] #dateの1文字目を取得
    print(f"month:{month}")
    return month
#test用
if __name__ == "__main__":
    import os

    # テスト対象のファイルパス
    upload_dir = "uploads"
    #test_filename = "勤務表2025.8ver4.pdf"
    test_filename = "血液浄化センター　早出勤務表　2025年 8月.pdf"
    test_path = os.path.join(upload_dir, test_filename)
    # 年の抽出（ファイル名から）
    match = re.search(r"(20\d{2})", test_filename)
    if match:
        year= int(match.group(1))
    else:
        year= None

    test_name="町田 つばさ"
    print("📄 [TEST] ファイル:", test_path)
    #extract_HD_schedule_from_PDF_B(test_path,year,test_name)
    extract_names_from_PDF_B(test_path)
    #extract_month_from_PDF_B(test_path)
    
    
    try:
        
        """         
        # 年月抽出テスト
        dates= extract_date_from_pdf(test_path)
        print("✅ 抽出された職員名一覧:")
        for date in dates:
            print("・", date)
        # 職員名一覧抽出テスト
        names = extract_names_from_pdf(test_path)
        print("✅ 抽出された職員名一覧:")
        for name in names:
            print("・", name) 
        """

        # 勤務予定抽出テスト（最初の職員で）
        """  
        if names:
            test_name = names[1]
            print(f"\n📆 {test_name} の勤務予定を抽出中...")
            work_days = extract_schedule_from_pdf(test_path, test_name)
        else:
            print("⚠ 職員名が1人も見つかりませんでした。") 
        """

    except Exception as e:
        print("❌ エラー:", e)