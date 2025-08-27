import fitz  # PyMuPDF
import re
from datetime import datetime, timedelta,date
import unicodedata
from collections import defaultdict
def extract_text(PDF_path):
    doc = fitz.open(PDF_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text
# PDFから1文字ずつのテキストと座標を抽出する関数
def extract_chars(PDF_path, page_mun=1):
    doc = fitz.open(PDF_path)
    page = doc[page_mun-1]
    rawdict = page.get_text("rawdict")

    chars_list = []

    for block in rawdict["blocks"]:
        if "lines" in block:  # テキストブロックのみ対象
            for line in block["lines"]:
                for span in line["spans"]:
                    for c in span["chars"]:  # ここで1文字ずつ
                        x0, y0, x1, y1 = c["bbox"]
                        chars_list.append({
                            "text": c["c"],
                            "area": (x0, y0, x1, y1)
                        })

    
    print(f"chars_list")
    for a in chars_list[:5]:
        print(f"・{a}")
    
    return chars_list

def extract_names_from_PDF_A(PDF_path):
    text = extract_text(PDF_path)

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
        full_name = f"{last}\u3000{first}"
        # 役職などの前置き除外
        if re.match(r"^(主|副|助|代\d?|振\d?)", full_name):
            continue 
        full_names.append(full_name)
    full_names=sorted(full_names)

    
    for a in full_names:
        print("☆", a)
    return full_names

def extract_text_top_area(PDF_path, height_ratio=0.1):
    """
    PDFの1ページ目の上部（ページ高さのheight_ratio分）だけテキストを抽出する。
    """
    doc = fitz.open(PDF_path)
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
def get_schedule_month_from_PDF_A(PDF_path):
    text = extract_text_top_area(PDF_path, height_ratio=0.15)  # 上部15%を抽出

    # 例：「2025 年 8 月」や「2025年8月」の形式を想定
    match = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        print(f"抽出された年月: {year}年{month}月")
        return year, month
    raise ValueError("年月情報が見つかりませんでした。")


def find_word_positions(PDF_path, keyword,search_height=200):
    """
    PDF内の指定文字の座標(x0, y0, x1, y1)をすべて返す。
    戻り値はページごとのリスト: [(page_num, x0, y0, x1, y1), ...]
    """
    # PDFを開いてテキスト化する
    doc = fitz.open(PDF_path)
    positions = []
    name_pattern=re.compile(r"([\u4E00-\u9FFF]{1,5})[ 　]{1}([\u4E00-\u9FFF\u3040-\u309Fー]{1,5})")
    m=name_pattern.match(keyword)


    for page_num, page in enumerate(doc, start=1):
        words = page.get_text("words")
        for x0, y0, x1, y1, text, *_ in words:
            if m: #keywordが名前の時
                last_name = m.group(1)  # 苗字
                first_name = m.group(2)  # 名前

                if y1 <= search_height and text == keyword:   #keywordが完全一致
                   positions.append((page_num, x0, y0, x1, y1))
                else:#keywordと一致無し
                    if y1 <= search_height and text == last_name: #苗字で検索してHit
                         
                         positions.append((page_num, x0, y0, x1, y1))
                    else:
                        if y1 <= search_height and text == first_name: #名前で検索
                            positions.append((page_num, x0, y0, x1, y1))

             
                     
                
            else: #keywordが名前ではないとき
               if y1 <= search_height and text == keyword:   #keywordが完全一致
                    positions.append((page_num, x0, y0, x1, y1))  
               
    print(f"[DEBUG] '{keyword}' の位置: {positions}")  # デバッグ用
    return words,positions


def search_keyword_in_pdf(pdf_path, keyword, search_height=200):
    doc = fitz.open(pdf_path)
    positions = []

    # 名前パターン判定（苗字1〜5漢字 + スペース + 名前1〜5）
    pattern_name = re.compile(r"^([\u4E00-\u9FFF]{1,5})[ 　]{1}([\u4E00-\u9FFF\u3040-\u309Fー]{1,5})$")
    m = pattern_name.match(keyword)

    if m:
        last_name, first_name = m.groups()
    else:
        last_name = first_name = None

    for page_num, page in enumerate(doc, start=1):
        words = page.get_text("words")  # [(x0, y0, x1, y1, text, ...), ...]

        if m:  # 名前パターン
            full_hits = []
            last_hits = []
            first_hits = []

            for x0, y0, x1, y1, text, *_ in words:
                if y1 > search_height:
                    continue
                if text == keyword:#フルネームヒット時はfull_hitsに格納
                    full_hits.append((page_num, x0, y0, x1, y1))
                elif text == last_name:#苗字ヒット時はlast_hitsに格納
                    last_hits.append((page_num, x0, y0, x1, y1))
                elif text == first_name:#名前ヒット時はfirst_hitsに格納
                    first_hits.append((page_num, x0, y0, x1, y1))

            if len(full_hits)==1:  # フルネームでヒット
                positions.extend(full_hits)
            elif len(full_hits)>1:
                 print(f"⚠ 同姓同名: {keyword}")
            elif len(last_hits) == 1:  # 苗字1件のみ
                positions.extend(last_hits)
            elif len(last_hits) > 1:  # 苗字2件以上
                if len(first_hits) == 1:
                    positions.extend(first_hits)
                elif len(first_hits) > 1:
                    print(f"⚠ 同姓同名: {keyword}")
                    positions.extend(first_hits)  # 必要なら全件追加

        else:  # 名前パターンじゃない
            for x0, y0, x1, y1, text, *_ in words:
                if y1 > search_height:
                    continue
                if text == keyword:
                    positions.append((page_num, x0, y0, x1, y1))
    print(f"{keyword}の座標は{positions}")
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
    #print(f"[DEBUG] keywordと同列のテキスト: {filtered}")  # デバッグ用
    return [{"text": text, "y": y} for y, text in filtered]

def pick_up_date_line(PDF_path,sub=10, add=-5):
    words,position=find_word_positions(PDF_path, "名前",search_height=200)
    # position は [(page_num, x0, y0, x1, y1)] 形式を想定
    y0 = position[0][2]
    y1 = position[0][4]
    print(f"y0:{y0}, y1:{y1}")
    y_min=y0-sub
    y_max=y1+add
    date_line=[]
    seen_texts = set()  # 出現済みの text を記録する集合
    for word in words:
        x0, y0, x1, y1, text,*_=word
        if y0>=y_min and y1<=y_max:
            if text not in seen_texts:  # 初めて出た text だけ追加
                date_line.append({"text": text, "area": (x0, y0, x1, y1)})
                seen_texts.add(text)    
    # print(f"date_line")
    # for a in date_line:
    #     print(f"・{a}")
    return date_line

def pick_up_row_text(PDF_path,keyword,page_num=1,sub=5,add=-3,search_height=800):
    all_text=extract_chars(PDF_path, page_mun=page_num)
    words,keyword_position=find_word_positions(PDF_path, keyword,search_height=search_height)
    print(f"keyword_position:{keyword_position}")
    y0=keyword_position[0][2]
    y1=keyword_position[0][4]
    print(f"y0:{y0}, y1:{y1}")
    y_min=y0-sub
    y_max=y1+add

    target_row=[]
    for text in all_text:
        if text["area"][1]>=y_min and text["area"][3]<=y_max:
            #print(f"・{text}")
            target_row.append({"text":text["text"],"area":text["area"]})
    return target_row

def merge_target_row_dataline(PDF_path, keyword):
    target_row=pick_up_row_text(PDF_path, keyword,page_num=1, sub=5,add=-3,search_height=800)
    date_line = pick_up_date_line(PDF_path,sub=10, add=-5)
    
    """     print("date_line")
    for a in date_line:
        print(f"・{a}")
    print("target_row")
    for a in target_row:
        print(f"・{a}")
     """
    merged=[]
    for d in date_line:
        d_x0 = d["area"][0]-10
        d_x1 = d["area"][2]+10
        for t in target_row:
            t_x0 = t["area"][0]
            t_x1 = t["area"][2]
            # X座標が近い target_row 要素を探す（差が最小のもの）
            if t_x0>d_x0 and t_x1<d_x1:
                new_t = t.copy()          # 辞書をコピー
                new_t["date"] = d["text"] # date を追加
                merged.append(new_t)
    # print("merged")
    # for a in merged:
    #     print(f"・{a}")

    grouped = defaultdict(list)
    # 日付ごとに grouped に格納
    for m in merged:   
        grouped[m["date"]].append(m)

    merged_date = []

    # grouped を処理
    for date, items in grouped.items():
        # x 座標順に並べる
        items = sorted(items, key=lambda it: it["area"][0])

        # テキストを連結
        merged_text = "".join([it["text"] for it in items])

        # 最小・最大座標を取って囲む
        x0 = min(it["area"][0] for it in items)
        y0 = min(it["area"][1] for it in items)
        x1 = max(it["area"][2] for it in items)
        y1 = max(it["area"][3] for it in items)

        # merged_date に追加
        merged_date.append({
            "date": date,
            "text": merged_text,
            "area": (x0, y0, x1, y1),
            
        })


    print("merged_date")
    for a in merged_date:
        print(f"・{a}")


    return merged_date    

def extract_schedule_from(PDF_path, keyword):
    merged_date=merge_target_row_dataline(PDF_path, keyword)
    year_A,month_A=get_schedule_month_from_PDF_A(PDF_path)
    convert_for_google = []
    timezone = "Asia/Tokyo"  # 日本時間

    for m in merged_date:
        day = int(m["date"])
        start=date(year_A, month_A, day).strftime("%Y-%m-%d")
        end=(date(year_A, month_A, day)+ timedelta(days=1)).strftime("%Y-%m-%d")
        if m["text"]=="代休":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "代替休日",
                                    "description": f"勤務表:MAIN 職員:{keyword}"})
        elif m["text"]=="年休":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "年次休暇",
                                    "description": f"勤務表:MAIN 職員:{keyword}"})
        elif m["text"]=="振休":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "振替休日",
                                    "description": f"勤務表:MAIN 職員:{keyword}"})
        elif m["text"]=="×":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "業務対応不可",
                                    "description": f"勤務表:MAIN 職員:{keyword}"})
        elif m["text"]=="⑯":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "当直",
                                    "description": f"勤務表:MAIN 職員:{keyword}"})
        elif m["text"]=="1":
            convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                "end": {"date": end,"timeZone": timezone},
                                "summary": "1st on call",
                                "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["text"]=="2":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "2nd on call",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
      
        elif m["text"]=="ＡＭ休":
                end = f"{start}T17:15:00"
                start = f"{start}T13:22:30"
                
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "午後出勤（午前休）",
                                    "description": f"勤務表:MAIN 職員:{keyword}"})

        elif m["text"]=="ＰＭ休":
                end = f"{start}T12:22:30"
                start = f"{start}T08:30:00"
                
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "午前出勤（午後休）",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["text"]=="明":
                end = f"{start}T08:30:00"
                start = f"{start}T00:00:00"
                
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "当直明け",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
                
        else:
            #何も追加しないようにする
             pass
    print("convert_for_google")
    for a in convert_for_google:
        print(f"・{a}")
    return convert_for_google


# 指定キーワードの列を抽出し、個々のy座標を取得
def extract_column_and_yrange_from_PDF_A(PDF_path,keyword,search_height=200,sub=10,add=10):
    
    find_column= search_keyword_in_pdf(PDF_path,keyword, search_height=search_height)
    print(f"[DEBUG] {keyword} 文字の座標: {find_column}")  # デバッグ用
    x = find_column[0][1]  # keywordのx座標を取得
    x_min=x- sub
    x_max=x+ add
    target_column=extract_text_in_xrange(PDF_path, x_min, x_max, page_num=1)
    target_column = [item for item in target_column if item['text'] != keyword]
    print(f"[DEBUG] {keyword} 列のテキスト: {target_column}")  # デバッグ用
    return target_column

# 指定キーワードの行を抽出し、個々のx座標を取得
def extract_row_and_xrange_from_PDF_A(PDF_path,keyword,search_height=300,sub=10,add=10):
    find_row= search_keyword_in_pdf(PDF_path,keyword, search_height=search_height)
    print(f"[DEBUG] {keyword} 文字の座標: {find_row}")  # デバッグ用
    y = find_row[0][2]  # keywordのy座標を取得
    #取得したい情報が取得できるy範囲を微調整
    y_min=y-sub
    y_max=y+add
    target_row=extract_text_in_yrange(PDF_path, y_min, y_max, page_num=1)
    target_row = [item for item in target_row if item['text'] != keyword]
    print(f"[DEBUG] {keyword} 行のテキスト: {target_row}")  # デバッグ用
    return target_row



#勤務表の抽出用関数
def extract_schedule_from_PDF_A(PDF_path, selected_name,x_tolerance=5):

    date_line = extract_row_and_xrange_from_PDF_A(PDF_path,"名前",search_height=200,sub=20,add=10)
    target_line = extract_row_and_xrange_from_PDF_A(PDF_path,selected_name,search_height=800,sub=10,add=10)
    

    for i, cell in enumerate(target_line):
        if re.search(r"(日|夜|日夜|勤務|勤)", cell["text"]):
            target_line = target_line[i + 1:]
            print(f"[DEBUG] Found work start position: \n{target_line}")  # デバッグ用
            break
    else:
        raise ValueError("勤務マーク開始位置（日夜など）が見つかりませんでした。")
    
    year_A,month_A=get_schedule_month_from_PDF_A(PDF_path)

    
    
    
    print(f"[DEBUG] 日付列: {date_line}")  # デバッグ用
    print(f"[DEBUG] {selected_name}列: {target_line}")  # デバッグ用
    

    #date_lineとtarget_lineを統合する
    merged = []
    
    for d in date_line:
        d_x = d["x"]
        # x座標が近い target_line 要素を探す（差が最小のもの）
        candidates = [(abs(n["x"] - d_x), n) for n in target_line]
        candidates = [c for c in candidates if c[0] <= x_tolerance]

        if candidates:
            # 差が最小のものを選択
            candidates.sort(key=lambda x: x[0])
            closest = candidates[0][1]
            merged.append({"date_text": d["text"], 
                           "work_mark": closest["text"]})
        else:
            # 見つからなければ None
            merged.append({"date_text": d["text"], 
                           "work_mark": None})
    print(f"merged:\n{merged}")
    
    #merged["work_mark"]の種類ごとにリスト内を書き換え且つgoogle calender用にかえる
    convert_for_google = []
    timezone = "Asia/Tokyo"  # 日本時間
    for m in merged:
        day = int(m["date_text"])
        start=date(year_A, month_A, day).strftime("%Y-%m-%d")
        end=(date(year_A, month_A, day)+ timedelta(days=1)).strftime("%Y-%m-%d")
        if m["work_mark"]=="1":
            convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                "end": {"date": end,"timeZone": timezone},
                                "summary": "1st on call",
                                "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="2":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "2nd on call",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="代休":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "代替休日",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="年休":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "年次休暇",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="振休":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "振替休日",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="×":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "業務対応不可",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="⑯":
                convert_for_google.append({"start": {"date": start,"timeZone": timezone},
                                    "end": {"date": end,"timeZone": timezone},
                                    "summary": "当直",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="ＡＭ休":
                end = f"{start}T17:15:00"
                start = f"{start}T13:22:30"
                
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "午後出勤（午前休）",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="ＰＭ休":
                end = f"{start}T12:22:30"
                start = f"{start}T08:30:00"
                
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "午前出勤（午後休）",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
        elif m["work_mark"]=="明":
                end = f"{start}T08:30:00"
                start = f"{start}T00:00:00"
                
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "当直明け",
                                    "description": f"勤務表:MAIN 職員:{selected_name}"})
                
        else:
            #何も追加しないようにする
             pass




    # selected_nameの苗字だけ抽出
    if not convert_for_google:
        print(f"[DEBUG] {selected_name} の勤務予定は無し")
        
    else:
        print(f"[DEBUG] {selected_name} のHD早出勤務イベント数: {len(convert_for_google)}")

    for a in convert_for_google[:5]:
        print("☆", a)

    return convert_for_google

#file_PDF_Bから名前だけ取り出す関数
#test用
if __name__ == "__main__":
    import os

    # テスト対象のファイルパス
    upload_dir = "uploads"
    test_filename = "勤務表2025.9ver3.1.pdf"
    #test_filename = "血液浄化センター　早出勤務表　2025年 8月.pdf"
    test_path = os.path.join(upload_dir, test_filename)
    search_height=800


    test_name="戸田　修一"
    print("📄 [TEST] ファイル:", test_path)


    print(fitz.__doc__)
    print(fitz.__version__)
    
    #extract_chars(test_path)
    #pick_up_date_line(test_path,sub=10, add=-5)
    #extract_names_from_PDF_A(test_path)
    #search_keyword_in_pdf(test_path, test_name, search_height)
    #find_word_positions(test_path,test_name,search_height=800)
    #pick_up_row_text(test_path,test_name,page_num=1,sub=5,add=-3,search_height=800)
    merge_target_row_dataline(test_path,test_name)
    extract_schedule_from(test_path, test_name)
    #extract_column_and_yrange_from_PDF_A(test_path,"名前",sub=40,add=30)
    #extract_row_and_xrange_from_PDF_A(test_path,"名前",search_height=200,sub=20,add=10)
    #extract_row_and_xrange_from_PDF_A(test_path,"大江　直義",search_height=500,sub=10,add=10)
    #extract_schedule_from_PDF_A(test_path,test_name,x_tolerance=6)
    #extract_names_from_PDF_A(test_path)

    #extract_month_from_PDF_A(test_path)
    
    
