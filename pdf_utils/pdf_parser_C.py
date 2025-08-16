import fitz  # PyMuPDF
import re
from datetime import datetime, timedelta,date
import unicodedata
import os
import fitz  # PyMuPDF
import re
from datetime import datetime, timedelta,date

import fitz
import re
import pprint

#テキストと座標を抽出
def load_date(pdf_path,page_num=1):
    doc = fitz.open(pdf_path)
    page = doc[page_num - 1]
    words= page.get_text("words")
    """
    print("words")
    for a in words:
        print(f"・{a}")
    """
    return words
#色付きセルの座標を抽出
def pick_up_color_area(pdf_path,page_num=1):
    doc = fitz.open(pdf_path)
    page = doc[page_num - 1]
    drawings   = page.get_drawings()
    color_area=[]
    for d in drawings:
        fill = d.get("fill")  # drawings内の塗潰し色を(R, G, B) 形式 or Noneで抽出
        rect = d.get("rect")  #drawings内の座標項目を抽出
        
        if fill is not None and fill != (0.0, 0.0, 0.0):      #色付きだけ抽出 
            color_area.append({"color": fill, "area": rect})
    """
    print(f"color_area")
    for a in color_area:
        print(f"・{a}") 
    """
    return color_area 

def pick_up_year_month_from_PDF_C(pdf_path,page_num=1):
    words=load_date(pdf_path,page_num=page_num)
    # 4桁の数字のみ
    year_pattern = re.compile(r"^\d{4}$")
    # 1～2桁の数字のみ
    month_pattern = re.compile(r"^\d{1,2}$")    
    year=[]
    month=[]
    for word in words:
        x0, y0, x1, y1, text,*_ = word
        x0, y0, x1, y1 = float(x0), float(y0), float(x1), float(y1) #一応float型に変換

        if year_pattern.fullmatch(text) and y1 <= 100:
            year.append({"text":text,"area":(x0, y0, x1, y1)})

        elif month_pattern.fullmatch(text):
            month.append({"text":text,"area":(x0, y0, x1, y1)})
    print(f"year:\n{year}")
    print(f"month:\n{month}")
    return year, month              

#日付け形式のテキストとその座標を返す関数
def extract_dates_with_coords(pdf_path, page_num=1):
    doc = fitz.open(pdf_path)
    page = doc[page_num - 1]

    pattern = re.compile(r"\d{1,2}/\d{1,2}")  # YYYY/M/D または YYYY/MM/DD

    results = []

    text_dict = page.get_text("dict")
    #pprint.pprint(text_dict['blocks'][0])

    for block in text_dict["blocks"]:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"]
                    # 部分一致で日付を探す
                    match = pattern.fullmatch(text)#patternと完全一致
                    if match:
                        x0, y0, x1, y1 = span["bbox"]
                        results.append({"text": match.group(),"area":( x0, y0, x1, y1)})
    
    print("日付00/00型抽出")
    for a in results:
        print(f"・{a}")
    
    return results

#特定の文字の座標を取得
def find_word_positions(pdf_path, keyword,search_height=200):
    words=load_date(pdf_path,page_num=1)
    keyword_positions=[]
    for word in words:
        x0, y0, x1, y1, text,*_ = word
        x0, y0, x1, y1 = float(x0), float(y0), float(x1), float(y1) #一応float型に変換
        if y1 <= search_height and text == keyword:   #keywordが完全一致
            keyword_positions.append({"text": text,"area":(x0, y0, x1, y1)})  
    print(f"{keyword}positions:\n{keyword_positions}")
    return keyword_positions #list内dict

#特定の文字のx軸の範囲内の情報を取得
def extract_text_in_xrange(pdf_path, keyword, add=10,sub=10,page_num=1):
    
    words = load_date(pdf_path,page_num=page_num)
    keyword_positions=find_word_positions(pdf_path, keyword,search_height=200 )
    first_keyword = keyword_positions[0]#keywordが複数存在した場合最初の値を採用する
    x_min=first_keyword["area"][0]-sub
    x_max=first_keyword["area"][2]+add
    y_min=first_keyword["area"][1]#keywordのy座標,ヘッダーをlist除外するため必要

    target_column=[]
    # X座標範囲でフィルタ
    for word in words:
        x0, y0, x1, y1, text,*_=word
        x0, y0, x1, y1 = float(x0), float(y0), float(x1), float(y1) #一応float型に変換
        if x0>=x_min and x1<=x_max and y0>y_min:
            target_column.append({"text":text,"area":(x0, y0, x1, y1)})
    """
    print(f"{keyword}列の情報取得")
    for a in target_column:
        print(f"・{a}")
    """
    return target_column

def pick_up_names_from_PDF_C(pdf_path):
    prepare_MAIN=extract_text_in_xrange(pdf_path, "心肺準備", add=10,sub=10,page_num=1)
    prepare_SCP=extract_text_in_xrange(pdf_path, "SCP", add=10,sub=10,page_num=1)    
    prepare_CP=extract_text_in_xrange(pdf_path, "CP準備", add=10,sub=10,page_num=1)
    prepare_OTHERS=extract_text_in_xrange(pdf_path, "外回業務準備", add=20,sub=20,page_num=1)
    prepare_hinotori=extract_text_in_xrange(pdf_path, "hinotori対応", add=10,sub=10,page_num=1)

    names=[]
    for n in prepare_MAIN:
        prepare_MAIN_names=n["text"]
        names.append(prepare_MAIN_names)
    for n in prepare_SCP:
        prepare_SCP_names=n["text"]
        names.append(prepare_SCP_names)    
    for n in prepare_CP:
        prepare_CP_names=n["text"]
        names.append(prepare_CP_names)
    for n in prepare_OTHERS:
        prepare_OTHERS_names=n["text"]
        names.append(prepare_OTHERS_names)
    for n in prepare_hinotori:
        prepare_hinotori_names=n["text"]
        names.append(prepare_hinotori_names)
    unique_names = list(set(names))


    
    print("unique_names:")
    for a in unique_names:
        print(f"・{a}")
    
    return unique_names
#特定の列内の情報と日付列の情報を合成する
def marge_datelist_and_target_column(pdf_path,keyword,page_num=1,add=10,sub=10,min_diff=1):
    datelist=extract_dates_with_coords(pdf_path, page_num=page_num)
    target_line=extract_text_in_xrange(pdf_path,keyword,add=add,sub=sub,page_num=page_num)

    for target in target_line:
        target_y0 = target["area"][1]
        # date_list の中で最も x0 が近いものを探す
        for date in datelist:
            date_y0 = date["area"][1]
            diff = abs(target_y0 - date_y0)
            if diff <= min_diff:
                target["date"] = date["text"]
                target["description"]=keyword
    
    print(f"{keyword}列に日付を合成")
    for a in target_line:
        print(f"・{a}")
    
    return target_line

def check_contain_color_area(pdf_path,keyword,page_num=1,add=10,sub=10,min_diff=1):
    color_area=pick_up_color_area(pdf_path,page_num=page_num,)
    target_line=marge_datelist_and_target_column(pdf_path,keyword,page_num=page_num,add=add,sub=sub,min_diff=min_diff)

    for target in target_line:
        x0_t=target["area"][0]
        y0_t=target["area"][1]
        x1_t=target["area"][2]
        y1_t=target["area"][3]
        for color in color_area:
            x0_c=color["area"][0]
            y0_c=color["area"][1]
            x1_c=color["area"][2]
            y1_c=color["area"][3]
            if x0_t>=x0_c and x1_c>=x1_t and y0_t>=y0_c and y1_c>=y1_t:
                target["color"]=color["color"]
    """            
    print(f"{keyword}列にcolorを追加")
    for a in target_line:
        print(f"・{a}")
    """
    return target_line

def convert_extracted_column_for_google(pdf_path,keyword,page_num=1,add=10,sub=10,min_diff=1):

    target_line=check_contain_color_area(pdf_path,keyword,page_num=page_num,add=add,sub=sub,min_diff=min_diff)
    year,month=pick_up_year_month_from_PDF_C(pdf_path,page_num=1)
    year_text=int(year[0]['text'])
    month_text=int(month[0]['text'])
    for target in target_line:
        m,d=target["date"].split("/")
        modify_date = f"{year_text}-{m.zfill(2)}-{d.zfill(2)}"#2025-8-20の形に変形
        target["date"] = modify_date  # 値を上書き
    print("target_line(日付表示を修正):")
    for a in target_line:
        print(f"・{a}")
    
    timezone = "Asia/Tokyo"  # 日本時間
    convert_for_google=[]
    for target in target_line:
        r,g,b=target['color']
        date=target["date"]
        name=target["text"]
        work=target["description"]
        if 0.78<r<=1 and 0.78<g<=1 and 0<=b<0.59:#黄色判定
                start = f"{date}T07:30:00"
                end = f"{date}T17:15:00"
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "OP早出7:30",
                                    "description": f"勤務内容:{work} 勤務表:OP 職員:{name}"})
        if 0.78<r<=1 and 0.39<g<=0.78 and 0<=b<0.39:#オレンジ色判定
                start = f"{date}T08:00:00"
                end = f"{date}T17:15:00"
                convert_for_google.append({"start": {"dateTime": start,"timeZone": timezone},
                                    "end": {"dateTime": end,"timeZone": timezone},
                                    "summary": "OP早出8:00",
                                    "description": f"勤務内容:{work} 勤務表:OP 職員:{name}"})
    

    
    
    
    print("convert_for_google")
    for a in convert_for_google:
        print(f"・{a}")
    #print(convert_for_google)
    return convert_for_google

#test用
if __name__ == "__main__":


    # テスト対象のファイルパス
    upload_dir = "uploads"
    #test_filename = "勤務表2025.8ver4.pdf"
    test_filename = "2025.8手術室業務担当予定表test.pdf"
    test_path = os.path.join(upload_dir, test_filename)
    keyword="hinotori対応"
    #all_words=load_date(test_path,page_num=1)#すべてのテキストが適切に解析分離抽出可能なことを確認
    #pick_up_color_area(test_path,page_num=1)
    #extract_dates_with_coords(test_path,page_num=1)
    #marge_datelist_and_target_column(test_path,keyword,page_num=1,add=10,sub=10,min_diff=1)
    #check_contain_color_area(test_path,keyword,page_num=1,add=10,sub=10,min_diff=1)
    #load_date(test_path,page_num=1)
    #pick_up_year_month_from_PDF_C(test_path,page_num=1)
    convert_extracted_column_for_google(test_path,keyword,page_num=1,add=10,sub=10,min_diff=1)
    #pick_up_names_from_PDF_C(test_path)