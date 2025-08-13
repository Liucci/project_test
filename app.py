# app.py（更新後のロジックフロー）

from flask import Flask, request, render_template, redirect, url_for, session
import os, pandas as pd
import re
import uuid
from datetime import datetime, timedelta,date
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from calendar_utils.pick_up_events import pick_up_events
from calendar_utils.delete_events import delete_events
from werkzeug.utils import secure_filename
from pdf_utils.pdf_parser_A import extract_names_from_PDF_A, get_schedule_month_from_PDF_A,extract_schedule_from_PDF_A
from pdf_utils.pdf_parser_B import extract_HD_schedule_from_PDF_B,extract_names_from_PDF_B,extract_month_from_PDF_B
from flask import Flask
from flask_session import Session  # ← 追加
from werkzeug.utils import secure_filename
app = Flask(__name__)


# ---- Flask-Sessionの設定 ----
app.config["SESSION_TYPE"] = "filesystem"  # サーバー側に保存
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_FILE_DIR"] = "flask_session"  # 保存先フォルダ（任意）
app.config["SESSION_USE_SIGNER"] = True  # セキュリティ強化

# ---- Sessionを初期化 ----
Session(app)


# 環境設定
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
load_dotenv()


app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev-secret-key-for-local")

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/userinfo.email",
]
CLIENT_SECRET_FILE = 'credentials.json'

@app.route("/", methods=["GET"])
def index():
    
        return render_template("index.html")




@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    # セッションからPDFファイルを削除
    # 残っているPDFファイルを削除する
    delete_upload_file("path_PDF_A")
    delete_upload_file("path_PDF_B")
    delete_session_keys("names",
                        "year_month_pdf_A",
                        "year_B",
                        "month_B",
                        "file_name_PDF_A_origin",
                        "file_name_PDF_B_origin",
                        "selected_name",
                        "deleted_events",
                        "html_events",
                        "year",
                        "month",
                        "selected_year",
                        "selected_month",
                        "events_to_delete_specificed_term",
                        "tags"
                        )
    print("session中身:")
    for a,b in session.items():
        print(f"・{a}:{b}")


    def unique_filename(original_filename, tag):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        random_str = uuid.uuid4().hex[:8]  # 8桁ランダムID
        name, ext = os.path.splitext(secure_filename(original_filename))
        return f"{tag}_{timestamp}_{random_str}{ext}"
    
    if request.method == "POST":
        #各ファイルをhtmlから受け取りre-name前の名前をsessionに保存
        file_PDF_A = request.files.get("file_PDF_A")
        session["file_name_PDF_A_origin"]=file_PDF_A.filename
        file_PDF_B = request.files.get("file_PDF_B")
        session["file_name_PDF_B_origin"]=file_PDF_B.filename

      
        # 年の抽出（ファイル名から）
        # 勤務表BはPDF内から年度を取り出せないためre-name前のfile名から取得しておく
        match = re.search(r"(20\d{2})", file_PDF_B.filename)
        if match:
                    year= int(match.group(1))
        else:
                    year= None
        session["year_B"] = year

        # PDFファイルの処理（オプション）
        # ファイルがアップロードされているか確認
        #安全と重複して上書きを避けるためfile名を変更
        #sessionにはファイルを保存できないのでpathとして保存する
        if file_PDF_A and file_PDF_A.filename:
            filename_PDF = unique_filename(file_PDF_A.filename, "PDF")
            path_PDF_A= os.path.join(app.config['UPLOAD_FOLDER'], filename_PDF)
            file_PDF_A.save(path_PDF_A)
            session["path_PDF_A"] = path_PDF_A
            print(f"[DEBUG] Uploaded PDF_A: {path_PDF_A}")
        else:
            path_PDF_A=None
            session["path_PDF_A"]=path_PDF_A
            

        if file_PDF_B and file_PDF_B.filename:
            filename_PDF = unique_filename(file_PDF_B.filename, "PDF")
            path_PDF_B= os.path.join(app.config['UPLOAD_FOLDER'], filename_PDF)
            file_PDF_B.save(path_PDF_B)
            session["path_PDF_B"] = path_PDF_B
            print(f"[DEBUG] Uploaded PDF_B: {path_PDF_B}")
        else:
            path_PDF_B=None
            session["path_PDF_B"]=path_PDF_B
            




        # 職員名の抽出

        if path_PDF_A:           
            try:
                names = extract_names_from_PDF_A(path_PDF_A)
                print("職員名")
                for a in names:
                    print("・", a)
            except Exception as e:
                return render_template("error_back_to_upload.html", message=f"勤務表から職員名簿作成失敗{e}") 
        elif path_PDF_B:
            try:
                names=extract_names_from_PDF_B(path_PDF_B)
                print("職員名")
                for a in names:
                    print("・", a)
            except Exception as e:
                return render_template("error_back_to_upload.html", message=f"血液浄化センター勤務表から職員名簿作成失敗{e}")
        
        else:
            return render_template("error_back_to_upload.html",messeage="勤務表が1つもアップロードされていません")
        

        session["names"]=names
        return render_template("select_name.html", names=names)

    # GET の場合：フォームを表示
    return render_template("upload.html")


@app.route("/select", methods=["POST","GET"])
def select_name():
    if request.method == "POST":
        selected_name = request.form.get("selected_name")
        if not selected_name:
            return "職員名が選択されていません。", 400

        session["selected_name"] = selected_name
        return redirect(url_for("show_schedule"))
    else:
        names=session.get("names")
        return render_template("select_name.html",names=names)
@app.route("/schedule")
def show_schedule():
    print("[DEBUG] show_schedule called")
    selected_name = session.get("selected_name")
    path_PDF_A = session.get("path_PDF_A")
    path_PDF_B=session.get("path_PDF_B")
    year_B=session.get("year_B")

    if not selected_name:
        return render_template("error_back_to_upload.html", message="職員名が未指定です。")

    html_events = []  # ← ★ HTML表示用イベント

    # 勤務表PDFからのスケジュール抽出
    if path_PDF_A:
        try:
            html_events_A = extract_schedule_from_PDF_A(path_PDF_A, selected_name,x_tolerance=7)
            if html_events_A is not None:
                html_events.extend(html_events_A)
                session["year_month_pdf_A"] = get_schedule_month_from_PDF_A(path_PDF_A)
        except Exception as e:
            return render_template("error_back_to_upload.html", message=f"PDF解析中にエラー: {e}")

    if path_PDF_B:
        try:
            html_events_B = extract_HD_schedule_from_PDF_B(path_PDF_B,year_B, selected_name,y_tolerance=5)
            session["month_B"]=extract_month_from_PDF_B(path_PDF_B)#eventをdeleteするときに年月が必要、年と月別のほうがpick_eventsに渡しやすい
            if html_events_B is not None and html_events_B:
                html_events.extend(html_events_B)
                
        except Exception as e:
            return render_template("error_back_to_upload.html", message=f"血液浄化センター勤務表解析中にエラー: {e}")


    # ★ HTML用イベントをセッションに保存
    #並び替え
    html_events = sorted(
    html_events,
    key=lambda e: e["start"].get("dateTime") or e["start"].get("date")
)
    session["html_events"] = html_events
   
   
    print("html_events:")
    for i, ev in enumerate(html_events):
        print(f"[DEBUG] html_events[{i}] type: {type(ev)}")
        for key in ['start', 'end', 'summary', 'description']:
            print(f"{key}: {ev.get(key)} (type: {type(ev.get(key))})")
        

    return render_template("show_schedule.html", selected_name=selected_name, html_events=html_events)


 

@app.route("/authorize")
def authorize():
    try:
        redirect_uri = url_for("oauth2callback", _external=True, _scheme="http" if os.getenv("FLASK_ENV") == "development" else "https")
        flow = Flow.from_client_secrets_file(CLIENT_SECRET_FILE, scopes=SCOPES, redirect_uri=redirect_uri)
        authorization_url, state = flow.authorization_url(access_type="offline", include_granted_scopes="true", prompt='select_account consent')
        session["state"] = state
        return redirect(authorization_url)
    
    except Exception as e:
                return render_template("error_back_to_index.html", message=f"Google認証失敗{e}") 

@app.route("/oauth2callback")
def oauth2callback():
    try:
        state = session.get("state")
        if state != request.args.get("state"):
            return "CSRFエラー", 400
        flow = Flow.from_client_secrets_file(CLIENT_SECRET_FILE, scopes=SCOPES, state=state, redirect_uri=url_for("oauth2callback", _external=True, _scheme="http" if os.getenv("FLASK_ENV") == "development" else "https"))
        flow.fetch_token(authorization_response=request.url)
        credentials = flow.credentials
        session["credentials"] = credentials_to_dict(credentials)
        return render_template("upload.html")
    except Exception as e:
                return render_template("error_back_to_index.html", message=f"Google認証失敗{e}") 

@app.route("/delete_registered_events")
def delete_registered_events():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    path_PDF_A = session.get("path_PDF_A")
    path_PDF_B=session.get("path_PDF_B")
    all_events_to_delete = []  # ← 全ての削除対象をここに集約
    print("[DEBUG] path_PDF_B repr:", repr(path_PDF_B), type(path_PDF_B))
    year_B=session.get("year_B")
    month_B=session.get("month_B")  
    print(f"year_B:\n{year_B}")
    print(f"month_B:\n{month_B}")

    
    
    if path_PDF_A:
        year_A,month_A = session.get("year_month_pdf_A", (None, None))
        events_to_delete_pdf_A= pick_up_events(
                                                service,
                                                calendar_id="primary",
                                                year=int(year_A),
                                                month=int(month_A),
                                                tags=["MAIN"]
                                            )
        print(f"[DEBUG] PDF(MAIN)から {len(events_to_delete_pdf_A)} 件削除予定")
        all_events_to_delete.extend(events_to_delete_pdf_A)
    if  path_PDF_B not in (None, "", "None"):
        year_B=int(session.get("year_B"))
        month_B=int(session.get("month_B"))
        events_to_delete_pdf_B= pick_up_events(
                service,
                calendar_id="primary",
                year=year_B,
                month=month_B,
                tags=["HD"]
            )
        print(f"[DEBUG] PDF(HD)から {len(events_to_delete_pdf_B)} 件削除予定")
        all_events_to_delete.extend(events_to_delete_pdf_B)

    print(f"[DEBUG] 全消去リスト: {len(all_events_to_delete)}")

    # ===== セッション保存・一括削除 =====
    # 削除対象を先に保存（必要なら .copy() で保護）
    #htmlで表示させるためにコピーしておく
    session["deleted_events"] = all_events_to_delete.copy()

    # 実際にall_events_to_deleteが削除される
    delete_events(service, calendar_id="primary", events=all_events_to_delete)

    return redirect(url_for("upload_to_calendar"))

@app.route("/pick_up_delete_events_specificed_term",methods=["POST"])
def pick_up_delete_events_specificed_term():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    selected_year = int(request.form.get("year"))
    selected_month= int(request.form.get("month"))
    session["selected_year"]=selected_year
    session["selected_month"]=selected_month
    tags=request.form.getlist("tags")

    if selected_year and selected_month and tags:
        events_to_delete_specificed_term= pick_up_events(
                service,
                calendar_id="primary",
                year=selected_year,
                month=selected_month,
                tags=tags
            )
    else:
        return render_template("error_back_to_upload.html", message=f"検索条件を設定してください。") 
         
    print(f"events_to_delete_specificed_term:\n{events_to_delete_specificed_term}")
    session["events_to_delete_specificed_term"] = events_to_delete_specificed_term
    return render_template("show_schedule_to_delete_events.html",
                           events_to_delete_specificed_term=events_to_delete_specificed_term,
                           selected_year=selected_year,
                           selected_month=selected_month)

@app.route("/delete_events_specificed_term",methods=["POST"])
def delete_events_specificed_term():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    events_to_delete_specificed_term=session.get("events_to_delete_specificed_term")
    selected_year=session.get("selected_year")
    selected_month=session.get("selected_month")
    if events_to_delete_specificed_term:
        delete_events(service, calendar_id="primary", events=events_to_delete_specificed_term)#events_todelete_specificed_termは空になる
        deleted_events=session.get("events_to_delete_specificed_term")#events_to_delete_specificed_termから再度データを取得
        print(f"{selected_year}年{selected_month}月の\n{deleted_events}を削除しました。")
        return render_template("result2.html",
                            deleted_events=deleted_events,
                            selected_year=selected_year,
                            selected_month=selected_month)
    else:
        return render_template("error_back_to_upload.html", message=f"イベント削除時にエラーが発生しました。") 

@app.route("/upload_to_calendar")
def upload_to_calendar():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)

    html_events = session.get("html_events", [])  # HTML表示用イベント
    selected_name = session.get("selected_name")
    deleted_events = session.get("deleted_events", [])

    # Google Calendarへイベント登録

    for event in html_events:
        service.events().insert(calendarId="primary", body=event).execute()



    print("[DEBUG] deleted_events type:", type(deleted_events))  # listであるべき
    print("[DEBUG] deleted_events（簡略版）:")
    for i, ev in enumerate(deleted_events[:3]):
        print(f"[DEBUG] deleted_events[{i}] type: {type(ev)}")
        for key in ['date', 'start', 'end', 'summary', 'description']:
            print(f"    {key}: {ev.get(key)} (type: {type(ev.get(key))})")
    
    # セッションからPDFファイルを削除
    # ここでPDFファイルを削除する
    delete_upload_file("path_PDF_A")
    delete_upload_file("path_PDF_B")
    



    return render_template(
        "result.html",
        selected_name=selected_name,
        html_events=html_events,
        deleted_events=deleted_events,
        
    )

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))





@app.route('/privacy')
def privacy():
    return render_template("privacy.html")

@app.route('/terms')
def terms():
    return render_template("terms.html")

@app.route('/about')
def about():
    return render_template("about.html")

@app.route("/contact")
def contact():
    return render_template("contact.html")


def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

def dict_to_credentials(d):
    return Credentials(**d)

#file_pathとsession内のfile_pathも消す関数
def delete_upload_file(session_key):
    file_path = session.get(session_key)
    if file_path:
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"[DEBUG] PDFファイル {file_path} を削除しました。")
        else:
            print(f"[DEBUG] PDFファイル {file_path} は存在しませんでした。")
        session.pop(session_key, None)
        print(f"[DEBUG] session の {session_key} を削除しました。")
    else:
        print(f"[DEBUG] session に {session_key} の情報がありません。")

#session消す関数
def delete_session_keys(*keys):
    for key in keys:
        session.pop(key, None)
    
    print("削除したsession key")
    for a in keys:
        print(f"・{a}")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
