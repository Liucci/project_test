# app.py（更新後のロジックフロー）

from flask import Flask, request, render_template, redirect, url_for, session
import os, pandas as pd
import re
from datetime import datetime, timedelta
from pytz import timezone
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from calendar_utils.pick_up_events import pick_up_events
from calendar_utils.delete_events import delete_events
from werkzeug.utils import secure_filename
from pdf_utils.pdf_parser import extract_names_from_PDF_A, get_schedule_month_from_PDF_A,extract_schedule_from_PDF_A
from pdf_utils.pdf_parser_4llm import extract_schedule_from_markdown, extract_names_from_pdf_with_4llm, get_schedule_month_from_pdf_with_4llm
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

    def unique_filename(original_filename, tag):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        name, ext = os.path.splitext(secure_filename(original_filename))
        return f"{tag}_{timestamp}{ext}"

    if request.method == "POST":
        # 初期化
        session.pop("names", None)
        session.pop("year", None)
        session.pop("month", None)

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
            print(f"[DEBUG] Uploaded PDF: {path_PDF_A}")
        else:
            path_PDF_A=None
            session["path_PDF_A"]=path_PDF_A
            

        if file_PDF_B and file_PDF_B.filename:
            filename_PDF = unique_filename(file_PDF_B.filename, "PDF")
            path_PDF_B= os.path.join(app.config['UPLOAD_FOLDER'], filename_PDF)
            file_PDF_B.save(path_PDF_B)
            session["path_PDF_B"] = path_PDF_B
            print(f"[DEBUG] Uploaded PDF: {path_PDF_B}")
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
                return render_template("error.html", message="勤務表から職員名簿作成失敗") 
        elif path_PDF_B:
            try:
                names=extract_names_from_PDF_B(path_PDF_B)
                print("職員名")
                for a in names:
                    print("・", a)
            except:
                return render_template("error.html", message="血液浄化センター勤務表から職員名簿作成失敗")
        
        else:
            return render_template("error.html",messeage="勤務表が1つもアップロードされていません")
        

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
        return render_template("error.html", message="職員名が未指定です。")

    html_events = []  # ← ★ HTML表示用イベント

    # 勤務表PDFからのスケジュール抽出
    if path_PDF_A:
        try:
            html_events_A = extract_schedule_from_markdown(path_PDF_A, selected_name)
            html_events.extend(html_events_A)
            session["year_month_pdf_A"] = get_schedule_month_from_PDF_A(path_PDF_A)
        except Exception as e:
            return render_template("error.html", message=f"PDF解析中にエラー: {e}")

    if path_PDF_B:
        try:
            html_events_B = extract_HD_schedule_from_PDF_B(path_PDF_B,year_B, selected_name,y_tolerance=5)
            html_events.extend(html_events_B)
            session["month_B"]=extract_month_from_PDF_B(path_PDF_B)#eventをdeleteするときに年月が必要、年と月別のほうがpick_eventsに渡しやすい
        except Exception as e:
            return render_template("error.html", message=f"血液浄化センター勤務表解析中にエラー: {e}")


    # ★ HTML用イベントをセッションに保存
    session["html_events"] = html_events
    print("html_events:")
    for i, ev in enumerate(html_events[:3]):
        print(f"[DEBUG] html_events[{i}] type: {type(ev)}")
        for key in ['date', 'start', 'end', 'summary', 'description']:
            print(f"    {key}: {ev.get(key)} (type: {type(ev.get(key))})")
        

    return render_template("show_schedule.html", selected_name=selected_name, html_events=html_events)


 

@app.route("/authorize")
def authorize():
    redirect_uri = url_for("oauth2callback", _external=True, _scheme="http" if os.getenv("FLASK_ENV") == "development" else "https")
    flow = Flow.from_client_secrets_file(CLIENT_SECRET_FILE, scopes=SCOPES, redirect_uri=redirect_uri)
    authorization_url, state = flow.authorization_url(access_type="offline", include_granted_scopes="true", prompt='select_account consent')
    session["state"] = state
    return redirect(authorization_url)

@app.route("/oauth2callback")
def oauth2callback():
    state = session.get("state")
    if state != request.args.get("state"):
        return "CSRFエラー", 400
    flow = Flow.from_client_secrets_file(CLIENT_SECRET_FILE, scopes=SCOPES, state=state, redirect_uri=url_for("oauth2callback", _external=True, _scheme="http" if os.getenv("FLASK_ENV") == "development" else "https"))
    flow.fetch_token(authorization_response=request.url)
    credentials = flow.credentials
    session["credentials"] = credentials_to_dict(credentials)
    return render_template("upload.html")

@app.route("/delete_registered_events")
def delete_registered_events():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    path_PDF_A = session.get("path_PDF_A")
    path_PDF_B=session.get("path_PDF_B")
    all_events_to_delete = []  # ← 全ての削除対象をここに集約

    
    
    if path_PDF_A:
        year_month_pdf_A = session.get("year_month_pdf_A", (None, None))
        if all(year_month_pdf_A):
            year_A, month_A = year_month_pdf_A
            events_to_delete_pdf_A= pick_up_events(
                service,
                calendar_id="primary",
                year=year_A,
                month=month_A,
                tag="MAIN"
            )
            print(f"[DEBUG] PDF(MAIN)から {len(events_to_delete_pdf_A)} 件削除予定")
            all_events_to_delete.extend(events_to_delete_pdf_A)
    if path_PDF_B:
        year_B=session.get("year_B")
        month_B=session.get("month_B")
        events_to_delete_pdf_B= pick_up_events(
                service,
                calendar_id="primary",
                year=year_B,
                month=month_B,
                tag="HD"
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

@app.route("/delete_events_specificed_term",methods=["POST"])
def delete_events_specificed_term():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    selected_year = int(request.form.get("year"))
    selected_month= int(request.form.get("month"))
    events_to_delete_specificed_term= pick_up_events(
                service,
                calendar_id="primary",
                year=selected_year,
                month=selected_month,
                tag="HD" or "MAIN"
            )
    session["deleted_events"] = events_to_delete_specificed_term.copy()
    
    delete_events(service, calendar_id="primary", events=events_to_delete_specificed_term)
    return render_template("result2.html",deleted_events=events_to_delete_specificed_term,selected_year=selected_year,selected_month=selected_month)

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

    path_PDF_A = session.get("path_PDF_A")
    if path_PDF_A: 
        os.remove(path_PDF_A)
        session.pop("path_PDF_A", None)
        print(f"[DEBUG] PDFファイル {path_PDF_A} を削除しました。")

    path_PDF_B = session.get("path_PDF_B")
    if path_PDF_B: 
        os.remove(path_PDF_B)
        session.pop("path_PDF_A", None)
        print(f"[DEBUG] PDFファイル {path_PDF_B} を削除しました。")
    


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

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
