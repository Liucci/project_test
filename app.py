from flask import Flask, request, render_template, redirect, url_for, session
import os,pandas as pd
from excel_utils.excel_parser import extract_names_from_excel
from datetime import datetime, timedelta
from pytz import timezone
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'#HTTP（非SSL）通信でもOAuth許可する

app = Flask(__name__)
app.secret_key = "your_secret_key"  # セッションにファイルパスを一時保存するため
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)



#スコープとGoogle認証関連の定数を追加
SCOPES = ['https://www.googleapis.com/auth/calendar']
CLIENT_SECRET_FILE = 'credentials.json'  # credentials.json がある場合


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename.endswith(".xlsx"):
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(save_path)

            # ファイルパスをセッションに保存
            session['uploaded_file'] = save_path

            # 名前一覧を抽出して次のページに渡す
            names = extract_names_from_excel(save_path)
            session['names'] = names
            print(f"抽出された名前:\n {names}")
            return render_template("select_name.html", names=names)#テンプレート変数名 = Python変数名

        return "有効な .xlsx ファイルを選択してください。"
    return render_template("upload.html")

@app.route("/select", methods=["GET", "POST"])
def select_name():
    "選択した名前をセッションに保存し、スケジュール表示ページへリダイレクト"
    names = session.get("names", [])
    
    if request.method == "POST":
        selected_name = request.form["selected_name"]
        session["selected_name"] = selected_name
        print(f"選択した名前: \n{selected_name}")
        return redirect(url_for("show_schedule"))
    return render_template("select_name.html", names=names)

@app.route("/schedule")
def show_schedule():
    "選択された名前の勤務情報を表示"
    filepath = session.get("uploaded_file")
    selected_name = session.get("selected_name")

    df = pd.read_excel(filepath, sheet_name="原本", header=None)
    
    # ===== 年・月を取得 =====
    c1 = df.iloc[0, 2]  # C1
    j1 = df.iloc[0, 9]  # J1

    if isinstance(c1, datetime):
        year = c1.year
    else:
        year = int(c1)

    if isinstance(j1, datetime):
        month = j1.month
    else:
        month = int(j1)

    print("C1 (年):", c1, type(c1))
    print("J1 (月):", j1, type(j1))
    name_rows = df.iloc[8:, :]  # 実データ（9行目以降）

    # ===== 指定した名前の行を取得 =====
    target_row = name_rows[name_rows.iloc[:, 2] == selected_name]
    print(f"選択した名前の行:\n {target_row}")
    if target_row.empty:
        return f"{selected_name} の勤務情報が見つかりませんでした"

    # ===== 先頭から最大31日分の勤務データを取得 =====
    start_col = 9  # J列（index=9）
    max_days = 31

    work_cells = target_row.iloc[0, start_col:start_col + max_days].tolist()
    days = list(range(1, max_days + 1))  # [1, 2, ..., 31]
    print(f"取得した勤務データ: \n{work_cells}")
    

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

        if str(cell).strip() == "1":
            # 終日イベント（翌日をend.dateとする）
            event = {
                "summary": "1st on call",
                "start": {
                    "date": date.strftime("%Y-%m-%d")
                },
                "end": {
                    "date": (date + timedelta(days=1)).strftime("%Y-%m-%d")
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "2":
            # 終日イベント（翌日をend.dateとする）
            event = {
                "summary": "2nd on call",
                "start": {
                    "date": date.strftime("%Y-%m-%d")
                },
                "end": {
                    "date": (date + timedelta(days=1)).strftime("%Y-%m-%d")
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "⑯":
            # 終日イベント（翌日をend.dateとする）
            event = {
                "summary": "当直",
                "start": {
                    "date": date.strftime("%Y-%m-%d")
                },
                "end": {
                    "date": (date + timedelta(days=1)).strftime("%Y-%m-%d")
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "年休" :
            # 終日イベント（翌日をend.dateとする）
            event = {
                "summary": "有給休暇",
                "start": {
                    "date": date.strftime("%Y-%m-%d")
                },
                "end": {
                    "date": (date + timedelta(days=1)).strftime("%Y-%m-%d")
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "振休"or str(cell).strip() == "代休":
            # 終日イベント（翌日をend.dateとする）
            event = {
                "summary": "振替休日",
                "start": {
                    "date": date.strftime("%Y-%m-%d")
                },
                "end": {
                    "date": (date + timedelta(days=1)).strftime("%Y-%m-%d")
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "ＡＭ休":
            # 午前休（例: 9:00-13:00 の時間帯イベントとして設定）
            event = {
                "summary": "午前休",
                "start": {
                    "dateTime": date.strftime("%Y-%m-%dT08:30:00"),
                    "timeZone": "Asia/Tokyo"
                },
                "end": {
                    "dateTime": date.strftime("%Y-%m-%dT12:22:30"),
                    "timeZone": "Asia/Tokyo"
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "ＰＭ休":
            # 午後休（例: 13:00-17:00 の時間帯イベントとして設定）
            event = {
                "summary": "午後休",
                "start": {
                    "dateTime": date.strftime("%Y-%m-%dT12:22:30"),
                    "timeZone": "Asia/Tokyo"
                },
                "end": {
                    "dateTime": date.strftime("%Y-%m-%dT17:15:00"),
                    "timeZone": "Asia/Tokyo"
                }
            }
            events.append(event)
            continue
        elif str(cell).strip() == "早HD":
            # 早番（例: 8:30-17:15 の時間帯イベントとして設定）
            event = {
                "summary": "HD早番",
                "start": {
                    "dateTime": date.strftime("%Y-%m-%dT07:30:00"),
                    "timeZone": "Asia/Tokyo"
                },
                "end": {
                    "dateTime": date.strftime("%Y-%m-%dT16:15:00"),
                    "timeZone": "Asia/Tokyo"
                }
            }
            events.append(event)
            continue
        else:
            continue  # 他の勤務種別は無視（または将来追加）

    session["events"] = events

    # === 確認用ログ ===
    print(f"=== {selected_name} の勤務イベント ===")
    for e in events:
            print(f"{e['start']} → {e['end']}:{e['summary']} ")

    return render_template("show_schedule.html", name=selected_name, events=events)

#Googleログイン開始
@app.route("/authorize")
def authorize():
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRET_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True)
    )
    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true"
    )
    session["state"] = state
    return redirect(authorization_url)

#認証後にトークン受け取り
@app.route("/oauth2callback")
def oauth2callback():
    state = session["state"]
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRET_FILE,
        scopes=SCOPES,
        state=state,
        redirect_uri=url_for("oauth2callback", _external=True)
    )
    flow.fetch_token(authorization_response=request.url)

    credentials = flow.credentials
    session["credentials"] = {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "token_uri": credentials.token_uri,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "scopes": credentials.scopes
    }
    return redirect(url_for("upload_to_calendar"))

#Googleカレンダーへ書き込み
@app.route("/upload_to_calendar")
def upload_to_calendar():
    if "credentials" not in session:
        return redirect("authorize")

    creds = Credentials(**session["credentials"])
    service = build("calendar", "v3", credentials=creds)

    events = session.get("events", [])
    if not events:
        return "カレンダーに登録する勤務予定がありません。"

    for event in events:
        service.events().insert(calendarId="primary", body=event).execute()

    return "勤務予定をGoogleカレンダーに登録しました。"


if __name__ == "__main__":
    app.run(host="localhost", port=5000, debug=True)

