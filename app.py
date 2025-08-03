# app.py（更新後のロジックフロー）

from flask import Flask, request, render_template, redirect, url_for, session
import os, pandas as pd
from excel_utils.excel_parser import extract_names_from_excel
from datetime import datetime, timedelta
from pytz import timezone
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from calendar_utils.pick_up_events import pick_up_events
from calendar_utils.delete_events import delete_events

# 環境設定
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
load_dotenv()

app = Flask(__name__)
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

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename.endswith(".xlsx"):
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(save_path)
            session['uploaded_file'] = save_path
            names = extract_names_from_excel(save_path)
            session['names'] = names
            return render_template("select_name.html", names=names)
        return "有効な .xlsx ファイルを選択してください。"
    return render_template("upload.html")

@app.route("/select", methods=["GET", "POST"])
def select_name():
    names = session.get("names", [])
    if request.method == "POST":
        selected_name = request.form["selected_name"]
        session["selected_name"] = selected_name
        return redirect(url_for("show_schedule"))
    return render_template("select_name.html", names=names)

@app.route("/schedule")
def show_schedule():
    filepath = session.get("uploaded_file")
    selected_name = session.get("selected_name")
    df = pd.read_excel(filepath, sheet_name="原本", header=None)
    c1, j1 = df.iloc[0, 2], df.iloc[0, 9]
    year, month = (c1.year if isinstance(c1, datetime) else int(c1)), (j1.month if isinstance(j1, datetime) else int(j1))
    session['year'], session['month'] = year, month
    name_rows = df.iloc[8:, :]
    target_row = name_rows[name_rows.iloc[:, 2] == selected_name]
    if target_row.empty:
        return f"{selected_name} の勤務情報が見つかりませんでした"

    start_col, max_days = 9, 31
    work_cells = target_row.iloc[0, start_col:start_col + max_days].tolist()
    days = list(range(1, max_days + 1))
    events, jst = [], timezone("Asia/Tokyo")

    for i, cell in enumerate(work_cells):
        if pd.isna(cell): continue
        day = days[i]
        try: date = datetime(year, month, day)
        except ValueError: continue

        cell = str(cell).strip()
        if cell == "1": summary = "1st on call"
        elif cell == "2": summary = "2nd on call"
        elif cell == "⑯": summary = "当直"
        elif cell == "年休": summary = "有給休暇"
        elif cell in ["振休", "代休"]: summary = "振替休日"
        elif cell == "ＡＭ休":
            events.append({
                "summary": "午後勤務（午前休）",
                "start": {"dateTime": date.strftime("%Y-%m-%dT13:22:30"), "timeZone": "Asia/Tokyo"},
                "end": {"dateTime": date.strftime("%Y-%m-%dT17:15:00"), "timeZone": "Asia/Tokyo"},
                "description": "勤務予定"
            })
            continue
        elif cell == "ＰＭ休":
            events.append({
                "summary": "午前勤務（午後休）",
                "start": {"dateTime": date.strftime("%Y-%m-%dT08:30:00"), "timeZone": "Asia/Tokyo"},
                "end": {"dateTime": date.strftime("%Y-%m-%dT12:22:30"), "timeZone": "Asia/Tokyo"},
                "description": "勤務予定"
            })
            continue
        elif cell == "早HD":
            events.append({
                "summary": "HD早番",
                "start": {"dateTime": date.strftime("%Y-%m-%dT07:30:00"), "timeZone": "Asia/Tokyo"},
                "end": {"dateTime": date.strftime("%Y-%m-%dT16:15:00"), "timeZone": "Asia/Tokyo"},
                "description": "勤務予定"
            })
            continue
        else: continue

        events.append({
            "summary": summary,
            "start": {"date": date.strftime("%Y-%m-%d")},
            "end": {"date": (date + timedelta(days=1)).strftime("%Y-%m-%d")},
            "description": "勤務予定"
        })

    session["parsed_events"] = events
    return render_template("show_schedule.html", name=selected_name, events=events)

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
    return redirect(url_for("delete_registered_events"))

@app.route("/delete_registered_events")
def delete_registered_events():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    year, month, name = session.get("year"), session.get("month"), session.get("selected_name")
    deleted_events = pick_up_events(service, calendar_id='primary', year=year, month=month, tag="勤務予定")
    session['deleted_events'] = deleted_events
    delete_events(service, calendar_id='primary', events=deleted_events)
    return redirect(url_for("upload_to_calendar"))

@app.route("/upload_to_calendar")
def upload_to_calendar():
    credentials = dict_to_credentials(session.get("credentials"))
    service = build("calendar", "v3", credentials=credentials)
    events_to_add = session.get("parsed_events")
    for ev in events_to_add:
        service.events().insert(calendarId="primary", body=ev).execute()
    name = session.get("selected_name")
    deleted_events = session.get("deleted_events", [])
    return render_template("result.html", name=name, added_events=events_to_add, deleted_events=deleted_events)

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

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
