from googleapiclient.discovery import Resource
from typing import List
from datetime import datetime, timezone
def pick_up_events(service: Resource, calendar_id: str, year: int, month: int, tag: str = None) -> List[dict]:
    """
    指定年月とタグで予定を抽出
    :param service: Google Calendar API サービス
    :param calendar_id: カレンダーID（通常 'primary'）
    :param year: 対象年
    :param month: 対象月
    :param tag: description に含まれるタグ（例: [勤務表=MAIN]）
    :return: 該当イベントのリスト
    """


    start = datetime(year, month, 1, tzinfo=timezone.utc)
    if month == 12:
        end = datetime(year + 1, 1, 1, tzinfo=timezone.utc)
    else:
        end = datetime(year, month + 1, 1, tzinfo=timezone.utc)

    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=start.isoformat(),
        timeMax=end.isoformat(),
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    events = events_result.get('items', [])

    if tag:
        filtered_events = [
            event for event in events
            if tag in event.get('description', '')
        ]
    else:
        filtered_events = events
    
    
    simplified_events = []
    for e in filtered_events:
        simplified_events.append({
            "start": e.get("start", {}),
            "end": e.get("end", {}),
            "summary": e.get("summary"),
            "description": e.get("description", ""),
            "id": e.get("id")
        })


    print(f"[DEBUG] pick_up_events() で抽出されたイベント数: {len(simplified_events)}")
    print(f"[DEBUG] pick_up_events()で抽出されたイベントの３例: {simplified_events[:3]}")  # 最初の3つのイベントを表示
    for i, ev in enumerate(simplified_events[:3]):
        print(f"[DEBUG] html_events[{i}] type: {type(ev)}")
        for key in [ 'start', 'end', 'summary', 'description']:
            print(f"    {key}: {ev.get(key)} (type: {type(ev.get(key))})")
        print(f"  {ev['start']}～{ev['end']}: {ev['summary']} | {ev['description']}")    
    
    return simplified_events
