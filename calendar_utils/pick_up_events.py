# calendar_utils/pick_up_events.py
from googleapiclient.discovery import Resource
from typing import List

def pick_up_events(service: Resource, calendar_id: str, year: int, month: int, tag: str = None) -> List[dict]:
    """
    指定年月とタグで予定を抽出
    :param service: Google Calendar API サービス
    :param calendar_id: カレンダーID（通常 'primary'）
    :param year: 対象年
    :param month: 対象月
    :param tag: summaryなどに含まれるタグ（例: 氏名）
    :return: イベントのリスト
    """
    from datetime import datetime, timezone

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

    filtered_events = [event for event in events if event.get('description') == tag]

    return filtered_events
