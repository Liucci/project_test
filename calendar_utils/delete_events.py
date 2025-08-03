# calendar_utils/delete_events.py
from googleapiclient.discovery import Resource
from typing import List

def delete_events(service: Resource, calendar_id: str, events: List[dict]):
    """
    指定されたイベントリストをすべて削除する
    :param service: Google Calendar API サービス
    :param calendar_id: カレンダーID
    :param events: 削除対象イベントのリスト
    """
    for event in events:
        event_id = event.get('id')
        summary = event.get('summary', '')
        start = event.get('start', {})
        print(f"削除中: {summary} - {start}")
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
