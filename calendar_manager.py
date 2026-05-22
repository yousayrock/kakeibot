#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Calendar 管理モジュール
家系Bot用 - カレンダーの追加・取得・削除・リマインダー
"""

import pickle
import logging
from pathlib import Path
from datetime import datetime, timedelta

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

logger = logging.getLogger("kakeibo")

SCOPES = ['https://www.googleapis.com/auth/calendar']
TOKEN_PATH = Path('token.pickle')
CREDENTIALS_PATH = Path('credentials.json')

_service = None  # キャッシュ用


def get_service():
    """Google Calendar APIサービスを取得（認証済み）"""
    global _service
    if _service:
        return _service

    creds = None
    if TOKEN_PATH.exists():
        with open(TOKEN_PATH, 'rb') as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_PATH), SCOPES
            )
            # ブラウザで認証
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'wb') as f:
            pickle.dump(creds, f)
        logger.info("✅ Google Calendar 認証完了")

    _service = build('calendar', 'v3', credentials=creds)
    return _service


def add_event(calendar_id: str, title: str, start_dt: datetime,
              end_dt: datetime = None, description: str = '',
              reminder_minutes: int = 30) -> dict:
    """
    予定を追加する

    Args:
        calendar_id: カレンダーID（'primary' または お店カレンダーID）
        title: 予定のタイトル
        start_dt: 開始日時
        end_dt: 終了日時（Noneの場合は1時間後）
        description: 詳細・メモ
        reminder_minutes: 何分前に通知するか

    Returns:
        作成されたイベントの辞書
    """
    service = get_service()

    if end_dt is None:
        end_dt = start_dt + timedelta(hours=1)

    event = {
        'summary': title,
        'description': description,
        'start': {
            'dateTime': start_dt.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'end': {
            'dateTime': end_dt.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': reminder_minutes},
            ],
        },
    }

    result = service.events().insert(calendarId=calendar_id, body=event).execute()
    logger.info(f"📅 予定追加: {title} {start_dt.strftime('%Y-%m-%d %H:%M')}")
    return result


def get_events(calendar_id: str, days: int = 7) -> list:
    """
    今後の予定を取得する

    Args:
        calendar_id: カレンダーID
        days: 何日先まで取得するか

    Returns:
        イベントのリスト
    """
    service = get_service()

    now = datetime.now().astimezone()
    end = now + timedelta(days=days)

    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=now.isoformat(),
        timeMax=end.isoformat(),
        maxResults=20,
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    return events_result.get('items', [])


def get_upcoming_for_reminder(calendar_ids: list, minutes_ahead: int = 30) -> list:
    """
    リマインダー用：指定分後に始まる予定を取得

    Args:
        calendar_ids: チェックするカレンダーIDのリスト
        minutes_ahead: 何分前にリマインドするか

    Returns:
        該当するイベントのリスト（_calendar_id と _calendar_name を付加）
    """
    service = get_service()

    now = datetime.now().astimezone()
    target_time = now + timedelta(minutes=minutes_ahead)
    # ±1分のウィンドウで検索
    window_start = target_time - timedelta(minutes=1)
    window_end   = target_time + timedelta(minutes=1)

    upcoming = []
    for cal_info in calendar_ids:
        cal_id   = cal_info['id']
        cal_name = cal_info.get('name', '')
        try:
            events_result = service.events().list(
                calendarId=cal_id,
                timeMin=window_start.isoformat(),
                timeMax=window_end.isoformat(),
                maxResults=5,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            for event in events_result.get('items', []):
                event['_calendar_id']   = cal_id
                event['_calendar_name'] = cal_name
                upcoming.append(event)
        except Exception as e:
            logger.warning(f"リマインダーチェックエラー ({cal_id}): {e}")

    return upcoming


def delete_event(calendar_id: str, event_id: str):
    """予定を削除する"""
    service = get_service()
    service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
    logger.info(f"🗑️ 予定削除: {event_id}")


def format_event_line(event: dict) -> str:
    """予定を1行テキストにフォーマット"""
    start = event['start'].get('dateTime', event['start'].get('date', ''))
    if 'T' in start:
        dt = datetime.fromisoformat(start)
        time_str = dt.strftime('%m/%d(%a) %H:%M')
    else:
        time_str = start
    title = event.get('summary', '（無題）')
    return f"📅 {time_str}　{title}"


def format_events_list(events: list, calendar_name: str, days: int) -> str:
    """複数の予定をDiscord向けにフォーマット"""
    if not events:
        return f"📭 {calendar_name}に今後{days}日以内の予定はありません。"

    lines = [f"📅 **{calendar_name}（今後{days}日間）**\n"]
    for event in events:
        lines.append(format_event_line(event))
    return "\n".join(lines)
