#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Calendar 管理モジュール
サービスアカウント方式（Railway対応）
"""

import os
import json
import logging
from datetime import datetime, timedelta, timezone

logger = logging.getLogger("kakeibo")

GOOGLE_SA_JSON        = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
PERSONAL_CALENDAR_ID  = os.getenv("PERSONAL_CALENDAR_ID", "primary")
SHOP_CALENDAR_ID      = os.getenv("SHOP_CALENDAR_ID", "")

JST = timezone(timedelta(hours=9))

_service = None


def _available() -> bool:
    return bool(GOOGLE_SA_JSON)


def get_service():
    global _service
    if _service:
        return _service
    if not _available():
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON が未設定です")
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    sa = json.loads(GOOGLE_SA_JSON)
    creds = service_account.Credentials.from_service_account_info(
        sa, scopes=["https://www.googleapis.com/auth/calendar"]
    )
    _service = build("calendar", "v3", credentials=creds)
    return _service


def add_event(calendar_id: str, title: str, start_dt: datetime,
              end_dt: datetime = None, description: str = "",
              reminder_minutes: int = 30) -> dict:
    """予定を追加する"""
    if not _available():
        raise RuntimeError("カレンダー未設定")

    if end_dt is None:
        end_dt = start_dt + timedelta(hours=1)

    # JSTに変換
    if start_dt.tzinfo is None:
        start_dt = start_dt.replace(tzinfo=JST)
    if end_dt.tzinfo is None:
        end_dt = end_dt.replace(tzinfo=JST)

    event = {
        "summary": title,
        "description": description,
        "start": {"dateTime": start_dt.isoformat(), "timeZone": "Asia/Tokyo"},
        "end":   {"dateTime": end_dt.isoformat(),   "timeZone": "Asia/Tokyo"},
        "reminders": {
            "useDefault": False,
            "overrides": [{"method": "popup", "minutes": reminder_minutes}]
        },
    }

    service = get_service()
    result = service.events().insert(calendarId=calendar_id, body=event).execute()
    logger.info(f"🗓️ 予定追加: {title} {start_dt.strftime('%Y-%m-%d %H:%M')}")
    return result


def delete_event(calendar_id: str, event_id: str) -> bool:
    """予定を削除する"""
    if not _available():
        return False
    try:
        get_service().events().delete(calendarId=calendar_id, eventId=event_id).execute()
        return True
    except Exception as e:
        logger.error(f"カレンダー削除失敗: {e}")
        return False


def get_upcoming_events(calendar_ids: list, days: int = 7) -> list:
    """今後の予定を取得する"""
    if not _available():
        return []
    now = datetime.now(JST)
    time_min = now.isoformat()
    time_max = (now + timedelta(days=days)).isoformat()
    events = []
    service = get_service()
    for cal_id in calendar_ids:
        try:
            result = service.events().list(
                calendarId=cal_id,
                timeMin=time_min, timeMax=time_max,
                singleEvents=True, orderBy="startTime"
            ).execute()
            for e in result.get("items", []):
                start = e.get("start", {})
                dt_str = start.get("dateTime") or start.get("date", "")
                events.append({
                    "id": e.get("id"), "title": e.get("summary", "無題"),
                    "start": dt_str, "calendar_id": cal_id,
                    "description": e.get("description", "")
                })
        except Exception as ex:
            logger.warning(f"カレンダー取得失敗 {cal_id}: {ex}")
    events.sort(key=lambda x: x["start"])
    return events


def format_event(event: dict) -> str:
    """予定を表示用テキストに変換"""
    try:
        dt = datetime.fromisoformat(event["start"])
        if dt.tzinfo:
            dt = dt.astimezone(JST)
        time_str = dt.strftime("%m/%d %H:%M")
    except Exception:
        time_str = event["start"]
    return f"🗓️ {time_str}　{event['title']}"


def format_events_list(events: list, calendar_name: str = "予定", days: int = 7) -> str:
    """予定リストを表示用テキストに変換"""
    lines = [f"🗓️ **{calendar_name}（今後{days}日間）**\n"]
    if not events:
        lines.append("予定はありません。")
    else:
        for e in events:
            lines.append(format_event(e))
    return "\n".join(lines)
