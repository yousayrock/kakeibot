"""
航海日誌モジュール（バグ修正済み版）
- N1: GOOGLE_SA_JSON空チェック
- N2: Notion処理を全てtry-exceptで保護
- N3: Google APIライブラリを遅延importしてbot起動を守る
- N4: メモトリガーが金額だけの場合は家計簿に流す
- N5: Drive APIをasyncio.to_threadで非同期化
"""

import os
import io
import json
import asyncio
import requests
from datetime import datetime, timezone, timedelta

# ── 定数 ────────────────────────────────────────────────
IS_RAILWAY             = os.environ.get("RAILWAY_ENVIRONMENT") is not None
NOTION_TOKEN           = os.getenv("NOTION_TOKEN", "")
GOOGLE_SA_JSON         = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "1AP6eDCgOA_l1RiGxRPobCg9VROQEwxn1")
MEMO_PAGE_ID           = os.getenv("MEMO_PAGE_ID", "367945a0ab9a8177bbbde89ca94e9656")
MEMO_CHANNEL_NAME      = os.getenv("MEMO_CHANNEL_NAME", "今日のメモ")
MEMO_TRIGGERS          = ["メモ", "アイデア", "memo", "idea"]
MAX_IMAGE_BYTES        = 5 * 1024 * 1024  # 5MB

JST = timezone(timedelta(hours=9))

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28",
}


# ── Google Drive（N3: 遅延import）────────────────────────
def _drive_available() -> bool:
    return bool(IS_RAILWAY and GOOGLE_SA_JSON)

def _drive_service():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        sa_info = json.loads(GOOGLE_SA_JSON)
        creds = service_account.Credentials.from_service_account_info(
            sa_info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        raise RuntimeError(f"Drive初期化失敗: {e}")

def _get_or_create_month_folder(service, now: datetime) -> str:
    month_name = f"{now.year}年{now.month:02d}月"
    query = (
        f"name='{month_name}' and "
        f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and "
        f"mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    folder = service.files().create(
        body={
            "name": month_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [GOOGLE_DRIVE_FOLDER_ID],
        },
        fields="id"
    ).execute()
    return folder["id"]

def upload_to_drive(image_bytes: bytes, original_filename: str, mime_type: str = "image/jpeg") -> str:
    """同期関数。asyncio.to_thread()経由で呼ぶこと（N5対策）"""
    from googleapiclient.http import MediaIoBaseUpload
    service = _drive_service()
    now = datetime.now(JST)
    month_folder_id = _get_or_create_month_folder(service, now)
    ext = original_filename.rsplit(".", 1)[-1] if "." in original_filename else "jpg"
    new_filename = f"{now.strftime('%Y%m%d_%H%M%S')}.{ext}"
    media = MediaIoBaseUpload(io.BytesIO(image_bytes), mimetype=mime_type)
    file = service.files().create(
        body={"name": new_filename, "parents": [month_folder_id]},
        media_body=media, fields="id"
    ).execute()
    service.permissions().create(
        fileId=file["id"], body={"type": "anyone", "role": "reader"}
    ).execute()
    return f"https://drive.google.com/uc?id={file['id']}"


# ── Notion（N2: 全てtry-exceptで保護）───────────────────
def _notion_patch(path: str, body: dict) -> bool:
    if not NOTION_TOKEN:
        return False
    try:
        r = requests.patch(
            f"https://api.notion.com/v1/{path}",
            headers=NOTION_HEADERS, json=body, timeout=10
        )
        if r.status_code == 429:
            import time; time.sleep(1)
            r = requests.patch(
                f"https://api.notion.com/v1/{path}",
                headers=NOTION_HEADERS, json=body, timeout=10
            )
        r.raise_for_status()
        return True
    except Exception as e:
        print(f"[nisshi] Notion書き込み失敗: {e}")
        return False

def add_text_to_memo(text: str) -> bool:
    return _notion_patch(f"blocks/{MEMO_PAGE_ID}/children", {"children": [{
        "object": "block", "type": "bulleted_list_item",
        "bulleted_list_item": {"rich_text": [{"text": {"content": text[:2000]}}]}
    }]})

def add_multiline_to_memo(lines: list, label: str = "") -> bool:
    now = datetime.now(JST).strftime("%H:%M")
    header = f"💡 {label}（{now}）" if label else f"📝 メモ（{now}）"
    blocks = [
        {"object": "block", "type": "paragraph",
         "paragraph": {"rich_text": [{"text": {"content": header}}]}}
    ]
    for line in lines:
        if line.strip():
            blocks.append({
                "object": "block", "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"text": {"content": line.strip()[:2000]}}]}
            })
    return _notion_patch(f"blocks/{MEMO_PAGE_ID}/children", {"children": blocks})

def add_image_to_memo(url: str, caption: str = "") -> bool:
    children = [{
        "object": "block", "type": "image",
        "image": {"type": "external", "external": {"url": url}}
    }]
    if caption:
        children.append({
            "object": "block", "type": "paragraph",
            "paragraph": {"rich_text": [{"text": {"content": f"📷 {caption[:500]}"}}]}
        })
    return _notion_patch(f"blocks/{MEMO_PAGE_ID}/children", {"children": children})


# ── 画像処理（N5: asyncio.to_thread） ────────────────────
async def _handle_image(att, caption: str = "") -> bool:
    """Discord添付画像をDrive（またはURL）→ Notionに保存"""
    # サイズチェック
    if att.size > MAX_IMAGE_BYTES:
        print(f"[nisshi] 画像スキップ（{att.size//1024}KB > 5MB）")
        return False
    try:
        resp = requests.get(att.url, timeout=15)
        resp.raise_for_status()
        mime = (att.content_type or "image/jpeg").split(";")[0]

        if _drive_available():
            # N5: Driveアップロードを別スレッドで実行
            drive_url = await asyncio.to_thread(
                upload_to_drive, resp.content, att.filename, mime
            )
        else:
            drive_url = att.url  # ローカル: Discord URLをそのまま使用
            print("[nisshi] ローカル環境: Discord URLを使用（有効期限あり）")

        return add_image_to_memo(drive_url, caption)
    except Exception as e:
        print(f"[nisshi] 画像処理失敗: {e}")
        return False


# ── N4: メモトリガー判定（金額だけの場合は家計簿に流す）──
import re as _re
_AMOUNT_ONLY = _re.compile(r'^[メアidemo\s]*[0-9０-９,，]+[円\s]*$')

def parse_memo_trigger(text: str):
    """
    メモ/アイデアトリガーを検出。
    (label, [lines]) or (None, [])
    金額だけの場合（「メモ 500円」等）はNoneを返して家計簿に流す。
    """
    for trigger in MEMO_TRIGGERS:
        if text.startswith(trigger):
            rest = text[len(trigger):].strip()
            # N4: 数字だけなら家計簿に流す
            if _re.fullmatch(r'[0-9０-９,，円\s]+', rest):
                return None, []
            lines = [l for l in rest.split("\n") if l.strip()]
            return trigger, lines
    return None, []


# ── メインハンドラ ────────────────────────────────────────
async def handle(message) -> bool:
    """
    航海日誌の処理を行う。
    処理した場合はTrue（bot.pyの家計簿処理をスキップ）。
    処理しなかった場合はFalse。
    NOTION_TOKENが未設定の場合は常にFalseを返す。
    """
    if not NOTION_TOKEN:
        return False

    text   = message.content.strip()
    images = [a for a in message.attachments
              if a.content_type and a.content_type.startswith("image")]

    # ── #今日のメモ チャンネル ──────────────────────────
    if message.channel.name == MEMO_CHANNEL_NAME:
        if text:
            add_text_to_memo(text)
        for img in images:
            await _handle_image(img, text)
        await message.add_reaction("📝")
        return True

    # ── メモ/アイデア トリガー ────────────────────────
    label, lines = parse_memo_trigger(text)
    if label:
        if images:
            caption = "\n".join(lines)
            for img in images:
                await _handle_image(img, caption)
        if lines:
            add_multiline_to_memo(lines, label)
        await message.add_reaction("📝")
        return True

    return False
