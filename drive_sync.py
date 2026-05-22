"""
drive_sync.py
Railway再起動時にrecords.jsonとconfig.jsonが消えないよう
Google Driveに自動バックアップ・リストアするモジュール。
"""

import os
import io
import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

IS_RAILWAY             = os.environ.get("RAILWAY_ENVIRONMENT") is not None
GOOGLE_SA_JSON         = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "")

# Driveの「data」サブフォルダのIDをキャッシュ
_data_folder_id: str | None = None


def _available() -> bool:
    return bool(IS_RAILWAY and GOOGLE_SA_JSON and GOOGLE_DRIVE_FOLDER_ID)


def _service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    sa = json.loads(GOOGLE_SA_JSON)
    creds = service_account.Credentials.from_service_account_info(
        sa, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def _get_data_folder(service) -> str:
    """Driveのルートフォルダ内に"data"フォルダを取得or作成してIDを返す"""
    global _data_folder_id
    if _data_folder_id:
        return _data_folder_id

    q = (f"name='data' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    if files:
        _data_folder_id = files[0]["id"]
    else:
        folder = service.files().create(
            body={"name": "data",
                  "mimeType": "application/vnd.google-apps.folder",
                  "parents": [GOOGLE_DRIVE_FOLDER_ID]},
            fields="id"
        ).execute()
        _data_folder_id = folder["id"]

    return _data_folder_id


def _drive_key_to_name(drive_key: str) -> str:
    """drive_key（パス風の文字列）をファイル名に変換 例: '914.../2026/05/records' → '914..._2026_05_records.json'"""
    return drive_key.replace("/", "_") + ".json"


def upload(local_path: Path, drive_key: str) -> bool:
    """ローカルファイルをDriveにアップロード（上書き）"""
    if not _available() or not local_path.exists():
        return False
    try:
        from googleapiclient.http import MediaIoBaseUpload
        service = _service()
        data_folder = _get_data_folder(service)
        filename = _drive_key_to_name(drive_key)

        # 既存ファイルを検索
        q = f"name='{filename}' and '{data_folder}' in parents and trashed=false"
        existing = service.files().list(q=q, fields="files(id)").execute().get("files", [])

        content = local_path.read_bytes()
        media = MediaIoBaseUpload(io.BytesIO(content), mimetype="application/json")

        if existing:
            service.files().update(
                fileId=existing[0]["id"], media_body=media
            ).execute()
        else:
            service.files().create(
                body={"name": filename, "parents": [data_folder]},
                media_body=media, fields="id"
            ).execute()

        logger.info(f"[drive_sync] アップロード完了: {drive_key}")
        return True
    except Exception as e:
        logger.warning(f"[drive_sync] アップロード失敗: {drive_key} → {e}")
        return False


def download(drive_key: str, local_path: Path) -> bool:
    """DriveからローカルパスにJSONをダウンロード"""
    if not _available():
        return False
    try:
        service = _service()
        data_folder = _get_data_folder(service)
        filename = _drive_key_to_name(drive_key)

        q = f"name='{filename}' and '{data_folder}' in parents and trashed=false"
        existing = service.files().list(q=q, fields="files(id)").execute().get("files", [])
        if not existing:
            return False

        content = service.files().get_media(fileId=existing[0]["id"]).execute()
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(content)
        logger.info(f"[drive_sync] ダウンロード完了: {drive_key}")
        return True
    except Exception as e:
        logger.warning(f"[drive_sync] ダウンロード失敗: {drive_key} → {e}")
        return False


def restore_all(kakeibo_dir: Path) -> None:
    """起動時: Drive上の全dataファイルをローカルに復元"""
    if not _available():
        logger.info("[drive_sync] Drive未設定のため復元スキップ")
        return
    try:
        service = _service()
        data_folder = _get_data_folder(service)

        q = f"'{data_folder}' in parents and trashed=false"
        files = service.files().list(q=q, fields="files(id,name)").execute().get("files", [])

        for f in files:
            name = f["name"]  # 例: 914..._2026_05_records.json
            if not name.endswith(".json"):
                continue
            # ファイル名からローカルパスを復元
            drive_key = name[:-5]  # .json除去
            parts = drive_key.split("_")
            # user_id_year_month_records or user_id_year_month_config
            if len(parts) < 4:
                continue
            file_type = parts[-1]  # "records" or "config"
            month = parts[-2]
            year = parts[-3]
            user_id = "_".join(parts[:-3])

            local_dir = kakeibo_dir / user_id / year / f"{int(month):02d}_{int(month)}月"
            local_dir.mkdir(parents=True, exist_ok=True)
            local_path = local_dir / f"{file_type}.json"

            if not local_path.exists():
                content = service.files().get_media(fileId=f["id"]).execute()
                local_path.write_bytes(content)
                logger.info(f"[drive_sync] 復元: {local_path}")

        logger.info(f"[drive_sync] 起動時復元完了: {len(files)}ファイル")
    except Exception as e:
        logger.warning(f"[drive_sync] 起動時復元失敗: {e}")
