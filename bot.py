#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
家計簿Bot - レシート自動読み取り・家計簿管理
"""

import os
import re
import json
import shutil
import base64
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime

import yaml
import discord
from discord.ext import commands
from dotenv import load_dotenv
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import aiohttp

load_dotenv()

DISCORD_TOKEN = os.getenv("DISCORD_TOKEN")
ANTHROPIC_KEY = os.getenv("ANTHROPIC_API_KEY")
KAKEIBO_DIR   = Path(os.getenv("KAKEIBO_DIR", "data"))
OUTPUT_DIR    = Path(os.getenv("OUTPUT_DIR", "output"))

KAKEIBO_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ai_client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)


# ────────────────────────────────────────────
# 設定読み込み
# ────────────────────────────────────────────

def _default_config() -> dict:
    return {
        "bot": {
            "allowed_channels": [],
            "log_file": "kakeibo.log",
            "log_max_bytes": 200_000,
            "log_backup_count": 2,
        },
        "categories": [
            "食費", "交通費", "通信費", "消耗品費", "仕事経費",
            "光熱費", "医療費", "娯楽費", "外食費", "衣服費", "日用品", "その他",
        ],
        "ai": {
            "model": "claude-haiku-4-5-20251001",
            "receipt_max_tokens": 512,
            "intent_max_tokens": 400,
            "category_advice_max_tokens": 200,
        },
    }


def load_config() -> dict:
    config_path = Path("config.yml")
    if not config_path.exists():
        print("⚠️  config.yml が見つかりません。デフォルト設定で起動します。")
        return _default_config()
    with open(config_path, encoding="utf-8") as f:
        loaded = yaml.safe_load(f) or {}
    # デフォルト値とマージ（キーが欠けていても動作するよう）
    default = _default_config()
    for section, values in default.items():
        if section not in loaded:
            loaded[section] = values
        elif isinstance(values, dict):
            for k, v in values.items():
                loaded[section].setdefault(k, v)
    return loaded


CONFIG     = load_config()
CATEGORIES = CONFIG["categories"]
AI_MODEL   = CONFIG["ai"]["model"]
ALLOWED_CHANNELS: list[int] = [int(c) for c in CONFIG["bot"].get("allowed_channels", [])]


# ────────────────────────────────────────────
# ログ設定
# ────────────────────────────────────────────

logger = logging.getLogger("kakeibo")
logger.setLevel(logging.INFO)

_log_file      = CONFIG["bot"]["log_file"]
_log_max_bytes = CONFIG["bot"]["log_max_bytes"]
_log_backups   = CONFIG["bot"]["log_backup_count"]

_handler = RotatingFileHandler(_log_file, maxBytes=_log_max_bytes, backupCount=_log_backups, encoding="utf-8")
_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S"))
logger.addHandler(_handler)

_error_handler = logging.FileHandler("kakeibo_error.log", encoding="utf-8")
_error_handler.setLevel(logging.WARNING)
_error_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S"))
logger.addHandler(_error_handler)

_console = logging.StreamHandler()
_console.setFormatter(logging.Formatter("%(asctime)s %(message)s", "%H:%M:%S"))
logger.addHandler(_console)


# ────────────────────────────────────────────
# Discord Bot
# ────────────────────────────────────────────

intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix="!", intents=intents)

# 状態管理
pending_date:        dict[int, dict] = {}  # 日付確認待ち
pending_delete:      dict[int, dict] = {}  # 削除確認待ち
pending_confirm:     dict[int, dict] = {}  # 手入力確認待ち
pending_manual:      dict[int, dict] = {}  # 手入力受付中
pending_reply_edit:  dict[int, dict] = {}  # リプライ修正確認待ち
last_record_context: dict[int, dict] = {}  # 直近の記録コンテキスト


# ────────────────────────────────────────────
# ユーティリティ
# ────────────────────────────────────────────

def get_month_dir(year: int, month: int) -> Path:
    d = KAKEIBO_DIR / str(year) / f"{month:02d}_{month}月"
    d.mkdir(parents=True, exist_ok=True)
    (d / "receipts").mkdir(exist_ok=True)
    return d


def get_excel_path(year: int, month: int) -> Path:
    return get_month_dir(year, month) / "kakeibo.xlsx"


def load_records(year: int, month: int) -> list:
    p = get_month_dir(year, month) / "records.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return []


def save_records(year: int, month: int, records: list):
    p = get_month_dir(year, month) / "records.json"
    if p.exists():
        shutil.copy2(p, p.with_suffix(".bak"))
    p.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def make_record_id(date_str: str, filename: str) -> str:
    ts   = datetime.now().strftime("%H%M%S%f")[:10]
    safe = re.sub(r'[^\w.]', '_', filename)
    return f"{date_str.replace('-', '')}_{ts}_{safe}"


def make_receipt_filename(date_str: str, store_name: str, original: str,
                           receipt_time: str | None, sent_at: datetime) -> str:
    date_compact = date_str.replace("-", "")
    time_compact = receipt_time.replace(":", "")[:4] if receipt_time else sent_at.strftime("%H%M%S")
    ext          = original.rsplit(".", 1)[-1].lower() if "." in original else "jpg"
    safe_store   = re.sub(r'[\\/:*?"<>|]', '', store_name)[:20]
    return f"{date_compact}_{time_compact}_{safe_store}.{ext}"


def add_record(record: dict):
    year, month = record["year"], record["month"]
    records = load_records(year, month)
    records.append(record)
    save_records(year, month, records)
    update_excel(year, month, records)


def set_last_context(user_id: int, record: dict):
    records = load_records(record["year"], record["month"])
    idx = next((i for i, r in enumerate(records) if r.get("id") == record.get("id")), len(records) - 1)
    last_record_context[user_id] = {
        "record": record,
        "year":   record["year"],
        "month":  record["month"],
        "index":  idx,
    }


def find_record_by_id(year: int, month: int, record_id: str) -> tuple[int, dict] | tuple[None, None]:
    for i, r in enumerate(load_records(year, month)):
        if r.get("id") == record_id:
            return i, r
    return None, None


def delete_record_by_index(year: int, month: int, index: int) -> dict:
    records = load_records(year, month)
    deleted = records.pop(index)
    save_records(year, month, records)
    update_excel(year, month, records)
    return deleted


def fix_year(year: int, now: datetime) -> tuple[int, bool]:
    if abs(year - now.year) > 1:
        return now.year, True
    return year, False


def parse_manual_entry(text: str, now: datetime) -> dict | None:
    date_match = (
        re.search(r'(\d{1,2})[/\-](\d{1,2})', text)
        or re.search(r'(\d{1,2})月(\d{1,2})日', text)
    )
    if not date_match:
        return None
    try:
        date_obj = datetime(now.year, int(date_match.group(1)), int(date_match.group(2)))
    except ValueError:
        return None

    amount_match = re.search(r'([0-9,]{3,})', text)
    if not amount_match:
        return None
    amount   = int(amount_match.group(1).replace(",", ""))
    category = next((c for c in CATEGORIES if c in text), "その他")

    cleaned = text
    cleaned = re.sub(r'\d{1,2}[/\-]\d{1,2}', '', cleaned)
    cleaned = re.sub(r'\d{1,2}月\d{1,2}日', '', cleaned)
    cleaned = re.sub(r'[0-9,]{3,}円?', '', cleaned)
    for c in CATEGORIES:
        cleaned = cleaned.replace(c, '')
    name = cleaned.strip() or "不明"

    return {
        "date":     date_obj.strftime("%Y-%m-%d"),
        "year":     date_obj.year,
        "month":    date_obj.month,
        "name":     name,
        "amount":   amount,
        "category": category,
    }


def update_excel(year: int, month: int, records: list):
    path = get_excel_path(year, month)
    wb   = openpyxl.Workbook()
    ws   = wb.active
    ws.title = "収支一覧"

    hfill = PatternFill("solid", fgColor="1F4E79")
    ifill = PatternFill("solid", fgColor="E2EFDA")
    efill = PatternFill("solid", fgColor="FCE4D6")
    hfont = Font(bold=True, color="FFFFFF")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )

    for col, h in enumerate(["日付", "種別", "店名・内容", "カテゴリ", "金額（円）", "備考"], 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = hfill; c.font = hfont; c.border = border
        c.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 20

    income_total = expense_total = 0
    for row_idx, r in enumerate(sorted(records, key=lambda x: x.get("date", "")), 2):
        is_income = r.get("type") == "収入"
        amount    = r.get("amount", 0)
        if is_income:
            income_total += amount
        else:
            expense_total += amount
        for col, v in enumerate([r.get("date",""), r.get("type",""), r.get("name",""),
                                  r.get("category",""), amount, r.get("note","")], 1):
            c = ws.cell(row=row_idx, column=col, value=v)
            c.fill = ifill if is_income else efill
            c.border = border
            if col == 5:
                c.alignment    = Alignment(horizontal="right")
                c.number_format = '#,##0'

    last = len(records) + 2
    for row_offset, label, val, color in [
        (0, "【収入合計】", income_total,            "2E75B6"),
        (1, "【支出合計】", expense_total,           "C00000"),
        (2, "【収支差額】", income_total-expense_total, "000000"),
    ]:
        ws.cell(row=last+row_offset, column=3, value=label).font = Font(bold=True)
        c = ws.cell(row=last+row_offset, column=5, value=val)
        c.font = Font(bold=True, color=color)
        c.number_format = '#,##0'

    wb.save(path)


def build_summary(year: int, month: int, records: list, max_items: int = 20) -> str:
    income  = sum(r["amount"] for r in records if r.get("type") == "収入")
    expense = sum(r["amount"] for r in records if r.get("type") == "支出")
    lines   = [
        f"📊 {year}年{month}月の収支\n",
        f"💰 収入合計：¥{income:,}",
        f"💸 支出合計：¥{expense:,}",
        f"📈 収支差額：¥{income - expense:,}\n",
        "--- 明細 ---",
    ]
    for r in sorted(records, key=lambda x: x.get("date", ""))[-max_items:]:
        emoji = "💰" if r["type"] == "収入" else "💸"
        lines.append(f"{emoji} {r['date']} {r['name']} ¥{r['amount']:,} [{r['category']}]")
    if len(records) > max_items:
        lines.append(f"...他{len(records)-max_items}件（Excelで全件確認できます）")
    return "\n".join(lines)


def create_annual_summary(year: int) -> Path:
    output_path = OUTPUT_DIR / f"{year}_確定申告_収支内訳書.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "収支内訳書"

    title_font  = Font(bold=True, size=14)
    hfill = PatternFill("solid", fgColor="1F4E79")
    hfont = Font(bold=True, color="FFFFFF")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )

    ws["A1"] = f"{year}年度 収支内訳書（白色申告用）"
    ws["A1"].font = title_font
    ws.merge_cells("A1:G1")

    for col, h in enumerate(["月", "収入合計", "支出合計", "収支差額"], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.fill = hfill; c.font = hfont; c.border = border
        c.alignment = Alignment(horizontal="center")

    annual_income = annual_expense = 0
    for month in range(1, 13):
        records = load_records(year, month)
        income  = sum(r["amount"] for r in records if r.get("type") == "収入")
        expense = sum(r["amount"] for r in records if r.get("type") == "支出")
        annual_income += income; annual_expense += expense
        row = month + 3
        for col, val in enumerate([f"{month}月", income, expense, income-expense], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.border = border
            if col > 1:
                c.number_format = '#,##0'

    tr = 16
    ws.cell(row=tr, column=1, value="年間合計").font = Font(bold=True)
    for col, val, color in [(2, annual_income, "2E75B6"), (3, annual_expense, "C00000"),
                             (4, annual_income-annual_expense, "000000")]:
        c = ws.cell(row=tr, column=col, value=val)
        c.font = Font(bold=True, color=color)
        c.number_format = '#,##0'; c.border = border

    ws2 = wb.create_sheet("カテゴリ別集計")
    ws2["A1"] = f"{year}年度 カテゴリ別支出集計"
    ws2["A1"].font = title_font
    for col, h in enumerate(["カテゴリ", "年間合計"], 1):
        c = ws2.cell(row=3, column=col, value=h)
        c.fill = hfill; c.font = hfont

    totals: dict[str, int] = {c: 0 for c in CATEGORIES}
    for month in range(1, 13):
        for r in load_records(year, month):
            if r.get("type") == "支出":
                totals[r.get("category", "その他")] = totals.get(r.get("category", "その他"), 0) + r.get("amount", 0)
    for i, (cat, total) in enumerate(totals.items(), 4):
        ws2.cell(row=i, column=1, value=cat).border  = border
        ws2.cell(row=i, column=2, value=total).border = border
        ws2.cell(row=i, column=2).number_format = '#,##0'

    ws2.column_dimensions["A"].width = 15
    ws2.column_dimensions["B"].width = 15
    wb.save(output_path)
    return output_path


# ────────────────────────────────────────────
# AI処理
# ────────────────────────────────────────────

async def analyze_receipt(image_bytes: bytes, media_type: str) -> dict | None:
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    prompt = f"""このレシート画像を解析して、以下のJSON形式で返してください。
日付が読み取れない場合は"unknown"にしてください。
カテゴリは必ず以下から選んでください：{', '.join(CATEGORIES)}

【重要】金額の読み取りルール：
- 「小計」「合計」「税込合計」「お買い上げ合計」が実際の支払金額です
- 「お預かり」「お預り」は除外してください
- 「お釣り」「おつり」「チェンジ」は除外してください
- ポイント・割引後の実際の支払額を使ってください

{{
  "date": "YYYY-MM-DD",
  "time": "HH:MM（24時間表記。読み取れない場合はnull）",
  "name": "店名または内容",
  "amount": 実際の支払金額（整数）,
  "category": "カテゴリ",
  "confidence": "high/medium/low",
  "note": "補足があれば"
}}

JSONのみ返してください。"""
    try:
        response = ai_client.messages.create(
            model=AI_MODEL,
            max_tokens=CONFIG["ai"]["receipt_max_tokens"],
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text",  "text": prompt},
            ]}],
        )
        m = re.search(r'\{.*\}', response.content[0].text.strip(), re.DOTALL)
        if m:
            return json.loads(m.group())
    except Exception as e:
        logger.error(f"AI解析エラー: {e}")
    return None


async def ask_ai_intent(text: str, now: datetime, last_ctx: dict | None) -> dict:
    """メッセージの意図をAIで解析して返す。このチャンネルは家計簿専用なので全メッセージを処理。"""
    ctx_text = "なし（まだ記録していない）"
    if last_ctx:
        r = last_ctx["record"]
        ctx_text = (
            f"日付：{r.get('date')}　店名：{r.get('name')}　"
            f"金額：{r.get('amount')}円　カテゴリ：{r.get('category')}　"
            f"備考：{r.get('note', '')}"
        )

    prev_year  = now.year if now.month > 1 else now.year - 1
    prev_month = now.month - 1 if now.month > 1 else 12

    prompt = f"""あなたは家計簿Botです。ユーザーのメッセージの意図を判定してJSONで返してください。
このチャンネルは家計簿専用なので、全てのメッセージが家計簿に関係すると想定してください。

今日：{now.strftime('%Y-%m-%d')}（{now.year}年{now.month}月）
先月：{prev_year}年{prev_month}月

【直近の記録】
{ctx_text}

【ユーザーのメッセージ】
{text}

【intent一覧と返すJSON】

"edit"   … 直近記録を修正・変更・直す・入れて・書いて・追加して（複数フィールド同時対応）
  → {{"intent":"edit","edits":[{{"field":"amount","value":"800"}},{{"field":"note","value":"接待"}}]}}
  ※ 1フィールドでも必ずeditsリスト形式で返すこと

"delete" … 直近記録を消す・削除・取り消す
  → {{"intent":"delete"}}

"show"   … 収支・明細・一覧を見たい・確認・教えて
  → {{"intent":"show","year":{now.year},"month":{now.month}}}

"category" … 特定カテゴリの合計・いくら
  → {{"intent":"category","category":"食費","year":{now.year},"month":{now.month}}}

"income" … 収入を記録したい
  → {{"intent":"income","amount":15000,"name":"給料"}}

"manual" … 手書き・手入力・写真なしで領収書を入力したい
  → {{"intent":"manual"}}

"summary" … 年間集計・確定申告・まとめ
  → {{"intent":"summary","year":{now.year}}}

"help"   … 使い方・ヘルプ
  → {{"intent":"help"}}

"category_advice" … 「これ何費？」「どのカテゴリ？」「経費になる？」などカテゴリの相談
  → {{"intent":"category_advice","description":"相談内容をそのまま"}}

"unknown" … 挨拶など全く関係ないメッセージ
  → {{"intent":"unknown"}}

JSONのみ返してください。"""

    try:
        response = ai_client.messages.create(
            model=AI_MODEL,
            max_tokens=CONFIG["ai"]["intent_max_tokens"],
            messages=[{"role": "user", "content": prompt}],
        )
        m = re.search(r'\{.*\}', response.content[0].text.strip(), re.DOTALL)
        if m:
            return json.loads(m.group())
    except Exception as e:
        logger.error(f"AI意図解析エラー: {e}")
    return {"intent": "unknown"}


async def ask_ai_category_advice(description: str, record: dict | None = None) -> str:
    """支出内容からカテゴリのアドバイスをAIが返す。"""
    cats = "、".join(CATEGORIES)
    record_info = ""
    if record:
        record_info = (
            f"\n【直近の記録】\n"
            f"店名：{record.get('name')}　金額：{record.get('amount')}円　"
            f"現在のカテゴリ：{record.get('category')}"
        )

    prompt = f"""あなたは家計簿の専門家です。以下の支出内容に最もふさわしいカテゴリをアドバイスしてください。
{record_info}

【ユーザーの質問・内容】
{description}

【選べるカテゴリ】
{cats}

以下の形式で答えてください：
- おすすめカテゴリ（1〜2個）とその理由を簡潔に
- 確定申告で経費にできるかどうかも一言添える
- 最後に「○○費でよいですか？」と確認する
- 100文字以内でコンパクトに答える"""

    try:
        response = ai_client.messages.create(
            model=AI_MODEL,
            max_tokens=CONFIG["ai"]["category_advice_max_tokens"],
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text.strip()
    except Exception as e:
        logger.error(f"カテゴリアドバイスエラー: {e}")
        return f"カテゴリの候補：{cats}\nどれが近いか教えてください！"


async def ask_ai_reply_edit(text: str, record: dict) -> dict | None:
    """リプライの修正指示をAIで解析する。"""
    record_summary = (
        f"日付：{record.get('date')}　店名：{record.get('name')}　"
        f"金額：{record.get('amount')}円　カテゴリ：{record.get('category')}　"
        f"備考：{record.get('note', '')}"
    )
    prompt = f"""以下の記録に対する修正指示を解析してJSONで返してください。

【現在の記録】
{record_summary}

【修正指示】
{text}

フィールド：amount（金額）/ name（店名）/ date（YYYY-MM-DD）/ category（{', '.join(CATEGORIES)}）/ note（備考・宛名・但し書き）

{{"field": "フィールド名", "value": "新しい値"}}

判断不能の場合は {{"error": "不明"}} を返してください。JSONのみ返してください。"""

    try:
        response = ai_client.messages.create(
            model=AI_MODEL,
            max_tokens=128,
            messages=[{"role": "user", "content": prompt}],
        )
        m = re.search(r'\{.*\}', response.content[0].text.strip(), re.DOTALL)
        if m:
            return json.loads(m.group())
    except Exception as e:
        logger.error(f"AI修正解析エラー: {e}")
    return None


# ────────────────────────────────────────────
# 修正値のバリデーション・変換（共通処理）
# ────────────────────────────────────────────

def validate_edit(field: str, raw_value: str, now: datetime) -> tuple[any, str, str | None]:
    """
    (new_value, display_value, error_message) を返す。
    error_message が None でなければ失敗。
    """
    field_labels = {
        "amount": "金額", "name": "店名", "date": "日付",
        "category": "カテゴリ", "note": "備考/宛名/但し書き",
    }
    label = field_labels.get(field, field)

    if field == "amount":
        try:
            v = int(str(raw_value).replace(",", "").replace("円", ""))
            return v, f"¥{v:,}", None
        except ValueError:
            return None, None, "❌ 金額の数値が読み取れませんでした。"

    elif field == "date":
        m = re.search(r'(\d{4}-\d{2}-\d{2})|(\d{1,2})[/\-月](\d{1,2})', str(raw_value))
        if not m:
            return None, None, "❌ 日付の形式が読み取れませんでした。例：4/20"
        try:
            if m.group(1):
                v = m.group(1)
            else:
                v = datetime(now.year, int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
            return v, v, None
        except ValueError:
            return None, None, "❌ 日付が正しくありません。"

    elif field == "category":
        cat = next((c for c in CATEGORIES if c in str(raw_value)), None)
        if not cat:
            return None, None, f"❌ カテゴリが見つかりません。\n使えるカテゴリ：{' / '.join(CATEGORIES)}"
        return cat, cat, None

    elif field in ("name", "note"):
        v = str(raw_value).strip()
        return v, v, None

    return None, None, f"❌ フィールド '{field}' は修正できません。"


def apply_edits(user_id: int, edits: list, year: int, month: int, index: int, now: datetime) -> tuple[list, list]:
    """複数フィールドを一括修正。日付変更時はフォルダ・ファイルも移動。
    (成功メッセージリスト, エラーメッセージリスト) を返す。
    """
    records = load_records(year, month)
    if index < 0 or index >= len(records):
        return [], ["❌ 記録が見つかりませんでした。"]

    field_labels = {"amount":"金額","name":"店名","date":"日付","category":"カテゴリ","note":"備考/宛名/但し書き"}
    ok_lines  = []
    err_lines = []
    changed_fields = set()

    # フィールドを一括更新
    for edit in edits:
        field     = edit.get("field", "")
        raw_value = str(edit.get("value", "")).strip()
        new_value, display_value, err = validate_edit(field, raw_value, now)
        if err:
            err_lines.append(err)
            continue
        old = records[index].get(field, "")
        records[index][field] = new_value
        changed_fields.add(field)
        ok_lines.append(f"{field_labels.get(field, field)}：{old} → {display_value}")
        logger.info(f"✏️ 修正: {field} {old}→{new_value}")

    if not ok_lines:
        return ok_lines, err_lines

    rec          = records[index]
    old_filename = rec.get("receipt_file")
    receipt_dir  = get_month_dir(year, month) / "receipts"

    # ── 日付が変わった場合：フォルダごと移動 ──
    if "date" in changed_fields:
        new_date_str = rec.get("date", "")
        try:
            new_date_obj = datetime.strptime(new_date_str, "%Y-%m-%d")
        except ValueError:
            err_lines.append("❌ 日付の変換に失敗しました。")
            return ok_lines, err_lines

        new_year  = new_date_obj.year
        new_month = new_date_obj.month

        if new_year != year or new_month != month:
            # 古い月から記録を削除
            records.pop(index)
            save_records(year, month, records)
            update_excel(year, month, records)

            # 新しい月に記録を追加
            rec["year"]  = new_year
            rec["month"] = new_month
            new_records  = load_records(new_year, new_month)
            new_records.append(rec)
            new_idx = len(new_records) - 1
            save_records(new_year, new_month, new_records)
            update_excel(new_year, new_month, new_records)

            # レシートファイルを新フォルダに移動＆リネーム
            if old_filename:
                old_path = receipt_dir / old_filename
                if old_path.exists():
                    parts     = old_filename.split("_", 2)
                    date_part = new_date_str.replace("-", "")
                    time_part = parts[1] if len(parts) >= 2 else now.strftime("%H%M%S")
                    ext       = old_filename.rsplit(".", 1)[-1] if "." in old_filename else "jpg"
                    safe_name = re.sub(r'[\/:*?"<>|]', '', rec.get("name", "不明"))[:20]
                    new_filename = f"{date_part}_{time_part}_{safe_name}.{ext}"
                    new_receipt_dir = get_month_dir(new_year, new_month) / "receipts"
                    new_path = new_receipt_dir / new_filename
                    old_path.rename(new_path)
                    new_records[-1]["receipt_file"] = new_filename
                    save_records(new_year, new_month, new_records)
                    ok_lines.append(f"📁 {year}/{month:02d}月/receipts/{old_filename}")
                    ok_lines.append(f"　→ {new_year}/{new_month:02d}月/receipts/{new_filename}")
                    logger.info(f"📁 移動: {old_filename} → {new_year}/{new_month:02d}/{new_filename}")

            if user_id in last_record_context:
                last_record_context[user_id] = {
                    "record": new_records[new_idx],
                    "year":   new_year,
                    "month":  new_month,
                    "index":  new_idx,
                }
            return ok_lines, err_lines

        # 同じ年月内での日付変更（ファイル名だけリネーム）
        if old_filename and changed_fields & {"name", "date"}:
            old_path  = receipt_dir / old_filename
            if old_path.exists():
                parts        = old_filename.split("_", 2)
                date_part    = new_date_str.replace("-", "")
                time_part    = parts[1] if len(parts) >= 2 else now.strftime("%H%M%S")
                ext          = old_filename.rsplit(".", 1)[-1] if "." in old_filename else "jpg"
                safe_name    = re.sub(r'[\/:*?"<>|]', '', rec.get("name", "不明"))[:20]
                new_filename = f"{date_part}_{time_part}_{safe_name}.{ext}"
                new_path     = receipt_dir / new_filename
                old_path.rename(new_path)
                records[index]["receipt_file"] = new_filename
                ok_lines.append(f"📁 {old_filename} → {new_filename}")
                logger.info(f"📁 リネーム: {old_filename} → {new_filename}")

    # ── 店名だけ変わった場合：ファイル名リネーム ──
    elif "name" in changed_fields and old_filename:
        old_path = receipt_dir / old_filename
        if old_path.exists():
            parts        = old_filename.split("_", 2)
            date_part    = rec.get("date", "").replace("-", "")
            time_part    = parts[1] if len(parts) >= 2 else now.strftime("%H%M%S")
            ext          = old_filename.rsplit(".", 1)[-1] if "." in old_filename else "jpg"
            safe_name    = re.sub(r'[\/:*?"<>|]', '', rec.get("name", "不明"))[:20]
            new_filename = f"{date_part}_{time_part}_{safe_name}.{ext}"
            new_path     = receipt_dir / new_filename
            old_path.rename(new_path)
            records[index]["receipt_file"] = new_filename
            ok_lines.append(f"📁 {old_filename} → {new_filename}")
            logger.info(f"📁 リネーム: {old_filename} → {new_filename}")

    save_records(year, month, records)
    update_excel(year, month, records)

    if user_id in last_record_context:
        last_record_context[user_id]["record"] = records[index]

    return ok_lines, err_lines


def record_confirm_msg(record: dict) -> str:
    return (
        f"```\n"
        f"日付：{record['date']}\n"
        f"店名：{record['name']}\n"
        f"金額：¥{record['amount']:,}\n"
        f"カテゴリ：{record['category']}\n"
        f"備考：{record.get('note', '')}\n"
        f"```\n"
        f"↩️ このメッセージに返信するとその記録を操作できます\n"
        f"||[record_id:{record['id']}]||"
    )


# ────────────────────────────────────────────
# Discordイベント
# ────────────────────────────────────────────

@bot.event
async def on_ready():
    logger.info(f"✅ Bot起動: {bot.user}")
    print(f"✅ 家計簿Bot起動しました: {bot.user}")
    if ALLOWED_CHANNELS:
        ch_list = ", ".join(str(c) for c in ALLOWED_CHANNELS)
        print(f"📌 動作チャンネル制限あり: [{ch_list}]")
    else:
        print("📌 チャンネル制限なし（全チャンネルで動作）")


@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return

    # ── チャンネル制限 ──
    if ALLOWED_CHANNELS and message.channel.id not in ALLOWED_CHANNELS:
        return

    text    = message.content.strip()
    now     = datetime.now()
    user_id = message.author.id

    # ══════════════════════════════════════
    # ② 削除確認待ち
    # ══════════════════════════════════════
    if user_id in pending_delete:
        state = pending_delete.pop(user_id)
        if text in ["はい", "yes", "YES", "うん", "おk", "ok", "OK"]:
            deleted = delete_record_by_index(state["year"], state["month"], state["index"])
            receipt_file = deleted.get("receipt_file")
            if receipt_file:
                (get_month_dir(state["year"], state["month"]) / "receipts" / receipt_file).unlink(missing_ok=True)
            logger.info(f"🗑️ 削除: {deleted.get('name')} ¥{deleted.get('amount',0):,}")
            await message.channel.send(
                f"🗑️ 削除しました！\n```\n日付：{deleted.get('date')}\n店名：{deleted.get('name')}\n金額：¥{deleted.get('amount',0):,}\n```"
            )
        else:
            await message.channel.send("❌ キャンセルしました。")
        return

    # ══════════════════════════════════════
    # ③ 手入力確認待ち
    # ══════════════════════════════════════
    if user_id in pending_confirm:
        state = pending_confirm.pop(user_id)
        if text in ["はい", "yes", "YES", "うん", "おk", "ok", "OK"]:
            record        = state["record"]
            receipt_bytes = state.get("receipt_bytes")
            sent_at       = state.get("sent_at", now)
            if receipt_bytes:
                fname = make_receipt_filename(record["date"], record["name"],
                                              state.get("receipt_ext","receipt.jpg"), None, sent_at)
                record["receipt_file"] = fname
                (get_month_dir(record["year"], record["month"]) / "receipts" / fname).write_bytes(receipt_bytes)
            add_record(record)
            set_last_context(user_id, record)
            logger.info(f"✅ 手入力記録: {record['name']} ¥{record['amount']:,}")
            await message.channel.send(f"✅ 記録しました！\n{record_confirm_msg(record)}")
        elif text in ["いいえ", "no", "NO", "修正"]:
            pending_manual[user_id] = state
            await message.channel.send(
                "✏️ 入力し直してください。\n書式：`月/日 店名 金額 カテゴリ`\n「キャンセル」で中止"
            )
        else:
            pending_confirm[user_id] = state
            await message.channel.send("「はい」で保存 / 「いいえ」で修正 / 「キャンセル」で中止")
        return

    # ══════════════════════════════════════
    # ④ 手入力受付中
    # ══════════════════════════════════════
    if user_id in pending_manual:
        state = pending_manual.pop(user_id)
        if "キャンセル" in text or "やめ" in text:
            await message.channel.send("❌ キャンセルしました。")
            return
        parsed = parse_manual_entry(text, now)
        if not parsed:
            await message.channel.send(
                "❓ 読み取れませんでした。\n書式：`月/日 店名 金額 カテゴリ`\n例：`4/20 薬局 1200 医療費`\n「キャンセル」で中止"
            )
            pending_manual[user_id] = state
            return
        record = {
            "id": make_record_id(parsed["date"], "manual"),
            "year": parsed["year"], "month": parsed["month"],
            "date": parsed["date"], "type": "支出",
            "name": parsed["name"], "amount": parsed["amount"],
            "category": parsed["category"], "note": "手入力", "receipt_file": None,
        }
        pending_confirm[user_id] = {**state, "record": record}
        await message.channel.send(
            f"📋 内容を確認してください：\n"
            f"```\n日付：{record['date']}\n店名：{record['name']}\n"
            f"金額：¥{record['amount']:,}\nカテゴリ：{record['category']}\n```\n"
            f"「はい」で保存 / 「いいえ」で修正 / 「キャンセル」で中止"
        )
        return

    # ══════════════════════════════════════
    # ⑤ 日付確認待ち
    # ══════════════════════════════════════
    if user_id in pending_date:
        state = pending_date.pop(user_id)
        if "キャンセル" in text or "やめ" in text:
            await message.channel.send("❌ 記録をキャンセルしました。")
            return
        date_match = re.search(r'(?:(\d{1,2})月)?(\d{1,2})日', text)
        if date_match:
            try:
                parsed_date = datetime(now.year,
                                       int(date_match.group(1)) if date_match.group(1) else now.month,
                                       int(date_match.group(2)))
                state["record"]["date"]  = parsed_date.strftime("%Y-%m-%d")
                state["record"]["year"]  = parsed_date.year
                state["record"]["month"] = parsed_date.month
            except ValueError:
                await message.channel.send("❌ 日付が正しくありません。キャンセルしました。")
                return
        else:
            await message.channel.send("❓ 「4月20日」のように入力してください。「キャンセル」で中止")
            pending_date[user_id] = state
            return

        record        = state["record"]
        receipt_bytes = state.get("receipt_bytes")
        if receipt_bytes:
            fname = make_receipt_filename(record["date"], record["name"],
                                          state.get("receipt_ext","receipt.jpg"),
                                          state.get("receipt_time"), state.get("sent_at", now))
            record["receipt_file"] = fname
            (get_month_dir(record["year"], record["month"]) / "receipts" / fname).write_bytes(receipt_bytes)

        add_record(record)
        set_last_context(user_id, record)
        logger.info(f"✅ 記録（日付確定後）: {record['name']} ¥{record['amount']:,}")
        await message.channel.send(f"✅ 記録しました！\n{record_confirm_msg(record)}")
        return

    # ══════════════════════════════════════
    # ⑥ リプライ処理（特定記録への操作）
    # ══════════════════════════════════════
    if message.reference:
        try:
            ref_msg = await message.channel.fetch_message(message.reference.message_id)
        except Exception:
            pass
        else:
            id_match = re.search(r'\[record_id:([^\]]+)\]', ref_msg.content)
            if id_match:
                record_id = id_match.group(1)
                date_part = record_id[:8]
                try:
                    r_year, r_month = int(date_part[:4]), int(date_part[4:6])
                except ValueError:
                    r_year, r_month = now.year, now.month

                idx, target = find_record_by_id(r_year, r_month, record_id)
                if idx is None:
                    await message.channel.send("❌ 既に削除済みか、記録が見つかりませんでした。")
                    return

                # リプライ内容をAIで解析
                intent = await ask_ai_intent(text, now, {"record": target, "year": r_year, "month": r_month, "index": idx})

                if intent.get("intent") == "delete":
                    pending_delete[user_id] = {"year": r_year, "month": r_month, "index": idx, "record": target}
                    await message.channel.send(
                        f"⚠️ 本当に削除しますか？\n```\n日付：{target.get('date')}\n店名：{target.get('name')}\n"
                        f"金額：¥{target.get('amount',0):,}\nカテゴリ：{target.get('category')}\n```\n"
                        f"「はい」で削除 / 「いいえ」でキャンセル"
                    )
                    return

                if intent.get("intent") == "edit":
                    edits = intent.get("edits", [])
                    if not edits and intent.get("field"):
                        edits = [{"field": intent.get("field"), "value": intent.get("value","")}]
                    ok_lines, err_lines = apply_edits(user_id, edits, r_year, r_month, idx, now)
                    if ok_lines:
                        await message.channel.send(f"✅ 修正しました！\n```\n" + "\n".join(ok_lines) + "\n```")
                    for err in err_lines:
                        await message.channel.send(err)
                    return

                # その他のリプライ操作案内
                await message.channel.send(
                    "💬 この記録に対してできる操作：\n"
                    "　🗑️ 削除したい → 「消して」「取り消し」\n"
                    "　✏️ 修正したい → 「金額を800円に」「備考に交通費と入れて」など"
                )
                return

    # ══════════════════════════════════════
    # ⑦ レシート画像
    # ══════════════════════════════════════
    if message.attachments:
        for attachment in message.attachments:
            if not any(attachment.filename.lower().endswith(ext)
                       for ext in [".jpg", ".jpeg", ".png", ".webp", ".gif"]):
                continue

            await message.channel.send("📸 レシートを読み取り中...")
            logger.info(f"📸 レシート受信: {attachment.filename} user={user_id}")

            async with aiohttp.ClientSession() as session:
                async with session.get(attachment.url) as resp:
                    image_bytes = await resp.read()

            ext = attachment.filename.lower().rsplit(".", 1)[-1]
            media_type = {"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png",
                          "webp":"image/webp","gif":"image/gif"}.get(ext, "image/jpeg")

            result = await analyze_receipt(image_bytes, media_type)
            if not result:
                logger.warning(f"❌ AI解析失敗: {attachment.filename}")
                pending_manual[user_id] = {"receipt_bytes": image_bytes,
                                           "receipt_ext": attachment.filename, "sent_at": now}
                await message.channel.send(
                    "❌ 読み取りに失敗しました。手入力で登録できます。\n\n"
                    "書式：`月/日 店名 金額 カテゴリ`\n例：`4/20 薬局 1200 医療費`\n「キャンセル」で中止"
                )
                return

            # 日付不明
            if result.get("date") in ("unknown", "", None):
                user_comment = text.strip()
                record = {
                    "id": make_record_id("00000000", attachment.filename),
                    "year": now.year, "month": now.month, "date": "",
                    "type": "支出", "name": result.get("name","不明"),
                    "amount": result.get("amount",0), "category": result.get("category","その他"),
                    "note": user_comment if user_comment else result.get("note",""), "receipt_file": None,
                }
                pending_date[user_id] = {
                    "record": record, "receipt_bytes": image_bytes,
                    "receipt_time": result.get("time"), "receipt_ext": attachment.filename, "sent_at": now,
                }
                await message.channel.send(
                    f"📅 日付が読み取れませんでした。何月何日のレシートですか？（例：4月20日）\n「キャンセル」で中止\n\n"
                    f"```\n店名：{record['name']}\n金額：¥{record['amount']:,}\nカテゴリ：{record['category']}\n```"
                )
                return

            date_obj = datetime.strptime(result["date"], "%Y-%m-%d")
            original_year = date_obj.year
            fixed_year, was_fixed = fix_year(original_year, now)
            if was_fixed:
                logger.warning(f"⚠️ 年補正: {original_year}→{fixed_year} ファイル={attachment.filename}")
                date_obj       = date_obj.replace(year=fixed_year)
                result["date"] = date_obj.strftime("%Y-%m-%d")
                await message.channel.send(
                    f"⚠️ 年が `{original_year}年` と読み取られましたが **{fixed_year}年** に自動補正しました。\n"
                    f"違う場合はこのメッセージに返信して修正してください。"
                )

            if result.get("confidence") == "low":
                await message.channel.send("⚠️ 読み取り精度が低いです。内容を確認してください。")

            fname = make_receipt_filename(result["date"], result.get("name","不明"),
                                          attachment.filename, result.get("time"), now)
            user_comment = text.strip()
            note = user_comment if user_comment else result.get("note", "")
            record = {
                "id": make_record_id(result["date"], attachment.filename),
                "year": date_obj.year, "month": date_obj.month, "date": result["date"],
                "type": "支出", "name": result.get("name","不明"),
                "amount": result.get("amount",0), "category": result.get("category","その他"),
                "note": note, "receipt_file": fname,
            }
            add_record(record)
            set_last_context(user_id, record)
            (get_month_dir(date_obj.year, date_obj.month) / "receipts" / fname).write_bytes(image_bytes)
            logger.info(f"✅ 記録: {record['name']} ¥{record['amount']:,} [{record['category']}]")
            await message.channel.send(f"✅ 記録しました！\n{record_confirm_msg(record)}")

            # カテゴリが「その他」の場合は自動でアドバイス
            if record["category"] == "その他":
                advice = await ask_ai_category_advice(
                    f"{record['name']} ¥{record['amount']}円", record
                )
                await message.channel.send(
                    f"💬 カテゴリが「その他」になりました。\n\n{advice}"
                )
        return

    # ══════════════════════════════════════
    # ⑧ 全テキストメッセージ → AIで意図判定
    # ══════════════════════════════════════
    last_ctx = last_record_context.get(user_id)
    intent   = await ask_ai_intent(text, now, last_ctx)
    action   = intent.get("intent", "unknown")
    logger.info(f"🤖 intent={action} user={user_id} text={text[:40]}")

    # 直近記録の修正（複数フィールド対応・即反映）
    if action == "edit":
        if not last_ctx:
            await message.channel.send("📭 まだ記録がありません。レシートを送るか「手入力」で記録してください。")
            return
        edits = intent.get("edits", [])
        # 旧形式（field/value）との後方互換
        if not edits and intent.get("field"):
            edits = [{"field": intent.get("field"), "value": intent.get("value","")}]
        if not edits:
            await message.channel.send("❓ 修正内容が読み取れませんでした。")
            return
        ok_lines, err_lines = apply_edits(
            user_id, edits,
            last_ctx["year"], last_ctx["month"], last_ctx["index"], now
        )
        if ok_lines:
            await message.channel.send(f"✅ 修正しました！\n```\n" + "\n".join(ok_lines) + "\n```")
        for err in err_lines:
            await message.channel.send(err)
        return

    # 直近記録の削除
    if action == "delete":
        if not last_ctx:
            await message.channel.send("📭 まだ記録がありません。")
            return
        target = last_ctx["record"]
        pending_delete[user_id] = {
            "year": last_ctx["year"], "month": last_ctx["month"],
            "index": last_ctx["index"], "record": target,
        }
        await message.channel.send(
            f"⚠️ この記録を削除しますか？\n```\n日付：{target.get('date')}\n店名：{target.get('name')}\n"
            f"金額：¥{target.get('amount',0):,}\nカテゴリ：{target.get('category')}\n```\n"
            f"「はい」で削除 / 「いいえ」でキャンセル"
        )
        return

    # 収支表示
    if action == "show":
        t_year, t_month = intent.get("year", now.year), intent.get("month", now.month)
        records = load_records(t_year, t_month)
        if not records:
            await message.channel.send(f"📭 {t_year}年{t_month}月の記録はまだありません。")
            return
        await message.channel.send(build_summary(t_year, t_month, records))
        return

    # カテゴリ別集計
    if action == "category":
        cat     = intent.get("category", "")
        t_year  = intent.get("year",  now.year)
        t_month = intent.get("month", now.month)
        records = load_records(t_year, t_month)
        items   = [r for r in records if r.get("category") == cat and r.get("type") == "支出"]
        total   = sum(r["amount"] for r in items)
        lines   = [f"📂 {t_year}年{t_month}月の【{cat}】\n合計：¥{total:,}\n"]
        for r in sorted(items, key=lambda x: x.get("date","")):
            lines.append(f"💸 {r['date']} {r['name']} ¥{r['amount']:,}")
        await message.channel.send("\n".join(lines))
        return

    # 収入記録
    if action == "income":
        amount = int(str(intent.get("amount", 0)).replace(",",""))
        name   = str(intent.get("name", "収入")).strip() or "収入"
        record = {
            "id": make_record_id(now.strftime("%Y-%m-%d"), "income"),
            "year": now.year, "month": now.month, "date": now.strftime("%Y-%m-%d"),
            "type": "収入", "name": name, "amount": amount, "category": "収入", "note": "",
        }
        add_record(record)
        set_last_context(user_id, record)
        logger.info(f"💰 収入: {name} ¥{amount:,}")
        await message.channel.send(f"✅ 収入を記録しました！\n{record_confirm_msg(record)}")
        return

    # 手入力モード
    if action == "manual":
        pending_manual[user_id] = {"receipt_bytes": None, "receipt_ext": None, "sent_at": now}
        await message.channel.send(
            f"✏️ 手入力モードです。\n書式：`月/日 店名 金額 カテゴリ`\n"
            f"例：`4/20 薬局 1200 医療費`\n\n使えるカテゴリ：{' / '.join(CATEGORIES)}\n\n「キャンセル」で中止"
        )
        return

    # 年間集計
    if action == "summary":
        s_year = intent.get("year", now.year)
        await message.channel.send(f"📊 {s_year}年の確定申告用ファイルを作成中...")
        path = create_annual_summary(s_year)
        await message.channel.send(
            f"✅ 作成しました！\n`{path}`\n\n白色申告の収支内訳書フォーマットで保存しました。",
            file=discord.File(str(path)),
        )
        return

    # カテゴリ相談
    if action == "category_advice":
        description = intent.get("description", text)
        advice = await ask_ai_category_advice(description, last_ctx["record"] if last_ctx else None)
        await message.channel.send(f"💬 {advice}")
        return

    # ヘルプ
    if action == "help":
        await message.channel.send(
            "📖 **家計簿Botの使い方**\n\n"
            "📸 **レシート記録** → 画像を送るだけ\n"
            "💰 **収入記録** → 「給料15万入った」「収入5000円 副業」\n"
            "📊 **収支確認** → 「今月見せて」「先月の明細」「3月の収支」\n"
            "📂 **カテゴリ別** → 「今月の食費いくら？」「先月の交通費教えて」\n"
            "✏️ **修正** → 記録後に「備考に領収書ありって入れて」「金額800円に直して」\n"
            "🗑️ **削除** → 「消して」「取り消し」\n"
            "✍️ **手入力** → 「手入力」または写真失敗時に自動移行\n"
            "📋 **確定申告** → 「まとめて」「確定申告用ファイル作って」"
        )
        return

    # unknown：家計簿と無関係なメッセージはスルー
    await bot.process_commands(message)


if __name__ == "__main__":
    print("kakeibo Bot starting...")
    bot.run(DISCORD_TOKEN)
