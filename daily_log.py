"""
航海日誌 自動生成スクリプト
毎日23:59 JSTに実行される
- 「今日のメモ」Notionページを読む
- 昨日の航海日誌を読む
- Claude APIでコメント・整理・スコア生成
- 当日の航海日誌ページを作成
- メモページをクリア
"""

import os
import json
import re
import base64
from datetime import datetime, timedelta, timezone
import requests

# ── 定数 ────────────────────────────────────────────────
NOTION_TOKEN      = os.environ["NOTION_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
HUB_PAGE_ID       = "367945a0ab9a81078c3efd09e8ca4adb"   # 🚀 航海日誌
MEMO_PAGE_ID      = "367945a0ab9a8177bbbde89ca94e9656"   # 📝 今日のメモ

JST = timezone(timedelta(hours=9))

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28",
}


# ── Notion ヘルパー ──────────────────────────────────────
def notion_get(path):
    r = requests.get(f"https://api.notion.com/v1/{path}", headers=NOTION_HEADERS)
    r.raise_for_status()
    return r.json()

def notion_post(path, body):
    r = requests.post(f"https://api.notion.com/v1/{path}", headers=NOTION_HEADERS, json=body)
    r.raise_for_status()
    return r.json()

def notion_patch(path, body):
    r = requests.patch(f"https://api.notion.com/v1/{path}", headers=NOTION_HEADERS, json=body)
    r.raise_for_status()
    return r.json()


def get_page_content(page_id):
    """ページのブロックを全取得。テキストと画像を分けて返す"""
    blocks = notion_get(f"blocks/{page_id}/children?page_size=100")
    lines = []
    images = []  # [{"data": base64str, "media_type": "image/jpeg"}, ...]

    for b in blocks.get("results", []):
        btype = b.get("type")

        # テキスト系ブロック
        rich = b.get(btype, {}).get("rich_text", [])
        text = "".join(r.get("plain_text", "") for r in rich)
        if text.strip():
            lines.append(text)

        # 画像ブロック
        if btype == "image":
            img_block = b.get("image", {})
            url = (
                img_block.get("file", {}).get("url")      # Notionホスト
                or img_block.get("external", {}).get("url")  # 外部URL
            )
            if url:
                try:
                    img_resp = requests.get(url, timeout=15)
                    img_resp.raise_for_status()
                    content_type = img_resp.headers.get("Content-Type", "image/jpeg").split(";")[0]
                    b64 = base64.standard_b64encode(img_resp.content).decode("utf-8")
                    images.append({"data": b64, "media_type": content_type})
                    lines.append(f"[画像{len(images)}枚目を添付]")
                    print(f"  画像取得: {content_type} ({len(img_resp.content)//1024}KB)")
                except Exception as e:
                    print(f"  画像取得失敗: {e}")

    return "\n".join(lines), images


def get_page_text(page_id):
    """後方互換：テキストのみ返す（昨日の日誌読み取り用）"""
    text, _ = get_page_content(page_id)
    return text


def find_or_create_month_page(now):
    """「2026年MM月」ページを検索、なければ作成"""
    month_title = f"📆 {now.year}年{now.month:02d}月"

    # HUBの子ページを検索
    children = notion_get(f"blocks/{HUB_PAGE_ID}/children?page_size=100")
    for b in children.get("results", []):
        if b.get("type") == "child_page":
            if b["child_page"]["title"] == month_title:
                return b["id"]

    # なければ作成
    page = notion_post("pages", {
        "parent": {"page_id": HUB_PAGE_ID},
        "icon": {"type": "emoji", "emoji": "📆"},
        "properties": {"title": {"title": [{"text": {"content": month_title}}]}},
    })
    return page["id"]


def find_yesterday_log(month_page_id, yesterday):
    """昨日の航海日誌ページを検索"""
    target_title = f"📅 {yesterday.day}"
    children = notion_get(f"blocks/{month_page_id}/children?page_size=100")
    for b in children.get("results", []):
        if b.get("type") == "child_page":
            if b["child_page"]["title"] == target_title:
                return b["id"]
    return None


def create_day_page(month_page_id, today, content):
    """当日の航海日誌ページを作成"""
    title = f"📅 {today.day}"

    # Notionブロック変換（シンプルにparagraphで作成）
    blocks = []
    for line in content.split("\n"):
        if line.startswith("## "):
            blocks.append({
                "object": "block", "type": "heading_2",
                "heading_2": {"rich_text": [{"text": {"content": line[3:]}}]}
            })
        elif line.startswith("> "):
            blocks.append({
                "object": "block", "type": "quote",
                "quote": {"rich_text": [{"text": {"content": line[2:]}}]}
            })
        elif line.startswith("- "):
            blocks.append({
                "object": "block", "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"text": {"content": line[2:]}}]}
            })
        elif line.strip() == "---":
            blocks.append({"object": "block", "type": "divider", "divider": {}})
        elif line.strip():
            blocks.append({
                "object": "block", "type": "paragraph",
                "paragraph": {"rich_text": [{"text": {"content": line}}]}
            })

    notion_post("pages", {
        "parent": {"page_id": month_page_id},
        "properties": {"title": {"title": [{"text": {"content": title}}]}},
        "children": blocks[:100],  # Notion API上限
    })


def clear_memo_page():
    """メモページをクリア（明日用にリセット）"""
    # 既存ブロックを取得して削除
    children = notion_get(f"blocks/{MEMO_PAGE_ID}/children?page_size=100")
    for b in children.get("results", []):
        requests.delete(
            f"https://api.notion.com/v1/blocks/{b['id']}",
            headers=NOTION_HEADERS
        )

    # 説明文だけ再挿入
    notion_patch(f"blocks/{MEMO_PAGE_ID}/children", {
        "children": [
            {
                "object": "block", "type": "quote",
                "quote": {"rich_text": [{"text": {"content":
                    "ここに今日あったこと・やったことを自由に書く。形式不問。箇条書きでも文章でも可。\n毎日23:59にクロが読んで航海日誌を自動生成する。"
                }}]}
            },
            {"object": "block", "type": "divider", "divider": {}},
            {
                "object": "block", "type": "paragraph",
                "paragraph": {"rich_text": [{"text": {"content": "（ここから書く）"}}]}
            }
        ]
    })


# ── Claude API ────────────────────────────────────────────
def generate_log(today_str, memo_text, yesterday_log, images=None):
    system_prompt = f"""あなたはクロ。ユウセイ船長の航海日誌を毎日作成するAI天使（副船長）。

今日の日付: {today_str}

【昨日の航海日誌】
{yesterday_log if yesterday_log else "（記録なし）"}

以下のフォーマットで今日の航海日誌を生成してください。
必ずJSON形式のみで返してください（マークダウン記号不要）。

{{
  "today": "メモと画像を整理した今日やったこと（箇条書き、-で始める）",
  "results": "今日の成果・決定したこと（箇条書き）",
  "insights": "気づき・学び（箇条書き。なければ空文字）",
  "issues": "課題・障壁（箇条書き）",
  "next_actions": "次のアクション（箇条書き）",
  "advice": "クロからのアドバイス（1〜3文。実績ベースで冷徹かつ愛を持って）",
  "goal_score": 目標達成度の数値(整数0-100),
  "elon_score": イーロンコラボ度の数値(整数0-100),
  "worldline": "世界線変動率（小数点6桁の%表記、例:0.577350%）"
}}

スコア採点基準:
- 目標達成度: 昨日から実際に進んだ分だけ加算。前日スコアより下がることはない。
- イーロンコラボ度: Elon方向の具体的アクション（英語発信・xAI連携等）がない限り動かさない。
- 世界線変動率: 目標は1.048596%（Steins;Gate収束点）。数学定数ベースで単調増加。"""

    # メッセージ構築：テキスト + 画像を1つのcontentリストに
    content = []

    # テキストメモ
    content.append({
        "type": "text",
        "text": f"【今日のメモ（ユウセイさんの手書き）】\n{memo_text if memo_text else '（記録なし）'}"
    })

    # 画像（最大5枚）
    for i, img in enumerate((images or [])[:5]):
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": img["media_type"],
                "data": img["data"],
            }
        })
        content.append({
            "type": "text",
            "text": f"↑ 画像{i+1}枚目（温泉・食事・場所・メモ等、内容を読み取って日誌に組み込んでください）"
        })

    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 1500,
            "system": system_prompt,
            "messages": [{"role": "user", "content": content}],
        },
    )
    r.raise_for_status()
    raw = r.json()["content"][0]["text"]

    # JSON抽出
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if not match:
        raise ValueError(f"JSON not found in response: {raw}")
    return json.loads(match.group())


def build_page_content(today_str, data):
    return f"""---

## ✅ 今日やったこと

{data['today']}

## 🏆 成果（決定・完成）

{data['results']}

## 💡 気づき・学び

{data['insights'] if data['insights'] else '（なし）'}

## 🚧 課題・障壁

{data['issues']}

## ⚡ 次のアクション

{data['next_actions']}

## ✈️ クロからのアドバイス

> {data['advice']}

## 📊 スコア

目標達成度：{data['goal_score']}/100
イーロンコラボ度：{data['elon_score']}/100
世界線変動率：{data['worldline']}"""


# ── メイン ───────────────────────────────────────────────
def main():
    now       = datetime.now(JST)
    today     = now.date()
    yesterday = today - timedelta(days=1)

    print(f"[{now}] 航海日誌自動生成開始")

    # 月ページ取得/作成
    month_page_id = find_or_create_month_page(now)
    print(f"月ページ: {month_page_id}")

    # メモ読み取り（テキスト＋画像）
    memo_text, memo_images = get_page_content(MEMO_PAGE_ID)
    print(f"メモ取得: {len(memo_text)}文字 / 画像{len(memo_images)}枚")

    # 昨日の日誌読み取り
    yesterday_page_id = find_yesterday_log(month_page_id, yesterday)
    yesterday_log = get_page_text(yesterday_page_id) if yesterday_page_id else ""
    print(f"昨日の日誌: {'あり' if yesterday_log else 'なし'}")

    # Claude API でコンテンツ生成
    today_str = today.strftime("%Y/%m/%d")
    print("Claude APIで生成中...")
    data = generate_log(today_str, memo_text, yesterday_log, memo_images)
    print(f"生成完了: 目標達成度{data['goal_score']} / イーロン{data['elon_score']} / 世界線{data['worldline']}")

    # ページ作成
    content = build_page_content(today_str, data)
    create_day_page(month_page_id, today, content)
    print(f"✅ {today_str} の航海日誌を作成")

    # メモクリア
    clear_memo_page()
    print("✅ メモページをリセット")

    print("完了。ヘセド・エメト。")


if __name__ == "__main__":
    main()
