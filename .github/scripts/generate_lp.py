import os
import anthropic

with open("README.md", "r", encoding="utf-8") as f:
    readme = f.read()

with open("docs/index.html", "r", encoding="utf-8") as f:
    current_lp = f.read()

DESIGN_SPEC = """
【絶対に変えないこと】
- ヒーロー固定コピー: 「家の、基本OS。」
- 背景: #080c18（ダークネイビー）× アクセント: #c9963a（ゴールド）
- フォント: Noto Serif JP（見出し）+ Noto Sans JP（本文）+ Bebas Neue（英字）
- hero-title: font-size: clamp(2.4rem, 10vw, 7.5rem) / letter-spacing: clamp(0.08em, 1.5vw, 0.35em)
- 星空アニメーション（canvas + requestAnimationFrame）
- スクロールフェードイン（IntersectionObserver + .reveal クラス）
- フッター: © 2026 寳家 / 未来ガジェット研究所 + ヘセド・エメト
- 報告システムセクション（COMING NEXT バッジ付き）は必ず残す

【ページ構成（この順番で）】
1. NAV（家系Bot ロゴ + GitHubリンク）
2. HERO（「家の、基本OS。」+ キャッチ + ボタン2つ）
3. VISION（詩的なコピー）
4. FEATURES（READMEの機能一覧から生成）
5. HOW TO（セットアップ手順）
6. PRICING（SELF無料 / INSTALL ¥5,000 / CUSTOM 要相談）
7. COMING NEXT（報告システム + Discordモックアップ）
8. PLATFORM（動作環境テーブル）
9. CTA
10. FOOTER
"""

client = anthropic.Anthropic()

message = client.messages.create(
    model="claude-opus-4-7",
    max_tokens=8000,
    messages=[
        {
            "role": "user",
            "content": f"""家系BotのランディングページHTMLを生成してください。

## 最新のREADME（この内容をLPに反映する）
{readme}

## デザイン仕様
{DESIGN_SPEC}

## 現在のLP（デザインとアニメーションはこれを維持する）
{current_lp}

## 指示
- READMEの最新機能・説明をFEATURESセクションに反映してください
- デザイン・アニメーション・構成は現在のLPを維持してください
- 完全なHTMLファイルのみ出力してください（コードブロック記号不要）
- <!DOCTYPE html>から</html>まで出力してください"""
        }
    ]
)

html = message.content[0].text.strip()

if html.startswith("```"):
    html = "\n".join(html.split("\n")[1:])
if html.endswith("```"):
    html = html.rsplit("```", 1)[0]
html = html.strip()

with open("docs/index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("✅ LP生成完了")
