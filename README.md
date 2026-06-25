# 💰 Kakei-Bot

**Discord で家計簿を自動化する、AI 駆動の家計管理ボット**

> 「レシート写真を送るだけで、面倒な家計簿入力が終わる。」

---

## Why Open Source

個人事業主・フリーランス・自営業者の多くは、**家計管理ツールを使えていません。**

理由は簡単：既存の家計簿アプリは「一般向け」で、事業収支・確定申告対応・複数チャネル対応が不足しています。

Kakei-Bot は、こうした課題を **「オープンソースなら解決できる」** という前提で開発しています。

- 🔓 **透明性**：家計データの処理ロジックが見える（信頼できる）
- 🛠️ **拡張性**：ユーザーが必要に応じてカテゴリ・ルール・出力形式をカスタマイズできる
- 🤝 **コミュニティ駆動**：利用者からのフィードバック・プルリクエストで進化する

このプロジェクトは、「技術があれば、誰でも自分の家計管理を自動化できる世界」を目指しています。

---

## 🎯 対象ユーザー

Kakei-Bot は以下のユーザーを想定して設計されています：

- **個人事業主・フリーランス**：月間数十万の収支を自動追跡したい
- **小規模店舗経営者**：日々の売上・経費をリアルタイム管理したい
- **非技術者**：GUI・CLI 不要。Discord で家計簿がしたい
- **確定申告準備者**：年間ファイルを自動生成して、税理士に提出したい

---

## ✨ 主な機能

| 機能 | 使い方 | 用途 |
| --- | --- | --- |
| 📸 **レシート自動読み取り** | 画像を送るだけ | 領収書からの自動項目・金額抽出 |
| 💰 **収入記録** | 「給料15万入った」「売上30万」 | 日々の収入を自然言語で記録 |
| 📊 **収支確認** | 「今月見せて」「先月の食費いくら？」 | リアルタイムで月間・カテゴリ別集計 |
| ✏️ **修正・削除** | 「金額800円に直して」「最後の記録削除」 | 誤入力の即座修正 |
| ✍️ **手入力** | 「手入力」で直接入力 | レシートなしの記録に対応 |
| 📋 **確定申告用 Excel 出力** | 「まとめて」 | 年間ファイル生成。税理士・e-Tax 提出対応 |
| 💬 **カテゴリ相談** | 「これ何費？」でAIがアドバイス | 迷った時の自動分類提案 |

---

## 🚀 セットアップ

### 必要なもの

- **Discord Bot トークン**（[Discord Developer Portal](https://discord.com/developers/applications) で取得）
- **Anthropic API キー**（[console.anthropic.com](https://console.anthropic.com/) で取得）
- **Python 3.10 以上**

---

### 方法 A：かんたんセットアップ（Windows・非技術者向け）

1. このリポジトリを [ZIP でダウンロード](https://github.com/yousayrock/kakeibot/archive/refs/heads/main.zip) して展開する
2. `setup.bat` をダブルクリックする
3. 画面の指示に従って `.env` ファイルに 2 つのキーを入力する
4. 自動でBot が起動する

**次回以降の起動は** `start_bot.bat` **をダブルクリックするだけです。**

---

### 方法 B：手動セットアップ（技術者向け）

```bash
# 1. リポジトリをクローン
git clone https://github.com/yousayrock/kakeibot.git
cd kakeibot

# 2. ライブラリをインストール
pip install -r requirements.txt

# 3. 環境変数を設定
cp .env.example .env
# .env を開いて DISCORD_TOKEN と ANTHROPIC_API_KEY を入力

# 4. 起動
python bot.py
```

---

## ⚙️ 設定のカスタマイズ

`config.yml` を編集することで動作をカスタマイズできます。

### カテゴリの変更

```yaml
categories:
  - 食費
  - 交通費
  - 通信費
  - 消耗品費
  - 仕事経費
  - 光熱費
  - 医療費
  - 娯楽費
  - 外食費
  - 衣服費
  - 日用品
  - その他
```

自由に追加・削除できます。変更後はBotを再起動してください。

### 特定チャンネルのみで動作させる

```yaml
bot:
  allowed_channels:
    - 123456789012345678
    - 987654321098765432
```

空（`[]`）のままにすると全チャンネルで反応します。

---

## 📁 ファイル構成

```
kakeibot/
├── bot.py              # メインコード
├── config.yml          # カスタマイズ設定
├── requirements.txt    # 依存ライブラリ
├── .env.example        # 環境変数テンプレート
├── setup.bat           # かんたんセットアップ（Windows）
├── start_bot.bat       # 起動スクリプト（Windows）
├── data/               # 記録データ（自動生成・Gitに含まれません）
└── output/             # 確定申告用Excel出力先（自動生成）
```

---

## 🔒 セキュリティについて

- `.env` ファイルは `.gitignore` で除外されています。**絶対にGitHubにアップロードしないでください。**
- `data/` フォルダにはあなたの家計データが保存されます。こちらもGitに含まれません。

---

## 📈 ロードマップ

### v1.1（開発中）

- [ ] 複数ユーザーの同時管理対応（家族・チーム用）
- [ ] Google Sheets へのリアルタイム同期
- [ ] レシート画像の OCR 精度向上（複数言語対応）

### v1.2（計画中）

- [ ] Slack ワークスペース対応
- [ ] LINE Bot 版リリース
- [ ] 銀行口座連携（自動データインポート）

### v2.0（ビジョン）

- [ ] Web Dashboard（PC/モバイル での確認）
- [ ] マルチプラットフォーム対応（Discord / Slack / LINE / Teams）
- [ ] AI による支出最適化提案

ご希望の機能があれば、[Issues](https://github.com/yousayrock/kakeibot/issues) で提案してください。

---

## 🤝 コントリビューション

- **バグ報告**：[Issues](https://github.com/yousayrock/kakeibot/issues) で詳細を報告
- **機能リクエスト**：欲しい機能がある場合も Issues へ
- **プルリクエスト**：コード改良をいただく場合は、まず Issue で相談ください

---

## 📄 ライセンス

MIT License — 自由に使用・改造・配布できます。

---

## 🙏 使用技術

- [discord.py](https://github.com/Rapptz/discord.py) — Discord Bot フレームワーク
- [Anthropic Claude API](https://www.anthropic.com/) — AI による自然言語処理・レシート読み取り
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel ファイル生成

---

## 📞 サポート

- **使い方の質問**：[Discussions](https://github.com/yousayrock/kakeibot/discussions) へ
- **バグ報告**：[Issues](https://github.com/yousayrock/kakeibot/issues) へ

---

Made with ❤️ by [yousayrock](https://github.com/yousayrock)
