# 💰 家計簿Bot

Discord上でレシート写真を送るだけで家計簿を自動管理するBotです。  
AIがレシートを読み取り、カテゴリ分けから確定申告用Excelの出力まで対応します。

## ✨ 主な機能

| 機能 | 使い方 |
|---|---|
| 📸 レシート自動読み取り | 画像を送るだけ |
| 💰 収入記録 | 「給料15万入った」 |
| 📊 収支確認 | 「今月見せて」「先月の食費いくら？」 |
| ✏️ 修正・削除 | 記録後に「金額800円に直して」など |
| ✍️ 手入力 | 「手入力」で直接入力 |
| 📋 確定申告用Excel出力 | 「まとめて」で年間ファイルを生成 |
| 💬 カテゴリ相談 | 「これ何費？」でAIがアドバイス |

---

## 🚀 セットアップ

### 必要なもの

- **Discord Bot トークン** （[Discord Developer Portal](https://discord.com/developers/applications) で取得）
- **Anthropic API キー** （[console.anthropic.com](https://console.anthropic.com/) で取得）
- **Python 3.10 以上**

---

### 方法A：かんたんセットアップ（Windows・非技術者向け）

1. このリポジトリを [ZIP でダウンロード](../../archive/refs/heads/main.zip) して展開する
2. `setup.bat` をダブルクリックする
3. 画面の指示に従って `.env` ファイルに2つのキーを入力する
4. 自動でBotが起動する

次回以降の起動は `start_bot.bat` をダブルクリックするだけです。

---

### 方法B：手動セットアップ（技術者向け）

```bash
# 1. リポジトリをクローン
git clone https://github.com/your-username/kakeibo-bot.git
cd kakeibo-bot

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
  - その他        # ← 末尾は「その他」推奨
```

自由に追加・削除できます。変更後はBotを再起動してください。

### 特定チャンネルのみで動作させる

```yaml
bot:
  allowed_channels:
    - 123456789012345678   # チャンネルID（右クリック→IDをコピー）
    - 987654321098765432
```

空（`[]`）のままにすると全チャンネルで反応します。

---

## 📁 ファイル構成

```
kakeibo-bot/
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

## 📄 ライセンス

MIT License

---

## 🙏 使用技術

- [discord.py](https://github.com/Rapptz/discord.py)
- [Anthropic Claude API](https://www.anthropic.com/)
- [openpyxl](https://openpyxl.readthedocs.io/)
