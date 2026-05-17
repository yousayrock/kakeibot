@echo off
chcp 65001 > nul
echo.
echo ╔══════════════════════════════════════════╗
echo ║        家計簿Bot セットアップ            ║
echo ╚══════════════════════════════════════════╝
echo.

:: Python チェック
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Python が見つかりません。
    echo.
    echo Python 3.10 以上をインストールしてください：
    echo https://www.python.org/downloads/
    echo.
    echo インストール時に「Add Python to PATH」に必ずチェックを入れてください。
    pause
    exit /b 1
)

echo [OK] Python を確認しました。
python --version
echo.

:: .env チェック
if not exist ".env" (
    echo [セットアップ] .env ファイルを作成します...
    copy ".env.example" ".env" > nul
    echo.
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    echo  .env ファイルが作成されました。
    echo  メモ帳で開いて、以下の2つを入力してください：
    echo.
    echo  DISCORD_TOKEN=（Discordのトークン）
    echo  ANTHROPIC_API_KEY=（AnthropicのAPIキー）
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    echo.
    start notepad ".env"
    echo .env の編集が終わったら何かキーを押してください...
    pause > nul
) else (
    echo [OK] .env ファイルを確認しました。
)

echo.
echo [インストール] 必要なライブラリをインストール中...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [エラー] ライブラリのインストールに失敗しました。
    echo インターネット接続を確認してから再度実行してください。
    pause
    exit /b 1
)

echo.
echo ╔══════════════════════════════════════════╗
echo ║  セットアップ完了！Botを起動します...   ║
echo ╚══════════════════════════════════════════╝
echo.
echo Botを止めるには、この画面を閉じるか Ctrl+C を押してください。
echo.
python bot.py
pause
