@echo off
chcp 65001 >nul
echo ========================================
echo   価格計算システム 起動
echo ========================================
echo.

cd /d "%~dp0"

echo [1/2] 依存パッケージ確認中...
pip install flask oracledb >nul 2>&1

echo [2/2] サーバー起動中...
echo.
echo   ブラウザで以下にアクセスしてください:
echo   http://localhost:5001/
echo.
echo   停止するには Ctrl+C を押してください
echo ========================================
echo.

python price_app.py

pause
