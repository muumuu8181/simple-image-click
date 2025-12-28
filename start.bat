@echo off
cd /d "%~dp0"

:: サーバー起動（バックグラウンド）
start /b python main.py

:: 少し待ってからブラウザ起動（小窓）
timeout /t 2 /nobreak >nul
start msedge --new-window --window-size=450,750 --window-position=50,50 http://localhost:8000

echo.
echo ========================================
echo   Simple Image Click 起動中...
echo   終了: このウィンドウを閉じる
echo ========================================
echo.

:: サーバーが動いている間待機
pause >nul
