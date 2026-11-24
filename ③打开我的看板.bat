@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "config\my_db.txt" goto NODB
set "DB="
set /p DB=<"config\my_db.txt"
echo 正在打开看板：http://127.0.0.1:8899/
start "O2O红包看板" /min cmd /c python scripts\serve_dashboard.py --db "%DB%" --port 8899
timeout /t 3 >nul
start "" http://127.0.0.1:8899/
echo 看板已启动，浏览器里如果没自动打开，就手动访问 http://127.0.0.1:8899/
echo 关掉黑色小窗口就停止看板；想看别人的数据不行，这个窗口只连你自己的库。
timeout /t 6 >nul
exit /b 0

:NODB
echo 还没建库。请先双击 ①建我自己的库.bat。
pause
exit /b 1
