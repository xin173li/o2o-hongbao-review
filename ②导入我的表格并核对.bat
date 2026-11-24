@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo  第二步：导入我自己的底表并自动核对
echo ==========================================
echo.
if not exist "config\my_db.txt" goto NODB
set "DB="
set /p DB=<"config\my_db.txt"
echo 我的库：%DB%
echo.
python scripts\import_data.py --db "%DB%"
if errorlevel 1 goto FAILED
echo.
python scripts\run_audit.py --db "%DB%"
if errorlevel 1 goto FAILED
echo.
echo 报告在 data\out\ 里，最新一份叫 最新核对报告.xlsx
echo 接下来双击 ③打开我的看板.bat
pause
exit /b 0

:NODB
echo 还没建库。请先双击 ①建我自己的库.bat。
pause
exit /b 1

:FAILED
echo 上面有报错信息；如果提示找不到表格，请确认 Excel 放进了 data\inbox 文件夹。
pause
exit /b 1
