@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo  第一步：建我自己的库（数据只存在这台电脑）
echo ==========================================
echo.
where python >nul 2>&1
if errorlevel 1 goto NOPY
set "NAME="
set /p NAME=请输入你的名字（例如 张三），然后回车：
if "%NAME%"=="" goto NONAME
python scripts\init_person.py --name "%NAME%"
if errorlevel 1 goto FAILED
> "config\my_db.txt" echo data/persons/%NAME%/redpacket.db
echo.
echo 已记住你的库：data\persons\%NAME%\redpacket.db
echo 接下来双击 ②导入我的表格并核对.bat
echo.
pause
exit /b 0

:NOPY
echo 这台电脑还没装 Python（只需要装一次）。
echo 打开 https://www.python.org/downloads/ 下载，安装时务必勾选
echo   Add python.exe to PATH
echo 装完再双击本文件。
echo.
pause
exit /b 1

:NONAME
echo 名字不能为空，请重新双击本文件并输入名字。
pause
exit /b 1

:FAILED
echo 建库失败，请看上面的提示；常见原因是压缩包没解压、或 Python 没装好。
pause
exit /b 1
