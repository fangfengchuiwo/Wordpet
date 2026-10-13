@echo off
cd /d %~dp0
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 word_pet.py
) else (
    python word_pet.py
)
if errorlevel 1 (
    echo.
    echo [启动失败] 请确认已安装 Python 3，并执行: py -3 -m pip install -r requirements.txt
    pause
)
