@echo off
cd /d "%~dp0"
echo Installing requirements (first time only)...
py -3.13 -m pip install -r requirements.txt
echo.
echo Starting the website. Open http://localhost:5000 in your browser.
echo Keep this window open while you use the site. Press Ctrl+C to stop.
start "" http://localhost:5000
py -3.13 app.py
pause
