@echo off
cd /d "%~dp0"
echo ====================================
echo  SnackTrack - Starting...
echo ====================================
echo.
echo App will open at: http://localhost:5000
echo Team members can access: http://YOUR-PC-IP:5000
echo.
echo Press Ctrl+C to stop the server.
echo.
python app.py
pause
