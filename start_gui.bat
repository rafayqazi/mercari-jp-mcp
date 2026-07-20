@echo off
title Mercari JP Search GUI
cd /d "%~dp0"
echo ============================================================
echo  Mercari JP Search GUI
echo ============================================================
echo.
echo  Starting server...
echo  Open http://127.0.0.1:5000 in your browser
echo  Close this window to stop the server
echo.
echo ============================================================
python mercari_gui.py
pause
