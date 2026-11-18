@echo off
title ExamSync — Build Tool
color 0A

echo.
echo  ============================================
echo    ExamSync Desktop App — Build Script
echo  ============================================
echo.

echo  [Step 1/3] Installing required libraries...
echo  This may take a few minutes on first run.
echo.
pip install flask python-docx pandas openpyxl reportlab pywebview pyinstaller
echo.

echo  [Step 2/3] Building EXE...
echo  Please wait, this takes 2-4 minutes.
echo.
pyinstaller ExamSync.spec --clean --noconfirm
echo.

echo  [Step 3/3] Done!
echo.
echo  ============================================
echo   Your EXE is at:  dist\ExamSync.exe
echo  ============================================
echo.
echo  To distribute:
echo    - Copy  dist\ExamSync.exe  to any Windows PC
echo    - Double-click to run — no Python needed
echo    - Works fully OFFLINE
echo.
pause
