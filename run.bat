@echo off
setlocal
title OCR Scanner Tool Launcher

:: ==========================================
:: בדיקת הרשאות אדמין (Administrator Check)
:: ==========================================
echo Checking for Administrator privileges...
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Success: Running as Administrator.
    goto :RunScript
) else (
    echo Requesting Administrator privileges...
    goto :GetAdmin
)

:GetAdmin
    :: הרצה מחדש של הסקריפט הזה באמצעות PowerShell עם הפקודה 'RunAs'
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b

:RunScript
    :: ==========================================
    :: שינוי תיקיית העבודה (חשוב מאוד!)
    :: ==========================================
    :: כאשר מריצים כאדמין, התיקייה משתנה ל-System32. הפקודה הבאה מחזירה אותנו לתיקיית הסקריפט.
    cd /d "%~dp0"

    echo Starting OCR Project...
    echo --------------------------------------
    
    :: הרצת הסקריפט
    :: אם יש לך סביבה וירטואלית (venv), צריך להפעיל אותה כאן קודם.
    :: אם לא, נשתמש בפייתון הגלובלי:
    python main.py

    :: ==========================================
    :: סיום
    :: ==========================================
    echo.
    echo The application has closed.
    :: ה-pause משאיר את החלון פתוח במקרה של שגיאה כדי שתוכל לראות מה קרה
    pause