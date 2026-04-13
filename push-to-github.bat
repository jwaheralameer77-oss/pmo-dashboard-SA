@echo off
cd /d "C:\Users\ameer\.verdent\verdent-projects\new-project"
echo ==========================================
echo    رفع PMO Dashboard على GitHub
echo ==========================================
echo.

"C:\Program Files\Git\bin\git.exe" remote add origin https://github.com/jwaheralameer77-oss/pmo-dashboard-SA.git 2>nul

echo جاري الرفع...
"C:\Program Files\Git\bin\git.exe" push -u origin main

echo.
echo ==========================================
echo اضغط أي مفتاح للخروج
pause >nul
