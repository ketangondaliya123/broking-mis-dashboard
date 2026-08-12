@echo off
title Update Public Broking MIS Dashboard
cls
echo ======================================================================
echo           BROKING MIS DASHBOARD — PUBLIC UPDATE SCRIPT
echo ======================================================================
echo.
echo  1. Compressing latest dashboard data...
python _compress.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ERROR] Compression failed! Please make sure you generated the dashboard locally first.
    echo.
    pause
    exit /b 1
)

echo.
echo  2. Pushing latest dashboard to public website...
git add outputs/dashboard.html.gz
git commit -m "Update public dashboard data"
git push origin master

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [WARNING] Push failed. If prompted, please log into GitHub.
)

echo.
echo ======================================================================
echo  SUCCESS! Your public live website will update in 1 minute at:
echo  https://broking-mis-dashboard.onrender.com/dashboard
echo ======================================================================
echo.
pause
