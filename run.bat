@echo off
title Anbar Sayimi Yardimcisi
echo ==================================================
echo         ANBAR SAYIMI YARDIMCISHI
echo ==================================================
echo.
echo Lazim olan kitabxanalar yoxlanilir va yuklanir...
python -m pip install openpyxl
if %errorlevel% neq 0 (
    echo Xata: Python yuklanmayib va ya pip xatasi bas verdi!
    echo Python-un quruldugundan va PATH-a alava edildiyindan amin olun.
    pause
    exit /b
)
echo.
echo Proqram runs...
python inventory_app.py
if %errorlevel% neq 0 (
    echo Proqram gozlanilmadan dayandi.
)
pause
