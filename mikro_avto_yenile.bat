@echo off
title Mikro Anbar Qaligi - Avtomatik Sinxronizator
color 0A
chcp 65001 >nul

echo ========================================================================
echo        MIKRO ANBAR QALIQLARI - CANLI AVTOMATIK SINXRONIZATOR
echo ========================================================================
echo.
echo Bu pencere aciq qaldiqca:
echo 1. 'Son anbar qaligi.xlsx' faylini daimi izleyir.
echo 2. Mikrodan bu fayli yenileyen kimi bazani derhal yenileyir.
echo 3. Operatorlarin etdiyi sayimlar 100%% qorunur, yalniz qaliqlar yenilenir.
echo.
echo Dayandirmaq ucun bu pencireni baglaya ve ya Ctrl+C sixa bilersiniz.
echo ========================================================================
echo.

py -3.14 -u sync_anbar_qaligi.py
if %ERRORLEVEL% NEQ 0 (
    python -u sync_anbar_qaligi.py
)

pause
