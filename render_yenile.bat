@echo off
title Render ve GitHub Yenileme Komekcisi
color 0A
echo ========================================================================
echo               ANBAR BOTU - AVTOMATIK YENILEME KOMEKCISI
echo ========================================================================
echo.
echo Kodda edilen deyisiklikler GitHub ve Render-e gonderilir...
echo.

git add .

set /p msg="Deyisiklik haqqinda qisa qeyd yazin (ve ya Enter sixin): "
if "%msg%"=="" set msg=Saytin avtomatik yenilenmesi

git commit -m "%msg%"
git push origin main

echo.
echo ========================================================================
echo  ELA! Butun deyisiklikler GitHub-a gonderildi.
echo  Render.com avtomatik olaraq saytinizi 30 saniyeye yenileyecek!
echo.
echo  Admin Paneli: https://anbar-botu.onrender.com/admin
echo  Operator Paneli: https://anbar-botu.onrender.com
echo ========================================================================
echo.
pause
