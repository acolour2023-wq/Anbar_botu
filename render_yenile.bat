@echo off
chcp 65001 > nul
title Render ve GitHub Yenileme Komekcisi
color 0A
echo ========================================================================
echo               ANBAR BOTU - AVTOMATIK YENILEME KOMEKCISI
echo ========================================================================
echo.
echo Kodda etdiyiniz bütün dəyişikliklər GitHub-a və Render-ə göndərilir...
echo.

git add .

set /p msg="Dəyişiklik haqqında qısa qeyd yazın (və ya birbaşa Enter sıxın): "
if "%msg%"=="" set msg=Saytın avtomatik yenilənməsi

git commit -m "%msg%"
git push origin main

echo.
echo ========================================================================
echo  ƏLA! Bütün dəyişikliklər GitHub-a göndərildi.
echo  Render.com avtomatik olaraq saytınızı 30 saniyəyə yeniləyəcək!
echo.
echo  🌐 Admin Paneli: https://anbar-botu-ip22.onrender.com/admin
echo  📱 Operator Paneli: https://anbar-botu-ip22.onrender.com
echo ========================================================================
echo.
pause
