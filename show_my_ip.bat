@echo off
echo.
echo ================================================================
echo           KOMPUTERIN REAL IP UNVANINI TAPMAQ UCUN YARDIMCI
echo ================================================================
echo.
echo Hazirki sebeke kartlarinizin IP unvanlari:
echo ----------------------------------------------------------------
ipconfig | findstr /i "IPv4"
echo ----------------------------------------------------------------
echo.
echo TELIMAT:
echo * Yuxaridaki siyahida "Wireless LAN adapter Wi-Fi" ve ya "Wi-Fi" 
echo   yazilan bolmenin qarsisindaki IPv4 unvanini tapin.
echo   (Meselen: 192.168.100.XX ve ya 192.168.1.XX)
echo.
echo * Hemin unvani telefonda brauzere yazin. Meselen:
echo   http://[WIFI_IP_UNVANINIZ]:5000
echo.
pause
