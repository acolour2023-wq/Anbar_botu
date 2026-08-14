@echo off
chcp 65001 > nul
:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Administrator icazəsi təsdiqləndi.
    echo Windows Firewall-dan Port 5000 qaydası silinir...
    
    netsh advfirewall firewall delete rule name="Anbar_Sayim_Port_5000" >nul 2>&1
    
    echo.
    echo Port 5000 qaydası uğurla silindi!
    pause
) else (
    echo ================================================================
    echo XƏTA: Zəhmət olmasa bu fayla sağ klikləyib 
    echo "Run as administrator" (Administrator olaraq başlat) seçin.
    echo ================================================================
    pause
)
