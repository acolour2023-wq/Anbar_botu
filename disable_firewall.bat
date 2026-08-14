@echo off
:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Administrator icazesi tesdiqlendi.
    echo Windows Firewall-a Port 5000 ucun qayda elave edilir...
    
    :: Remove old rule if exists, then add new one
    netsh advfirewall firewall delete rule name="Anbar_Sayim_Port_5000" >nul 2>&1
    netsh advfirewall firewall add rule name="Anbar_Sayim_Port_5000" dir=in action=allow protocol=TCP localport=5000
    
    echo.
    echo Port 5000 ugurla acildi! (Butun Firewall sondurulmedi, yalnız bu proqrama icaze verildi).
    echo Indi telefondan yeniden qosulmagi yoxlayin.
    pause
) else (
    echo ================================================================
    echo XETA: Zehmet olmasa bu fayla sag klikleyib 
    echo "Run as administrator" (Administrator olaraq baslat) secin.
    echo ================================================================
    pause
)
