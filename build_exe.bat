@echo off
echo PyInstaller yoxlanilir...
python -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller tapilmadi. Qurulur...
    python -m pip install pyinstaller openpyxl
)
echo.
echo .EXE fayli hazirlanir, zehmet olmasa gozleyin...
python -m PyInstaller --onefile --noconsole --name "Anbar_Sayim_Botu" --clean inventory_app.py
echo.
if %errorlevel% equ 0 (
    echo Ugurlu! .EXE fayli "dist" qovlugunda yaradildi.
) else (
    echo Xeta bas verdi!
)
pause
