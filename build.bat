@echo off
chcp 65001 > nul

echo [1/2] Installing packages...
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Python not found. Please install Python and add to PATH.
    pause
    exit /b 1
)

echo [2/2] Building exe...
python -m PyInstaller --onefile --name insurance_lookup test3_reconstructed.py

echo.
echo Done! Check dist\insurance_lookup.exe
pause
