@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 partes_aitex.py
) else (
    python partes_aitex.py
)
if errorlevel 1 (
    echo.
    echo No se ha podido iniciar. Necesitas Python 3.10 o superior.
    echo Descargalo de https://www.python.org/downloads/ marcando Add Python to PATH.
    pause
)
