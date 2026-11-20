@echo off
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto error
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto error
echo Instalacion completa. Ejecute iniciar_windows.bat en la computadora principal.
pause
exit /b 0
:error
echo No se pudo instalar. Verifique Python 3.11 o superior e internet.
pause
exit /b 1
