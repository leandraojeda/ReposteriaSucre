@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
 echo Primero ejecute instalar_windows.bat
 pause
 exit /b 1
)
echo Sistema: http://localhost:5000
echo Segunda computadora: http://IP-DE-ESTA-PC:5000
echo Consulte la IPv4 de su adaptador de red con ipconfig.
echo Mantenga esta ventana abierta mientras se utiliza el sistema.
.venv\Scripts\python.exe app.py
pause
