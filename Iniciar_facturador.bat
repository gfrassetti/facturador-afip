@echo off
REM Preferimos pythonw / pyw: no dejan abierta la consola negra (solo la ventana de la app).
cd /d "%~dp0"

where pyw >nul 2>&1
if %errorlevel%==0 (
  start "" pyw -3 "%~dp0facturador_ui.py"
  exit /b 0
)
where pythonw >nul 2>&1
if %errorlevel%==0 (
  start "" pythonw "%~dp0facturador_ui.py"
  exit /b 0
)
where py >nul 2>&1
if %errorlevel%==0 (
  start "" py -3 "%~dp0facturador_ui.py"
  exit /b 0
)
where python >nul 2>&1
if %errorlevel%==0 (
  start "" python "%~dp0facturador_ui.py"
  exit /b 0
)

echo No se encontró Python. Instalá desde https://www.python.org y marcá "Add Python to PATH".
pause
exit /b 1
