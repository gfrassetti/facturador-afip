@echo off
cd /d "%~dp0"

REM 1) Entorno virtual (tiene openpyxl, selenium, webdriver-manager)
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0facturador_ui.py"
  exit /b 0
)
if exist "%~dp0.venv\Scripts\python.exe" (
  start "" "%~dp0.venv\Scripts\python.exe" "%~dp0facturador_ui.py"
  exit /b 0
)

REM 2) Python global (puede fallar si no instalaste pip install -r requirements.txt ahí)
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

echo No se encontro Python ni la carpeta .venv
echo Crear venv: py -3 -m venv .venv
echo Instalar: .venv\Scripts\pip install -r requirements.txt
pause
exit /b 1
