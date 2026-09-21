@echo off
setlocal enabledelayedexpansion
title ILKIN SERVIS to Avto 326 Converter
cd /d "%~dp0"

set "LOGFILE=startup.log"
echo === Startup log === > "%LOGFILE%"
echo Date: %date% %time% >> "%LOGFILE%"
echo Dir:  %CD% >> "%LOGFILE%"

echo ============================================================
echo   ILKIN SERVIS  --^>  Avto 326   Converter
echo ============================================================
echo   Working directory: %CD%
echo.

REM ---------- 1. Find real Python ----------
echo [1/7] Looking for a real Python interpreter...

set "PY_CMD="

REM 1a. Try py launcher first (always real if installed)
py -3 --version >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=py -3"
    for /f "tokens=*" %%v in ('py -3 --version 2^>^&1') do set "PYVER=%%v"
    echo       Found launcher: !PY_CMD!  (!PYVER!)
    echo Python: !PYVER! via py -3 >> "%LOGFILE%"
    goto :python_ok
)

REM 1b. Try python.exe but validate it is not the Store stub
python --version >nul 2>nul
if not errorlevel 1 (
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do set "PYOUT=%%v"
    echo !PYOUT! | findstr /i /c:"Python 3" >nul
    if not errorlevel 1 (
        set "PY_CMD=python"
        set "PYVER=!PYOUT!"
        echo       Found interpreter: python  (!PYVER!)
        echo Python: !PYVER! via python >> "%LOGFILE%"
        goto :python_ok
    )
)

REM 1c. Nothing real found
echo.
echo [ERROR] No working Python 3 installation found.
echo.
echo What happened:
echo   - The "py" launcher is not installed, AND
echo   - the "python" command in PATH is the Microsoft Store stub
echo     (it opens the Store instead of running Python).
echo.
echo How to fix:
echo   1. Download Python 3.10+ from https://www.python.org/downloads/
echo   2. During install, CHECK "Add python.exe to PATH".
echo   3. Also install the "py" launcher (it is included by default).
echo   4. Close and reopen this window, then run start.bat again.
echo.
echo If the Store stub keeps hijacking "python":
echo   Settings - Apps - Advanced app settings - App execution aliases
echo   Turn OFF both "python.exe" and "python3.exe" aliases.
echo.
pause
exit /b 1

:python_ok

REM ---------- 2. Create / activate venv ----------
echo [2/7] Checking virtual environment...

if not exist ".venv\Scripts\python.exe" (
    echo       Creating .venv with !PY_CMD! ...
    !PY_CMD! -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create .venv
        echo -------- Log --------
        type "%LOGFILE%"
        pause
        exit /b 1
    )
    echo       .venv created.
) else (
    echo       .venv already exists.
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Failed to activate .venv
    pause
    exit /b 1
)

REM Confirm we are inside venv
where python | findstr /i ".venv" >nul
if errorlevel 1 (
    echo [WARN] venv activation did not change python path. Continuing anyway.
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set "VENV_PY=%%v"
echo       Active interpreter: !VENV_PY!
echo Venv interpreter: !VENV_PY! >> "%LOGFILE%"

REM ---------- 3. Dependencies ----------
echo [3/7] Checking dependencies...
python -c "import flask" >nul 2>nul
if errorlevel 1 (
    echo       Installing Flask ...
    python -m pip install --quiet --upgrade pip
    python -m pip install --quiet flask
    if errorlevel 1 (
        echo [ERROR] Failed to install Flask
        pause
        exit /b 1
    )
)
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo       Installing openpyxl ...
    python -m pip install --quiet openpyxl
    if errorlevel 1 (
        echo [ERROR] Failed to install openpyxl
        pause
        exit /b 1
    )
)
echo       OK

REM ---------- 4. Project files ----------
echo [4/7] Checking project files...
if not exist "app.py"                goto :no_app
if not exist "converter.py"          goto :no_conv
if not exist "templates\index.html"  goto :no_tpl
echo       OK

REM ---------- 5. Syntax ----------
echo [5/7] Checking Python syntax...
python -m py_compile app.py converter.py >> "%LOGFILE%" 2>&1
if errorlevel 1 goto :syntax_fail
echo       OK

REM ---------- 6. Imports ----------
echo [6/7] Checking module imports...
python -c "import app, converter" >> "%LOGFILE%" 2>&1
if errorlevel 1 goto :import_fail
echo       OK

REM ---------- 7. Port ----------
echo [7/7] Finding free port...
set "PORT="
for %%P in (5000 5050 8080 8000 5500) do (
    if not defined PORT (
        netstat -ano ^| findstr "LISTENING" ^| findstr ":%%P " >nul
        if errorlevel 1 set "PORT=%%P"
    )
)
if not defined PORT (
    echo [ERROR] No free port in 5000 5050 8080 8000 5500
    pause
    exit /b 1
)
echo       Using port %PORT%
set "FLASK_PORT=%PORT%"
set "FLASK_DEBUG=0"
echo Port: %PORT% >> "%LOGFILE%"

echo.
echo ============================================================
echo   Server: http://localhost:%PORT%
echo   Stop:   Ctrl + C
echo   Log:    %LOGFILE%
echo ============================================================
echo.

start "" cmd /c "timeout /t 2 >nul & start http://localhost:%PORT%"
python app.py 2>&1 | powershell -NoProfile -Command "$input | Tee-Object -FilePath '%LOGFILE%' -Append"

set "EXITCODE=%errorlevel%"
echo.
if not "%EXITCODE%"=="0" (
    echo [SERVER CRASHED] Exit code: %EXITCODE%
    echo -------- Log --------
    type "%LOGFILE%"
) else (
    echo [SERVER STOPPED]
)
echo.
pause
exit /b 0

:no_app
echo [ERROR] app.py not found in %CD%
echo Files in this folder:
dir /b
pause
exit /b 1

:no_conv
echo [ERROR] converter.py not found in %CD%
echo Files in this folder:
dir /b
pause
exit /b 1

:no_tpl
echo [ERROR] templates\index.html not found in %CD%
echo Files in this folder:
dir /b
pause
exit /b 1

:syntax_fail
echo [ERROR] Python syntax error in app.py or converter.py
echo -------- Log --------
type "%LOGFILE%"
pause
exit /b 1

:import_fail
echo [ERROR] Import failed. See log below.
echo -------- Log --------
type "%LOGFILE%"
pause
exit /b 1