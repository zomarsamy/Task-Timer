@echo off
setlocal
cd /d "%~dp0"
echo ================================================
echo             Dodo Task Timer
echo ================================================
echo.
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    ) else (
        echo Python was not found on this computer.
        echo Install Python 3.11 or newer from https://www.python.org/downloads/windows/
        echo Enable "Add Python to PATH" during installation, then run this file again.
        pause
        exit /b 1
    )
)
if not exist ".venv\Scripts\python.exe" (
    echo Setting up the app environment for the first time...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 goto :error
)
echo Checking required libraries...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :error
echo Starting Dodo Task Timer...
".venv\Scripts\python.exe" dodo_task_timer.py
if errorlevel 1 goto :error
exit /b 0
:error
echo.
echo Setup or startup failed. Check your internet connection and the error message above.
pause
exit /b 1
