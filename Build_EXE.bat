@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    ) else (
        echo Install Python 3.11 or newer, then run this file again.
        pause
        exit /b 1
    )
)
if not exist ".venv\Scripts\python.exe" (
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 goto :error
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt pyinstaller
if errorlevel 1 goto :error
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name DodoTaskTimer --collect-all pynput --hidden-import openpyxl dodo_task_timer.py
if errorlevel 1 goto :error
echo.
echo Build complete:
echo %~dp0dist\DodoTaskTimer.exe
echo You can copy the EXE wherever you like. App data remains in the user data folder.
pause
exit /b 0
:error
echo Failed to build the executable. Review the error message above.
pause
exit /b 1
