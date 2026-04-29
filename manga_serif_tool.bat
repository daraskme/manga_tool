@echo off
setlocal
cd /d "%~dp0"

set VENV=.venv_manga_serif
set PYEXE=%VENV%\Scripts\python.exe
set PYWEXE=%VENV%\Scripts\pythonw.exe

if not exist "%PYEXE%" (
    echo [setup] Creating venv ^(%VENV%^)...
    where py >nul 2>&1
    if not errorlevel 1 (
        py -3 -m venv "%VENV%"
    ) else (
        python -m venv "%VENV%"
    )
    if errorlevel 1 (
        echo.
        echo ERROR: Python not found or venv creation failed.
        echo Install Python 3 from https://www.python.org/
        pause
        exit /b 1
    )
)

"%PYEXE%" -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo [setup] Installing PyQt6 ^(first run only^)...
    "%PYEXE%" -m pip install --upgrade pip
    "%PYEXE%" -m pip install PyQt6
    if errorlevel 1 (
        echo.
        echo ERROR: PyQt6 install failed.
        pause
        exit /b 1
    )
)

echo [run] Launching manga serif tool...
"%PYEXE%" "%~dp0manga_serif_tool.py"
set EXITCODE=%errorlevel%
if not "%EXITCODE%"=="0" (
    echo.
    echo ERROR: tool exited with code %EXITCODE%
    echo See log: "%~dp0manga_serif_tool.log"
    pause
)

endlocal
