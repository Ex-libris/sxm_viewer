@echo off
setlocal enabledelayedexpansion

REM Installer wrapper for SXM Viewer. Runs python install.py with best available interpreter.
cd /d "%~dp0"

echo ========================================
echo SXM Viewer Installer
echo ========================================
echo.

set "PYTHON_EXE="

REM Priority 1: PYTHON environment variable
if defined PYTHON (
    if exist "%PYTHON%" (
        set "PYTHON_EXE=%PYTHON%"
        echo Using PYTHON env var: !PYTHON_EXE!
    )
)

REM Priority 2: Look for Python from python.org (with SSL support)
if not defined PYTHON_EXE (
    echo Searching for Python installations...
for %%V in (313 312 311 310 39) do (
        REM Check LocalAppData (standard python.org install for current user)
        if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
            set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python%%V\python.exe"
            echo Found Python 3.%%V in LocalAppData
            goto :found_python
        )
        REM Check ProgramData (all users install)
        if exist "%ProgramData%\Python%%V\python.exe" (
            set "PYTHON_EXE=%ProgramData%\Python%%V\python.exe"
            echo Found Python 3.%%V in ProgramData
            goto :found_python
        )
        REM Check Program Files
        if exist "%ProgramFiles%\Python%%V\python.exe" (
            set "PYTHON_EXE=%ProgramFiles%\Python%%V\python.exe"
            echo Found Python 3.%%V in Program Files
            goto :found_python
        )
        REM Check old-style C:\Python3X
        if exist "C:\Python%%V\python.exe" (
            set "PYTHON_EXE=C:\Python%%V\python.exe"
            echo Found Python 3.%%V in C:\Python%%V
            goto :found_python
        )
    )
)

REM Priority 3: py launcher (gets the best available Python)
if not defined PYTHON_EXE (
    py -3 --version >nul 2>&1
    if !errorlevel! equ 0 (
        REM Get the actual path from py launcher
        for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)"') do set "PY_PATH=%%i"
        if exist "!PY_PATH!" (
            set "PYTHON_EXE=!PY_PATH!"
            echo Found Python via py launcher: !PY_PATH!
            goto :found_python
        ) else (
            set "PYTHON_EXE=py -3"
            echo Found py launcher
            goto :found_python
        )
    )
)

REM Priority 4: Windows Store Python
if not defined PYTHON_EXE (
    if exist "%LocalAppData%\Microsoft\WindowsApps\python.exe" (
        REM Test if it's real or just a stub
        "%LocalAppData%\Microsoft\WindowsApps\python.exe" --version >nul 2>&1
        if !errorlevel! equ 0 (
            set "PYTHON_EXE=%LocalAppData%\Microsoft\WindowsApps\python.exe"
            echo Found Windows Store Python
            goto :found_python
        )
    )
)

REM Priority 5: python in PATH
if not defined PYTHON_EXE (
    python --version >nul 2>&1
    if !errorlevel! equ 0 (
        REM Get the actual path
        for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)"') do set "PY_PATH=%%i"
        if defined PY_PATH (
            set "PYTHON_EXE=!PY_PATH!"
            echo Found Python in PATH: !PY_PATH!
        ) else (
            set "PYTHON_EXE=python"
            echo Found python in PATH
        )
        goto :found_python
    )
)

REM Priority 6: Check for Anaconda/Miniconda as last resort
if not defined PYTHON_EXE (
    echo Checking for Anaconda/Miniconda...
    for %%P in (
        "%USERPROFILE%\anaconda3\python.exe"
        "%USERPROFILE%\miniconda3\python.exe"
        "%USERPROFILE%\miniconda\python.exe"
        "C:\ProgramData\Anaconda3\python.exe"
        "C:\ProgramData\Miniconda3\python.exe"
        "%LocalAppData%\Continuum\anaconda3\python.exe"
    ) do (
        if exist %%~P (
            set "PYTHON_EXE=%%~P"
            echo Found Anaconda/Miniconda: !PYTHON_EXE!
            echo WARNING: Anaconda Python may have SSL issues
            echo TIP: Consider using install_sxm_viewer_conda.bat instead
            goto :found_python
        )
    )
)

:found_python

if not defined PYTHON_EXE (
    echo.
    echo ERROR: No Python installation found!
    echo.
    echo Please install Python 3.9-3.13 from:
    echo   https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

REM Test if the interpreter works
echo Testing Python interpreter...
if /i "!PYTHON_EXE!"=="py -3" (
    py -3 -c "import sys; print(f'Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
    if !errorlevel! neq 0 (
        echo ERROR: py launcher failed
        pause
        exit /b 1
    )
) else if /i "!PYTHON_EXE!"=="python" (
    python -c "import sys; print(f'Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
    if !errorlevel! neq 0 (
        echo ERROR: python command failed
        pause
        exit /b 1
    )
) else (
    "!PYTHON_EXE!" -c "import sys; print(f'Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
    if !errorlevel! neq 0 (
        echo ERROR: Python interpreter failed
        pause
        exit /b 1
    )
)

echo.
echo Running installer...
echo ========================================
echo.

if /i "!PYTHON_EXE!"=="py -3" (
    py -3 install.py %*
) else if /i "!PYTHON_EXE!"=="python" (
    python install.py %*
) else (
    "!PYTHON_EXE!" install.py %*
)

echo.
echo ========================================
pause
