@echo off
setlocal enabledelayedexpansion

REM Launcher for SXM Viewer. Prefers the local .venv, falls back to Conda envs, then PATH python.
set "SCRIPT_DIR=%~dp0"
set "ROOT_DIR=%~dp0.."
REM Run from repo root so package imports resolve correctly.
cd /d "%ROOT_DIR%"

set "PYTHON_EXE="

if defined PYTHON (
    set "PYTHON_EXE=%PYTHON%"
)

if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    if not defined PYTHON_EXE set "PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"
)

if exist ".venv\Scripts\python.exe" (
    if not defined PYTHON_EXE set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
)

if not defined PYTHON_EXE if defined CONDA_PREFIX if exist "%CONDA_PREFIX%\python.exe" (
    set "PYTHON_EXE=%CONDA_PREFIX%\python.exe"
)

for %%P in (
    "%USERPROFILE%\miniconda3\envs\sxm_viewer\python.exe"
    "%USERPROFILE%\miniconda\envs\sxm_viewer\python.exe"
    "%USERPROFILE%\anaconda3\envs\sxm_viewer\python.exe"
) do (
    if not defined PYTHON_EXE if exist %%~P set "PYTHON_EXE=%%~P"
)

if not defined PYTHON_EXE (
    py -3 -c "import sys" >nul 2>&1
    if !errorlevel! equ 0 set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
    python -c "import sys" >nul 2>&1
    if !errorlevel! equ 0 set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
    echo Could not find a working Python interpreter. Run install_sxm_viewer.bat first.
    pause
    exit /b 1
)

REM Test if the interpreter works
if /i "!PYTHON_EXE!"=="py -3" (
    py -3 -c "import sys" >nul 2>&1
) else if /i "!PYTHON_EXE!"=="python" (
    python -c "import sys" >nul 2>&1
) else (
    "!PYTHON_EXE!" -c "import sys" >nul 2>&1
)

if !errorlevel! neq 0 (
    echo Interpreter !PYTHON_EXE! is not runnable. Run install_sxm_viewer.bat or fix PYTHON.
    pause
    exit /b 1
)

REM Check if dependencies are installed
set "IMPORT_FAILED="
if /i "!PYTHON_EXE!"=="py -3" (
    py -3 -c "import sxm_viewer; import PyQt5" >nul 2>&1 || set "IMPORT_FAILED=1"
) else if /i "!PYTHON_EXE!"=="python" (
    python -c "import sxm_viewer; import PyQt5" >nul 2>&1 || set "IMPORT_FAILED=1"
) else (
    "!PYTHON_EXE!" -c "import sxm_viewer; import PyQt5" >nul 2>&1 || set "IMPORT_FAILED=1"
)

if defined IMPORT_FAILED (
    echo.
    echo Launch failed: dependencies are missing for !PYTHON_EXE!.
    echo Run "install_sxm_viewer.bat" or "python install.py --reset" to rebuild the environment.
    pause
    exit /b 1
)

echo Using !PYTHON_EXE!

REM Try to load a local .env automatically if python-dotenv is available
set "DOTENV_OK="
if /i "!PYTHON_EXE!"=="py -3" (
    py -3 -m dotenv --version >nul 2>&1 && set "DOTENV_OK=1"
) else if /i "!PYTHON_EXE!"=="python" (
    python -m dotenv --version >nul 2>&1 && set "DOTENV_OK=1"
) else (
    "!PYTHON_EXE!" -m dotenv --version >nul 2>&1 && set "DOTENV_OK=1"
)

if defined DOTENV_OK (
    echo Loading .env ^(if present^) via python-dotenv...
    if /i "!PYTHON_EXE!"=="py -3" (
        py -3 -m dotenv -q run -- "!PYTHON_EXE!" -m sxm_viewer
    ) else if /i "!PYTHON_EXE!"=="python" (
        python -m dotenv -q run -- "!PYTHON_EXE!" -m sxm_viewer
    ) else (
        "!PYTHON_EXE!" -m dotenv -q run -- "!PYTHON_EXE!" -m sxm_viewer
    )
) else (
    echo Tip: install python-dotenv to auto-load .env ^(pip install python-dotenv^)
    if /i "!PYTHON_EXE!"=="py -3" (
        py -3 -m sxm_viewer
    ) else if /i "!PYTHON_EXE!"=="python" (
        python -m sxm_viewer
    ) else (
        "!PYTHON_EXE!" -m sxm_viewer
    )
)
