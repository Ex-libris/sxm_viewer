@echo off
setlocal

REM Launcher for SXM Viewer. Prefers the local .venv, falls back to Conda envs, then PATH python.
cd /d "%~dp0"

set "PYTHON_EXE="

if defined PYTHON (
    set "PYTHON_EXE=%PYTHON%"
)

if exist ".venv\Scripts\python.exe" (
    set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
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
    py -3 -c "import sys" >nul 2>&1 && set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
    set "PYTHON_EXE=python"
)

set "IMPORT_CHECK=import importlib; importlib.import_module('sxm_viewer'); importlib.import_module('PyQt5')"

if /i "%PYTHON_EXE%"=="python" (
    python -c "import sys" >nul 2>&1 || (
        echo Could not find a working Python interpreter. Run install_sxm_viewer.bat first.
        pause
        exit /b 1
    )
    python -c "%IMPORT_CHECK%" >nul 2>&1 || set "IMPORT_FAILED=1"
) else if /i "%PYTHON_EXE%"=="py -3" (
    py -3 -c "import sys" >nul 2>&1 || (
        echo Could not run "py -3". Install Python 3.9-3.12 or adjust the PYTHON variable.
        pause
        exit /b 1
    )
    py -3 -c "%IMPORT_CHECK%" >nul 2>&1 || set "IMPORT_FAILED=1"
) else (
    "%PYTHON_EXE%" -c "import sys" >nul 2>&1 || (
        echo Interpreter %PYTHON_EXE% is not runnable. Run install_sxm_viewer.bat or fix PYTHON.
        pause
        exit /b 1
    )
    "%PYTHON_EXE%" -c "%IMPORT_CHECK%" >nul 2>&1 || set "IMPORT_FAILED=1"
)

if defined IMPORT_FAILED (
    echo.
    echo Launch failed: dependencies are missing for %PYTHON_EXE%.
    echo Run "install_sxm_viewer.bat" or "python install.py --reset" to rebuild the environment.
    pause
    exit /b 1
)

echo Using %PYTHON_EXE%
if /i "%PYTHON_EXE%"=="py -3" (
    py -3 -m sxm_viewer
) else if /i "%PYTHON_EXE%"=="python" (
    python -m sxm_viewer
) else (
    "%PYTHON_EXE%" -m sxm_viewer
)
