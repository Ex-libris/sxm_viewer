@echo off
setlocal

REM Installer wrapper for SXM Viewer. Runs python install.py with best available interpreter.
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
    "%USERPROFILE%\miniconda3\python.exe"
    "%USERPROFILE%\miniconda\python.exe"
    "%USERPROFILE%\anaconda3\python.exe"
) do (
    if not defined PYTHON_EXE if exist %%~P set "PYTHON_EXE=%%~P"
)

if not defined PYTHON_EXE (
    py -3 -c "import sys" >nul 2>&1 && set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
    set "PYTHON_EXE=python"
)

if /i "%PYTHON_EXE%"=="python" (
    python -c "import sys" >nul 2>&1 || (
        echo Could not find a working Python interpreter. Install Python 3.9-3.12 and retry.
        pause
        exit /b 1
    )
) else if /i "%PYTHON_EXE%"=="py -3" (
    py -3 -c "import sys" >nul 2>&1 || (
        echo Could not run "py -3". Install Python 3.9-3.12 or adjust the PYTHON variable.
        pause
        exit /b 1
    )
) else (
    "%PYTHON_EXE%" -c "import sys" >nul 2>&1 || (
        echo Interpreter %PYTHON_EXE% is not runnable. Adjust the PYTHON variable or install Python 3.9-3.12.
        pause
        exit /b 1
    )
)

echo Using %PYTHON_EXE%
if /i "%PYTHON_EXE%"=="py -3" (
    py -3 install.py %*
) else if /i "%PYTHON_EXE%"=="python" (
    python install.py %*
) else (
    "%PYTHON_EXE%" install.py %*
)
pause
