@echo off
setlocal

REM Installer wrapper for SXM Viewer. Runs python install.py with best available interpreter.
cd /d "%~dp0"

set "PYTHON_EXE="

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
    set "PYTHON_EXE=python"
)

echo Using %PYTHON_EXE%
"%PYTHON_EXE%" install.py
pause
