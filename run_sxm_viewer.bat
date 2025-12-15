@echo off
setlocal

REM Launcher for SXM Viewer. Prefers the local .venv, falls back to Conda envs, then PATH python.
cd /d "%~dp0"

set "PYTHON_EXE="

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
    set "PYTHON_EXE=python"
)

echo Using %PYTHON_EXE%
"%PYTHON_EXE%" -m sxm_viewer
if errorlevel 1 (
    echo.
    echo Launch failed. Please run "python install.py" once to install dependencies.
    pause
)
