@echo off
setlocal

cd /d "%~dp0"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "LOCAL_PYTHONPATH=%~dp0.codex\python-packages"

if exist "%LOCAL_PYTHONPATH%" (
  if defined PYTHONPATH (
    set "PYTHONPATH=%LOCAL_PYTHONPATH%;%PYTHONPATH%"
  ) else (
    set "PYTHONPATH=%LOCAL_PYTHONPATH%"
  )
)

set "PYTHON_EXE="

if exist "%~dp0.venv\Scripts\python.exe" (
  set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
)

if not defined PYTHON_EXE (
  if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
    set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  )
)

if not defined PYTHON_EXE (
  where python >nul 2>nul
  if not errorlevel 1 set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
  where py >nul 2>nul
  if not errorlevel 1 set "PYTHON_EXE=py"
)

if not defined PYTHON_EXE (
  echo Python was not found.
  echo Please install Python, or ask the maintainer to configure this launcher.
  pause
  exit /b 1
)

"%PYTHON_EXE%" "%~dp0toolshub_gui.py"
if errorlevel 1 (
  echo.
  echo Toolshub failed to start.
  echo Please send this window's error message to the maintainer.
  pause
  exit /b 1
)
