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
  "%~dp0.venv\Scripts\python.exe" -c "import PySide6" >nul 2>nul
  if not errorlevel 1 set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
)

if not defined PYTHON_EXE (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -c "import PySide6" >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=python"
  )
)

if not defined PYTHON_EXE (
  if exist "%USERPROFILE%\Documents\Codex\.codex-python\Python311\python.exe" (
    "%USERPROFILE%\Documents\Codex\.codex-python\Python311\python.exe" -c "import PySide6" >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=%USERPROFILE%\Documents\Codex\.codex-python\Python311\python.exe"
  )
)

if not defined PYTHON_EXE (
  if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
    "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" -c "import PySide6" >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  )
)

if not defined PYTHON_EXE (
  echo A Python environment with PySide6 was not found.
  echo Run: python -m pip install -e "%~dp0"
  pause
  exit /b 1
)

"%PYTHON_EXE%" "%~dp0toolshub_gui.py" %*
if errorlevel 1 (
  echo.
  echo Toolshub failed to start.
  echo Please send this window's error message to the maintainer.
  pause
  exit /b 1
)
