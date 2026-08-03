@echo off
chcp 65001 >nul

if not "%~1"=="" goto run

"%~dp0QAtools-CLI.exe" --help
echo.
pause
exit /b

:run
"%~dp0QAtools-CLI.exe" %*
exit /b %errorlevel%
