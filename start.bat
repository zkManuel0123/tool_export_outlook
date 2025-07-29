@echo off
cd /d "%~dp0"
echo Starting Outlook Calendar Export Tool...
echo.

REM Check if PowerShell is available
powershell.exe -Command "Get-Host" >nul 2>&1
if errorlevel 1 (
    echo Error: Cannot find PowerShell
    pause
    exit /b 1
)

REM 运行PowerShell脚本，指定编码
powershell.exe -ExecutionPolicy Bypass -Command "& {Set-ExecutionPolicy Bypass -Scope Process; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; & '.\export_outlook.ps1'}"

REM If script execution fails, show error message
if errorlevel 1 (
    echo.
    echo Script execution may have encountered problems, please check the error messages above
    pause
)

echo.
echo Program execution completed
pause