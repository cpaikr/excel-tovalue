@echo off
echo Building ExcelFormulaRemover (Development - Multi-file)...
cd /d "%~dp0.."
dotnet publish src/ExcelFormulaRemover.csproj -c Debug -r win-x64 --self-contained -o published-dev
echo.
if %ERRORLEVEL% EQU 0 (
    echo Development build successful! Files created in published-dev\
    echo Use this for faster development builds and debugging.
) else (
    echo Build failed!
)
pause