@echo off
echo Building ExcelFormulaRemover (Development - Multi-file)...
cd /d "%~dp0.."

echo Formatting code...
dotnet format src/ExcelFormulaRemover.csproj --verbosity minimal
if %ERRORLEVEL% NEQ 0 (
    echo Code formatting failed!
    pause
    exit /b 1
)

echo Building...
dotnet publish src/ExcelFormulaRemover.csproj -c Debug -r win-x64 --self-contained -o published-dev
echo.
if %ERRORLEVEL% EQU 0 (
    echo Development build successful! Files created in published-dev\
    echo Use this for faster development builds and debugging.
) else (
    echo Build failed!
)
pause