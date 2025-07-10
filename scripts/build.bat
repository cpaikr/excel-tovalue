@echo off
echo Building ExcelFormulaRemover...
cd /d "%~dp0.."

echo Formatting code...
dotnet format src/ExcelFormulaRemover.csproj --verbosity minimal
if %ERRORLEVEL% NEQ 0 (
    echo Code formatting failed!
    pause
    exit /b 1
)

echo Building...
dotnet publish src/ExcelFormulaRemover.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:DebugType=None -o published
echo.
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Executable created in published\ExcelFormulaRemover.exe
) else (
    echo Build failed!
)
pause