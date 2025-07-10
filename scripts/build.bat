@echo off
echo Building ExcelFormulaRemover...
cd /d "%~dp0.."
dotnet publish src/ExcelFormulaRemover.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:DebugType=None -o published
echo.
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Executable created in published\ExcelFormulaRemover.exe
) else (
    echo Build failed!
)
pause