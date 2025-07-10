@echo off
echo Formatting C# code...
cd /d "%~dp0.."

dotnet format src/ExcelFormulaRemover.csproj --verbosity normal

if %ERRORLEVEL% EQU 0 (
    echo Code formatting completed successfully!
) else (
    echo Code formatting failed!
)
pause