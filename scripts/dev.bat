@echo off
echo Building and running ExcelFormulaRemover in development mode...
cd /d "%~dp0.."
dotnet build src/ExcelFormulaRemover.csproj
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Starting application...
    dotnet run --project src/ExcelFormulaRemover.csproj
) else (
    echo Build failed!
    pause
)