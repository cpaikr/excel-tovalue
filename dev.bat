@echo off
echo Building and running ExcelFormulaRemover in development mode...
dotnet build
if %ERRORLEVEL% EQU 0 (
    echo ✓ Build successful! Starting application...
    dotnet run
) else (
    echo ✗ Build failed!
    pause
)