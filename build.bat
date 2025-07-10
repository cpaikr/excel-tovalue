@echo off
echo Building ExcelFormulaRemover...
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o published
echo.
if %ERRORLEVEL% EQU 0 (
    echo ✓ Build successful! Executable created in published\ExcelFormulaRemover.exe
) else (
    echo ✗ Build failed!
)
pause