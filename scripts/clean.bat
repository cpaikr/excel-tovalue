@echo off
echo Cleaning build outputs...
cd /d "%~dp0.."

if exist "published" (
    echo Removing published\
    rmdir /s /q "published"
)

if exist "published-dev" (
    echo Removing published-dev\
    rmdir /s /q "published-dev"
)

if exist "src\bin" (
    echo Removing src\bin\
    rmdir /s /q "src\bin"
)

if exist "src\obj" (
    echo Removing src\obj\
    rmdir /s /q "src\obj"
)

echo Clean complete!
pause