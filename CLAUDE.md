# Excel Formula Remover Development Notes

## Environment Issues

### WSL Compatibility Problem
- **Issue**: .NET WinForms applications cannot be built in WSL environment
- **Error**: Missing Windows Desktop SDK targets (`Microsoft.NET.Sdk.WindowsDesktop.targets`)
- **Root Cause**: WinForms is Windows-specific and requires Windows .NET SDK

### Solution
- **Development Environment**: Use Windows Command Prompt/PowerShell instead of WSL
- **Build Commands**: Run `dotnet build` and `dotnet publish` from Windows terminal
- **File Access**: Project files are accessible from Windows at `C:\Users\user\Documents\IT\excel-tovalue\`

## Build Instructions

### Development Workflow
When working on this project in WSL, **ask the user to run Windows commands** since WSL cannot build WinForms applications.

**Prompt the user with:**
"Please run the following command in Windows Command Prompt or PowerShell:"

### From Windows Command Prompt:
```cmd
cd C:\Users\user\Documents\IT\excel-tovalue
dotnet build
dotnet run
```

### For Single-File Distribution:
```cmd
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## Project Status
- ✅ Complete C# implementation exists
- ✅ Project file created with correct dependencies
- ❌ Cannot build in WSL - requires Windows environment
- ⏳ Needs user to run Windows terminal commands for build/test process

## Development Pattern
1. Edit code files in WSL/Claude Code
2. Ask user to run build/test commands in Windows terminal
3. User provides output back to continue development
4. Repeat as needed for debugging/testing