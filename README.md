# Excel Formula Remover

A Windows Forms application that removes formulas from Excel files and converts them to static values.

## Features

- Drag and drop interface for Excel files (.xlsx, .xlsm)
- Batch processing of multiple files
- Preserves original files by creating new copies with "_NoFormulas" suffix
- Progress tracking and error handling
- Clean removal of calculation chains to prevent corruption warnings

## Project Structure

```
excel-tovalue/
├── src/                          # Source code
│   ├── Program.cs               # Main application code
│   └── ExcelFormulaRemover.csproj # Project file
├── scripts/                     # Build and development scripts
│   ├── build.bat               # Production build script
│   └── dev.bat                 # Development build and run script
├── published/                   # Build output directory
├── output/                      # Processed Excel files
├── docs/                        # Documentation
├── CLAUDE.md                    # Development notes and environment setup
└── README.md                    # This file
```

## Requirements

- .NET 9.0 or later
- Windows operating system (WinForms dependency)
- DocumentFormat.OpenXml package (automatically restored)

## Usage

### Development

Run the development script to build and start the application:
```cmd
scripts\dev.bat
```

### Production Build

Create a single-file executable:
```cmd
scripts\build.bat
```

The executable will be created in the `published/` directory.

### Manual Commands

From Windows Command Prompt or PowerShell:

```cmd
# Build for development
dotnet build src/ExcelFormulaRemover.csproj

# Run in development mode
dotnet run --project src/ExcelFormulaRemover.csproj

# Create production build
dotnet publish src/ExcelFormulaRemover.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o published
```

## How It Works

1. **File Selection**: Drag and drop Excel files into the application or use the file browser
2. **Output Directory**: Choose where processed files should be saved
3. **Processing**: Click "Remove Formulas" to start batch processing
4. **Results**: Files are saved with "_NoFormulas" suffix containing only static values

## Technical Details

- Uses DocumentFormat.OpenXml for Excel file manipulation
- Preserves cell values while removing formula definitions
- Removes calculation chains to prevent Excel corruption warnings
- Handles both .xlsx and .xlsm file formats

## Environment Notes

This project requires Windows for building and running due to WinForms dependencies. See `CLAUDE.md` for WSL compatibility information.