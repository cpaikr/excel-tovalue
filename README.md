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

## Known Issues

### File Size Increase After Formula Removal

When converting formulas to values, the output file size may increase rather than decrease. This counterintuitive behavior occurs because:

1. **Loss of Formula Optimizations**: Excel uses shared formulas and array formulas to store repeated patterns efficiently. For example, if column C contains `=A1+B1`, `=A2+B2`, etc., Excel stores the pattern once and references it multiple times. When converted to values, each cell must store its value individually.

2. **Data Type Storage**: The current implementation preserves values as text strings, which can take more space than optimized numeric storage.

3. **Full Precision Storage**: Numbers are stored with full double precision (15-17 digits), even when the original display format showed fewer decimal places.

4. **Metadata Overhead**: Cell metadata from formulas isn't fully cleaned up during conversion.

## Possible Enhancements

### File Size Optimization

1. **Implement Proper Data Type Conversion**
   - Parse numeric strings back to native number types
   - Use appropriate cell data types (numeric vs string)
   - Remove unnecessary text formatting from numeric cells

2. **Add Precision Control**
   - Detect original number formatting
   - Round values to displayed precision
   - Provide user option for decimal place limits

3. **Optimize Cell Storage**
   - Clean up formula-specific cell attributes
   - Remove unnecessary cell styles
   - Implement shared string table for repeated text values

4. **Advanced Features**
   - Add file size comparison reporting
   - Provide compression options for final save
   - Implement batch optimization for large worksheets
   - Add option to preserve formula results in comments

### Example of Storage Difference

```xml
<!-- Before: Shared formula (compact) -->
<c r="C1"><f t="shared" ref="C1:C100">A1+B1</f><v>10</v></c>
<c r="C2"><f t="shared" si="0"/><v>20</v></c>

<!-- After: Individual values (larger) -->
<c r="C1"><v>10.000000000000</v></c>
<c r="C2"><v>20.000000000000</v></c>
```

### Temporary Workarounds

Until these optimizations are implemented, users can:
- Open the converted file in Excel and use "Save As" to re-optimize
- Use Excel's "Reduce File Size" feature after conversion
- Consider whether formula removal is necessary for their use case
