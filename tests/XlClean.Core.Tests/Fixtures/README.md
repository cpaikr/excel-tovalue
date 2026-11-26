# Excel-saved test fixtures

These workbooks were saved by desktop Excel 16 on 2026-11-26 so tests run
against the exact records Excel writes, not records the tests construct.

- `ExcelStyles.xlsb`: five unused custom styles (`Junk 1` to `Junk 5`), a
  used custom style, the built-in `Good` style, direct fonts, fills, and a
  custom number format on numbers, text, booleans, errors, formulas, and a
  blank cell, a formatted row (3), a formatted column (J), and a table.
- `ExcelTableStyles.xlsb`: a table whose header row and second column have
  cell styles (`Head style`, `Body style`), which the table records name.

Before committing, `docProps/core.xml` was replaced (author fields and dates)
and the save-folder record group (`BrtAbsPath15` in `workbook.bin`) was
removed. Every other part's content is what Excel saved; the two edited zip
entries carry new timestamps.
