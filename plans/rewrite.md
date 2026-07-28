# 2026 rewrite

## Audit summary

The tagged implementation builds, but it has no tests and reports success for
states it cannot guarantee. The form owns UI layout, input validation, batch
orchestration, cancellation, filesystem writes, and OpenXML mutation. This
causes concrete failures:

- cancellation is swallowed inside the cell loop;
- output files can overwrite or delete existing workbooks;
- failed conversions can leave partial files behind;
- controls remain mutable during processing and a worker reads a UI control;
- missing cached formula values are silently left as formulas;
- each workbook is scanned twice while progress remains indeterminate;
- file intake is drag-only and the fixed-pixel layout is not DPI-accessible;
- current dependencies include a vulnerable transitive package;
- build scripts format source as a side effect and pause for manual input.

## Target shape

### Conversion module

Expose one batch interface that accepts immutable input paths, output
directory, start row, and excluded-sheet patterns. Hide path normalization,
unique output naming, temporary-file cleanup, OpenXML mutation, and per-file
results behind it. Tests exercise this interface with generated workbooks in
temporary directories.

### Command-line adapter

Provide a dependency-free command-line adapter for unattended use. It prints
deterministic per-file results and returns distinct success, conversion-error,
and usage-error exit codes.

### Windows adapter

Keep one WinForms window. It captures an immutable request, runs the conversion
module off the UI thread, and renders progress and a single batch summary. Use
native controls, system colors and fonts, scalable layout containers, keyboard
file selection, clear focus order, accessible descriptions, and safe close /
cancel behavior.

## Delivery slices

1. [Complete] Restructure the solution, upgrade to supported .NET 10 LTS and
   OpenXML, implement the conversion module, and add behavior tests.
2. Add the CLI and test its parsing and exit behavior.
3. Rebuild the WinForms app and verify its lifecycle and layout.
4. Replace automation and documentation, remove tracked local IDE state, and
   run format, build, test, dependency, and publish checks.

## Completion gate

- formatting verification is clean;
- Release build has zero warnings;
- all tests pass;
- vulnerable-package scan is clean;
- GUI and CLI publish successfully;
- docs match the final behavior;
- every slice is committed with a Conventional Commit message.
