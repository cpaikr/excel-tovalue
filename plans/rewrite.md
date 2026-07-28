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
2. [Complete] Add the CLI and test its parsing and exit behavior.
3. [Complete] Rebuild the WinForms app and verify its lifecycle and layout.
4. [Complete] Replace automation and documentation, remove tracked local IDE
   state, and run format, build, test, dependency, and publish checks.

## Completion gate

- [x] Formatting verification is clean.
- [x] Release build has zero warnings.
- [x] All tests pass.
- [x] Vulnerable-package scan is clean.
- [x] GUI and CLI publish successfully.
- [x] Docs match the final behavior.
- [x] Every slice is committed with a Conventional Commit message.
