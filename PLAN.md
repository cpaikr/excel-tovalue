# Excel Formula Remover plan

The current rewrite is tracked in [plans/rewrite.md](plans/rewrite.md).

## Goal

Provide a dependable Windows utility and command-line workflow that replace
Excel formulas with the cached values stored in `.xlsx`, `.xlsm`, and `.xlsb`
files, without modifying source workbooks or overwriting earlier outputs.

## Milestones

- [x] Preserve the former implementation as tag
  `before-rewrite-2026-07-28`.
- [x] Audit correctness, architecture, dependencies, automation, and UX.
- [x] Extract and test the workbook conversion engine.
- [x] Add a non-interactive command-line workflow.
- [x] Rebuild the WinForms workflow around the tested engine.
- [x] Replace stale build scripts and documentation; run the full release gate.
- [x] Support `.xlsb` by rewriting BIFF12 worksheet and table records with the
  same scope and boundary rules as `.xlsx`.

## Product constraints

- The tool does not calculate formulas. It uses the last cached results saved
  in the workbook.
- A target formula without a cached result is a file-level failure; no partial
  output is published.
- Existing output files are preserved. Name collisions receive a numbered
  suffix.
- `.xlsm`/`.xlsb` VBA content and formulas outside the selected scope are
  preserved.
- `.xlsb` is rewritten at the [MS-XLSB] record level without a third-party
  library; only formula-bearing records change. Record layouts were confirmed
  against files saved by Excel. A BIFF12 formula record always stores its
  result, so `.xlsb` cannot fail with a missing cached value.
- Table totals-row functions are kept in both formats while calculated-column
  and custom totals formulas are removed. Excel refuses to open a table whose
  totals cell holds a value but whose column has no totals function; removing
  the function was the earlier `.xlsx` behavior and produced such files.
- Windows is required only for the graphical app. The conversion engine and
  command-line workflow remain independently testable.

## Validation

- Pinned .NET SDK `10.0.110` locked restore and Release build: clean, zero
  warnings.
- Automated tests: 93 passed across Core, CLI, and App projects.
- Excel desktop round trip (2026-11-05): `.xlsb` and `.xlsx` outputs of
  Excel-saved workbooks with every formula result type, array, dynamic-array,
  data-table, and table calculated/totals formulas reopened in Excel with
  identical values and no remaining formulas in scope.
- Direct and transitive dependency vulnerability audit: clean.
- GUI single-file and portable CLI release bundles: `2.1.1` (`.xlsb` support)
  published and checksum verified; the released CLI converted Excel-saved
  `.xlsb` samples.

## Open questions

None. The user delegated implementation and product trade-offs for this
rewrite.
