# Excel Formula Remover plan

The current rewrite is tracked in [plans/rewrite.md](plans/rewrite.md).

## Goal

Provide a dependable Windows utility and command-line workflow that replace
Excel formulas with the cached values stored in `.xlsx` and `.xlsm` files,
without modifying source workbooks or overwriting earlier outputs.

## Milestones

- [x] Preserve the former implementation as tag
  `before-rewrite-2026-07-28`.
- [x] Audit correctness, architecture, dependencies, automation, and UX.
- [x] Extract and test the workbook conversion engine.
- [x] Add a non-interactive command-line workflow.
- [x] Rebuild the WinForms workflow around the tested engine.
- [ ] Replace stale build scripts and documentation; run the full release gate.

## Product constraints

- The tool does not calculate formulas. It uses the last cached results saved
  in the workbook.
- A target formula without a cached result is a file-level failure; no partial
  output is published.
- Existing output files are preserved. Name collisions receive a numbered
  suffix.
- `.xlsm` VBA content and formulas outside the selected scope are preserved.
- Windows is required only for the graphical app. The conversion engine and
  command-line workflow remain independently testable.

## Open questions

None. The user delegated implementation and product trade-offs for this
rewrite.
