# Excel Formula Remover plan

The 2026 rewrite is recorded in [plans/rewrite.md](plans/rewrite.md). The
agent-facing CLI and consumer skill are designed in
[plans/agent-cli.md](plans/agent-cli.md).

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
- [x] Make the CLI an agent-facing process contract: bare invocation equals
  `--help`, `--version`, one JSON document on stdout.
- [x] Publish standalone CLI executables for Windows x64 and Linux x64 with
  checksum-verifying installers and a tag-driven release workflow.
- [x] Ship the `skill/excel-formula-remover` consumer skill.
- [ ] Release 3.0.0 and verify the documented installation from an
  independent terminal.
- [ ] Evaluate the skill in fresh agent contexts, then decide on implicit
  invocation.

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
- The CLI is a versioned process API. Stdout is one newline-terminated JSON
  document for every run that attempts a conversion, stderr stays empty, and
  failure kinds are stable identifiers. JSON is the only encoding; the 2.x
  TOON and NDJSON outputs were removed in 3.0.0.
- CLI executables are self-contained and untrimmed because the OpenXML SDK
  depends on reflection. They run in invariant-globalization mode, so Linux
  hosts need no ICU; the engine uses only ordinal and invariant operations.
- The repository and its releases are private, so installation downloads
  through an authenticated GitHub CLI. The CLI has no update check or
  self-upgrade; reinstalling from a newer release is the upgrade path.
- The skill is explicit-invocation only until trigger trials justify more.

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

- Agent-facing CLI (2026-11-17, unreleased): 108 tests pass, including five
  that run the CLI as an external process. A local `publish.ps1` run installed
  the Windows x64 archive through `install.ps1` and passed the process
  contract tests against the installed executable. The Linux x64 archive was
  installed with `install.sh` and run under WSL Ubuntu 24.04 without a .NET
  runtime (`--version`, help, usage error, checksum rejection); the contract
  test suite has not yet run against it. The GitHub workflows are unexercised
  until pushed.

## Open questions

None. The user delegated implementation and product trade-offs for this
rewrite.
