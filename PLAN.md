# xlclean plan

The 2026 rewrite is recorded in [plans/rewrite.md](plans/rewrite.md). The
agent-facing CLI and consumer skill are designed in
[plans/agent-cli.md](plans/agent-cli.md). The next release renames the tool
`xlclean` and adds XLStylesTool-style cleanup; its design and order are in
[plans/style-cleanup.md](plans/style-cleanup.md).

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
- [x] Ship the `skills/excel-formula-remover` consumer skill.
- [x] Release 3.0.0 and verify the documented installation from an
  independent process.
- [x] Release 3.0.1 so every release archive carries the license terms.
- [x] Release 3.0.2 so every release archive carries third-party notices.
- [x] Evaluate the skill in fresh agent contexts
  ([2026-11-17](evals/skill/evaluation-2026-11-17.md)): four cases passed.
- [ ] Run trigger trials before considering implicit invocation, and repeat
  the installation case against a real release.
- [x] Plan the `xlclean` rename and style cleanup
  ([plans/style-cleanup.md](plans/style-cleanup.md), 2026-11-26).
- [x] Rename the source tree, executables, release assets, and installers to
  `xlclean` with no behavior change. The skill, installation docs, and
  repository keep the 3.x names until 4.0.0.
- [x] Move the CLI to subcommands: `xlclean values`, with `command` and a
  per-file `changes` object in the JSON result. `styles` arrives with the
  cleanup engine.
- [x] Extract the shared per-file safety pipeline (`WorkbookBatchPipeline`) from
  the formula engine; failure kinds and progress are now `Workbook*` types.
- [x] Style reader, planner, and independent verifier for `.xlsx`/`.xlsm`
  (`src/XlClean.Core/Styles/`).
- [x] `xlclean styles` for `.xlsx`/`.xlsm`: rewriter, `StyleCleanupEngine`,
  and the CLI command.
- [x] `xlclean styles` for `.xlsb`: BIFF12 reader, rewriter, and independent
  verifier, with record layouts confirmed against Excel-saved files and
  Excel-saved test fixtures (`tests/XlClean.Core.Tests/Fixtures/`).
- [x] 4.0.0 documentation: README install, usage, and output for `xlclean`,
  the skill moved to `skills/xlclean` and extended to `styles`, and its
  eval cases. The skill has not been re-evaluated in fresh contexts since.
- [x] Release 4.0.0 (2026-11-26): repository renamed to `cpaikr/xlclean`,
  documentation merged, `v4.0.0` tagged and published, documented install
  verified on Windows and Linux.
- [ ] Re-evaluate the `xlclean` skill in fresh agent contexts
  ([evals/skill/cases.json](evals/skill/cases.json)), including the two
  `styles` cases.
- [ ] Style cleanup in the GUI (4.1.0).

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
- The repository is public at `cpaikr/xlclean` (moved from a private
  `sjunepark` repository on 2026-11-18 as `cpaikr/excel-tovalue`, renamed for
  4.0.0; GitHub redirects the old name), and releases download without a
  login. The CLI has no update check or self-upgrade; reinstalling from a
  newer release is the upgrade path.
- The source is licensed under the Elastic License 2.0
  ([LICENSE.md](LICENSE.md)), the same license as `cpaikr/darty`. From
  3.0.1 on, every release archive holds `LICENSE.md` beside the executable;
  the installers require it and install only the executable. Archives of
  3.0.0 and earlier predate this and stay as published.
- From 3.0.2 on, release archives also hold `THIRD-PARTY-NOTICES.txt`:
  the checked-in Open XML SDK notice followed by the license and notices of
  the runtime packs MSBuild resolved for that build, so the text matches the
  bundled .NET version. A pack missing an expected file fails the publish.
  The GUI archive adds the WinForms and WPF notices vendored in
  `third-party/` from the upstream `v10.0.15` tags, because the Windows
  Desktop runtime pack ships only its license.
- The skill is explicit-invocation only until trigger trials justify more.

## Validation

- Pinned .NET SDK `10.0.110` locked restore and Release build: clean, zero
  warnings.
- Automated tests (2026-11-26, `main` toward 4.0.0): 216 passed across Core,
  CLI, and App projects, including pipeline tests with a non-formula operation.
- Excel desktop round trip of `xlclean styles` (2026-11-26, Excel 16, its own
  hidden instance): Excel built `.xlsx` and `.xlsm` workbooks with 300 junk
  custom styles, a used custom style, a built-in style, direct formatting, a
  date format, and a table. The CLI cleaned each from 303 to 3 stored styles.
  Excel reopened and resaved both outputs, wrote no repair log, and reported
  identical formatting on every checked cell. The same round trip passed for
  `.xlsb`, and for an `.xlsb` whose cell indices were renumbered after a merge
  (binary index removed). Not yet covered: workbooks with duplicate cell
  formats as other tools write them, and real-world samples.
- Excel desktop round trip (2026-11-05): `.xlsb` and `.xlsx` outputs of
  Excel-saved workbooks with every formula result type, array, dynamic-array,
  data-table, and table calculated/totals formulas reopened in Excel with
  identical values and no remaining formulas in scope.
- Direct and transitive dependency vulnerability audit: clean.
- GUI single-file and portable CLI release bundles: `2.1.1` (`.xlsb` support)
  published and checksum verified; the released CLI converted Excel-saved
  `.xlsb` samples.
- Agent-facing CLI `3.0.0` (released 2026-11-18): 108 tests pass, including
  five that run the CLI as an external process. The tag-driven release
  workflow built the artifacts and, before publishing, installed the Windows
  x64 and Linux x64 archives through their installers on their own runners
  and passed the process contract tests against each installed executable.
- Windows consumer install of `3.0.0` (2026-11-18): the README commands
  downloaded the release through `gh`. From a shell started by a packaged
  desktop app, `install.ps1` refused the redirected default directory as
  designed; with `-BinDirectory "$env:USERPROFILE\.local\bin"` it installed,
  and `verify-windows-install.ps1` passed from an independent process: file
  visible, `--version` reports `3.0.0`, and the command name resolves from
  the persisted `PATH`. The default `%LOCALAPPDATA%` destination and the
  Linux `gh` download path were not exercised outside CI.
- Login-free install after the repository became public (2026-11-18): the
  README commands downloaded and installed `3.0.0` under Windows PowerShell
  5.1 and under WSL Ubuntu 24.04, each into a scratch directory, and the
  installed executable reported `3.0.0`.
- License in archives, `3.0.1` (released 2026-11-18): the release workflow
  passed the installer and process contract tests on both runners. The
  downloaded assets matched `SHA256SUMS.txt`, and the Windows CLI, Linux CLI,
  and GUI archives each list the executable and a `LICENSE.md` identical to
  the repository's. The consumer install was not repeated for `3.0.1`.
- Third-party notices in archives, `3.0.2` (released 2026-11-20): the release
  workflow passed the installer and process contract tests on both runners.
  The downloaded assets matched `SHA256SUMS.txt`, and each archive lists the
  executable, `LICENSE.md`, and `THIRD-PARTY-NOTICES.txt`. Each notices file
  holds the Open XML SDK section and the license and notices of the
  `10.0.15` runtime pack for its own platform; the GUI file adds the Windows
  Desktop license and the WinForms and WPF notices. The consumer install
  was not repeated for `3.0.2`.
- `xlclean` 4.0.0 (released 2026-11-26): the release workflow passed version
  resolution, the Windows and Linux certification jobs, and publication. The
  release is the latest and not a pre-release. All three archives matched
  `SHA256SUMS.txt`, and the CLI archive holds `xlclean.exe`, `LICENSE.md`, and
  `THIRD-PARTY-NOTICES.txt`. The README's Windows commands, run from a scratch
  folder with a scratch `-BinDirectory`, installed `xlclean 4.0.0`, which
  cleaned the Excel-saved `.xlsb` fixture. The README's Linux commands
  installed it under WSL Ubuntu 24.04. The old repository URL redirects to the
  new one. Not exercised: the default install folders and `PATH`
  registration, and `verify-windows-install.ps1` from an independent process.
- Not covered by a test: the notices content. The installers check only the
  CLI archive entry names.
- Not covered by a test: SIGTERM cancellation.

## Open questions

- Style cleanup samples: whether the owner supplies real bloated workbooks
  (see [plans/style-cleanup.md](plans/style-cleanup.md#open-questions)).
