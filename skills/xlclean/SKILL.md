---
name: xlclean
disable-model-invocation: true
description: "Clean .xlsx, .xlsm, and .xlsb workbooks into new files with the xlclean CLI, leaving the originals untouched: replace formulas with their last saved values, or remove unused cell styles and formats and merge identical formats (Excel's 'Too many different cell formats', thousands of copied styles). Excludes recalculating formulas, editing or reading cell contents, and .xls or CSV files. Explicit invocation only."
---

# xlclean

Turn a request to "paste values", "remove formulas", "값으로 변환", "clean up
styles", or "too many different cell formats" into one CLI run and a verified
report of what was written. The tool copies each workbook, changes the copy,
checks it against the source, and publishes it under a new name. Requires a
shell that can run the executable; it needs no network and no Excel
installation.

## Start with the available executable

Reuse a successful check from this execution context. Otherwise run
`xlclean --version` once; recheck after an executable or environment change,
or a failure to start.

- If the command is missing, cannot start, or setup or an upgrade is
  requested, read [installation](references/installation.md). A failed run is
  not an installation failure.
- An installed `excel-formula-remover` is the 3.x release of this tool. Treat
  `xlclean` as missing and follow [installation](references/installation.md);
  do not adapt to the 3.x command line.
- Otherwise run `xlclean --help` and `xlclean <command> --help` for exact
  options and the result schema. Help is the syntax authority for the
  installed version.

## Choose the command

| Request | Command |
| --- | --- |
| Values instead of formulas, paste-as-values, freeze results | `values` |
| Unused or duplicate styles, too many cell formats, slow or bloated formatting | `styles` |
| Both | `values`, then `styles` on the `values` output |

## Constraints that change the answer

- **`values` never calculates.** It can only keep the value stored in the
  file. A workbook written by a library such as openpyxl, or never opened in
  Excel after its formulas changed, has no stored values and fails with
  `missing-cached-value`. Do not work around this by inventing values. A
  recalculation by another engine, such as LibreOffice, can differ from
  Excel: offer it only as a named alternative, never as the default.
- **Stored values can be stale.** A warning about manual calculation means
  the saved values may not match the formulas. Pass that warning on; the
  conversion still succeeded, but only a recalculate-and-save in Excel makes
  the values current.
- **`styles` changes no cell's appearance.** It keeps built-in styles, styles
  a table names, styles cells use (even when two look alike under different
  names), and every number format, font, fill, border, and conditional
  format. Do not promise it removes those, or that it fixes slowness with
  other causes (excess used range, volatile formulas).
- **Nothing is overwritten.** Sources stay unchanged and outputs never replace
  an existing file, so the output name is not predictable. Read each `output`
  path from the result.
- **`values` scope is the user's choice.** `--start-row` keeps formulas above
  that row (for example a formula header block); `--exclude` keeps whole
  sheets whose name contains the text, case-insensitively. Use them only when
  the request asks to keep something; the default converts every formula.
  `styles` has no scope options.
- **Supported formats** are `.xlsx`, `.xlsm`, and `.xlsb`. For `.xls`, CSV, or
  an encrypted workbook, say it is unsupported instead of converting through
  another tool unasked.

## Run and interpret

Pass every workbook of the request in one invocation with an explicit
`--output` directory; quote paths. Capture the exit status and stdout. Stdout
is one JSON document; stderr stays empty.

| Exit | Meaning | Next step |
| --- | --- | --- |
| 0 | Every workbook succeeded | Report outputs, changes, and warnings |
| 1 | At least one workbook failed, or a top-level `error` | Keep the `converted` files; read [recovery](references/recovery.md) for each `failure.kind` |
| 2 | Invalid command line | Fix the invocation from `error.message` and `--help` |
| 130 | Canceled | `files` holds only workbooks finished before cancellation; the rest were not attempted |

Read each file's `changes` before reporting success:

- `values`: `removedFormulas: 0` means nothing was converted. Check whether
  the workbook had formulas in scope, or whether `--start-row` or `--exclude`
  removed them all from scope.
- `styles`: equal `before` and `after` counts with `mergedCellFormats: 0`
  mean no visible reduction; report that rather than a cleanup.

## Finish the task

Report, per workbook: the `output` path, what changed (formulas removed, or
cell styles and formats before and after with merges), any warnings verbatim,
and for failures the `failure.kind` with the recovery the user must perform.
For `values`, state sheets left untouched by `--exclude` and rows left
untouched by `--start-row` when they were used, and do not let a zero exit
status stand for current values when a staleness warning was returned. For
`styles`, pass on a macro warning: a macro that applies a removed style by
name will no longer find it.
