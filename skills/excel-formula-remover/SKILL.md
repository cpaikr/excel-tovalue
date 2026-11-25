---
name: excel-formula-remover
disable-model-invocation: true
description: "Replace formulas in .xlsx, .xlsm, and .xlsb workbooks with their last saved values using the excel-formula-remover CLI, producing new files and leaving the originals untouched. Excludes recalculating formulas, editing or reading cell contents, and .xls or CSV files. Explicit invocation only."
---

# Excel Formula Remover

Turn a request to "paste values", "remove formulas", or "값으로 변환" into one
CLI run and a verified report of what was written. The tool copies each
workbook, swaps every in-scope formula for the value Excel last saved with it,
and publishes the copy under a new name. Requires a shell that can run the
executable; it needs no network and no Excel installation.

## Start with the available executable

Reuse a successful check from this execution context. Otherwise run
`excel-formula-remover --version` once; recheck after an executable or
environment change, or a failure to start.

- If the command is missing, cannot start, or setup or an upgrade is
  requested, read [installation](references/installation.md). A failed
  conversion is not an installation failure.
- Otherwise run `excel-formula-remover --help` for exact options and the
  result schema. Help is the syntax authority for the installed version.

## Constraints that change the answer

- **The tool never calculates.** It can only keep the value stored in the
  file. A workbook written by a library such as openpyxl, or never opened in
  Excel after its formulas changed, has no stored values and fails with
  `missing-cached-value`. Do not work around this by inventing values. A
  recalculation by another engine, such as LibreOffice, can differ from
  Excel: offer it only as a named alternative, never as the default.
- **Stored values can be stale.** A warning about manual calculation means
  the saved values may not match the formulas. Pass that warning on; the
  conversion still succeeded, but only a recalculate-and-save in Excel makes
  the values current.
- **Nothing is overwritten.** Sources stay unchanged and outputs never replace
  an existing file, so the output name is not predictable. Read each `output`
  path from the result.
- **Scope is the user's choice.** `--start-row` keeps formulas above that row
  (for example a formula header block); `--exclude` keeps whole sheets whose
  name contains the text, case-insensitively. Use them only when the request
  asks to keep something; the default converts every formula.
- **Supported formats** are `.xlsx`, `.xlsm`, and `.xlsb`. For `.xls`, CSV, or
  an encrypted workbook, say it is unsupported instead of converting through
  another tool unasked.

## Run and interpret

Pass every workbook of the request in one invocation with an explicit
`--output` directory; quote paths. Capture the exit status and stdout. Stdout
is one JSON document; stderr stays empty.

| Exit | Meaning | Next step |
| --- | --- | --- |
| 0 | Every workbook converted | Report outputs and warnings |
| 1 | At least one workbook failed, or a top-level `error` | Keep the `converted` files; read [recovery](references/recovery.md) for each `failure.kind` |
| 2 | Invalid command line | Fix the invocation from `error.message` and `--help` |
| 130 | Canceled | `files` holds only workbooks finished before cancellation; the rest were not attempted |

A file with `removedFormulas: 0` converted nothing. Check whether the
workbook had formulas in scope, or whether `--start-row` or `--exclude`
removed them all from scope, before reporting success.

## Finish the task

Report, per workbook: the `output` path, the number of formulas removed, any
warnings verbatim, and for failures the `failure.kind` with the recovery the
user must perform. State sheets left untouched by `--exclude` and rows left
untouched by `--start-row` when they were used. A zero exit status alone does
not establish that the values are current; say so when a staleness warning
was returned.
