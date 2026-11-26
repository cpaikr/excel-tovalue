# Recover from a failed workbook

Each failed entry in `files[]` has `failure.kind` and `failure.message`. A
failed workbook publishes nothing: there is no partial output to clean up, and
other workbooks in the same run are unaffected. Fix the cause, then rerun only
the failed workbooks.

## `missing-cached-value` (values)

A formula in scope has no stored result. Typical causes: the file was
generated or edited by a library (openpyxl, a reporting tool) that writes
formulas without results, or it has never been calculated and saved by Excel.

The fix is outside this tool: the workbook must be opened in Excel,
recalculated, and saved, then converted again. Ask the user to do that, or do
it only if they have already authorized driving Excel. Do not:

- compute the values yourself and write them into the workbook;
- fall back to a different converter and present its values as Excel's.

A headless recalculation by another spreadsheet engine can differ from Excel.
Offer it only as an explicit alternative and name the engine in the report.

If the formulas without results sit only in rows or sheets the user wants
kept as formulas anyway, `--start-row` or `--exclude` can move them out of
scope. That changes what is converted, so confirm it matches the request.

## `formula-range-crosses-boundary` (values)

`--start-row` cuts through a multi-cell formula definition that cannot be
split safely: an array formula, a data table, or an Excel table formula range.
Choose a start row that falls before or after that range, or drop
`--start-row` to convert the whole sheet. Use the message to locate the
range. Do not retry the same boundary.

## `invalid-workbook`

The file is not a readable workbook of its extension: corrupted, encrypted or
password-protected, or a different format under a renamed extension.
Confirm the file opens in Excel.
An encrypted workbook must be saved without a password first. Under `styles`,
a file that opens in Excel can also fail this way when its style table does
not resolve, for example a cell that refers to a cell format the file does not
define; report the message rather than editing the file.

## `invalid-input`

The path or extension is not acceptable: the file does not exist, is not
`.xlsx`, `.xlsm`, or `.xlsb`, or was listed twice in one run. Correct the
path. For `.xls` or CSV, report that the format is unsupported.

## `unsupported-workbook-content` (styles)

The workbook contains something whose style references this tool does not
rewrite, so it cannot prove a cleaned copy looks the same: an Excel 4.0 macro
sheet, shared-workbook revision history, a Strict Open XML file, or, in
`.xlsb`, a cell record type it does not know or a row style index it cannot
keep valid. The message names which. Report
it as a limit of the tool, not a defect in the file. Do not remove the part,
convert the file to another format, or switch to another cleaner without the
user's decision. Saving a Strict file as a regular Excel workbook, or saving
an `.xlsb` as `.xlsx`, are alternatives the user can choose.

## `file-access`

The workbook could not be read or the output could not be written. Most
often the workbook is open in Excel, the output directory is not writable, or
the file is still syncing from cloud storage. Have the user close the
workbook or choose another output directory, then retry once.

## `unexpected`

An error the tool does not classify. Keep the message, retry once, and if it
repeats report the workbook and message as a tool defect instead of
substituting another conversion.

For `styles`, a message starting "The cleaned workbook failed verification"
means the independent check found a difference between the source and the
cleaned copy, so nothing was published. That is a tool defect: report the
workbook and the message, and do not retry or work around it.

## Invocation errors

A top-level `error` means no workbook was attempted. `kind: usage` (exit 2)
is a command-line mistake; correct it from the message and `--help`.
`kind: unexpected` (exit 1) is a failure outside any single workbook, such as
an output directory that cannot be created; check it, then retry once.
