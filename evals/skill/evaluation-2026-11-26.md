# Skill evaluation, 2026-11-26

Evaluation of `skills/xlclean` as released in 4.0.0, which renamed
`excel-formula-remover` and added the `styles` command. Each trial ran once in
a fresh agent context (Sonnet 5.5) that was given a copy of the package outside
its source path and told to read only what the skill routes to. Requests are in
[cases.json](cases.json).

The two live trials executed the CLI for real against workbooks that desktop
Excel and openpyxl wrote: the installed 4.0.0 release, and for the rerun the
build with the fix below. The other four are decision trials with simulated
observations; they show what the agent would do, not that the commands work.
Activation was not tested: the skill is explicit-invocation only and every
request invoked it by name.

| Case | Files read | Result |
| --- | --- | --- |
| `partial-batch` (live) | `SKILL.md`; rerun also `recovery.md` | Pass on rerun; the first run found a tool defect, now fixed |
| `style-bloat` (live) | `SKILL.md` | Pass |
| `missing-cli` | `SKILL.md`, `installation.md` | Pass |
| `boundary-failure` | `SKILL.md`, `recovery.md` | Pass |
| `unsupported-request` | `SKILL.md` | Pass |
| `unsupported-styles` | `SKILL.md`, `recovery.md` | Pass |

## Observations

- `partial-batch`: one `values` invocation for both workbooks. Against 4.0.0 it
  exited 0: Forecast.xlsx, written by openpyxl, "converted" with one formula
  removed. openpyxl writes `<v></v>` for formulas it never calculated, and the
  engine accepted the empty value as a cached result, so the output cell was
  silently blank. The agent reported what the tool returned, including the
  manual-calculation warning for Sales.xlsx verbatim, so the skill behaved;
  the tool did not. The engine now treats an empty value as missing unless
  the result is text (`t="str"`, which Excel uses for a formula returning
  ""). With the fix the run exits 1 with `missing-cached-value` for
  Forecast.xlsx. In the rerun the agent treated exit 1 as a partial batch,
  read `recovery.md`, reported the converted Sales.xlsx output and the
  warning verbatim, reported `missing-cached-value` with an Excel
  recalculate-and-save, invented no values, and named LibreOffice only as an
  alternative it had not used. The rerun used a local build, which reports
  version 1.0.0 because only release builds are stamped.
- `style-bloat`: chose `styles` for "Too many different cell formats". It
  reported the output path, 2001 to 1 cell styles and no merges, passed on the
  macro warning verbatim with what to test, and said it had not confirmed the
  Excel message is gone. It claimed nothing about number formats or fonts.
- `missing-cli`: read only `SKILL.md` and `installation.md`, did not install
  or build from source, and asked for authorization to install from the
  public release.
- `boundary-failure`: did not retry the same start row; laid out the two
  safe boundaries, said how each changes the requested scope, and asked the
  user to choose.
- `unsupported-request`: predicted `missing-cached-value` for the openpyxl
  workbook without running it, declined to recalculate by default, offered
  LibreOffice only as a named alternative with the caveat that its results
  can differ from Excel's, and reported `.xls` as unsupported. The caveat
  that the 2026-11-17 evaluation found missing now reaches the agent from
  `SKILL.md`.
- `unsupported-styles`: reported Data.xlsb's output and counts, reported the
  macro sheet as a tool limit, and did not remove it or switch tools.

## Not covered

- Installation against a real release inside a trial; the release install was
  verified separately (PLAN.md).
- Implicit activation and near-miss requests.
- A second model family or client.
- Repeated runs; each case ran once.
