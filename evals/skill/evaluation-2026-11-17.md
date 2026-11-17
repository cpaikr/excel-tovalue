# Skill evaluation, 2026-11-17

First evaluation of `skill/excel-formula-remover`, against no prior skill.
Each trial ran once in a fresh agent context (Sonnet 5.5) that was given a
copy of the package outside its source path and told to read only what the
skill routes to. Requests are in [cases.json](cases.json).

One trial executed the CLI for real against generated workbooks, using the
locally built executable rather than an installed release. The other three
are decision trials with simulated observations; they show what the agent
would do, not that the commands work. Activation was not tested: the skill is
explicit-invocation only and every request invoked it by name.

| Case | Files read | Result |
| --- | --- | --- |
| `partial-batch` (live) | `SKILL.md`, `installation.md`, `recovery.md` | Pass |
| `missing-cli` | `SKILL.md`, `installation.md` | Pass |
| `boundary-failure` | `SKILL.md`, `recovery.md` | Pass |
| `unsupported-request` | `SKILL.md` | Pass, one gap fixed |

## Observations

- `partial-batch`: one invocation for both workbooks, exit 1. Reported the
  converted file's actual output path, the count of two formulas, and the
  manual-calculation warning verbatim with its meaning. Reported
  `missing-cached-value` for the other file, did not invent values, and
  offered to rerun only that file after an Excel save. It read
  `installation.md` unnecessarily: the trial told it to prepend a Windows
  path to `PATH`, which Git Bash did not resolve, so the first `--version`
  failed. That is a harness artifact, not skill guidance.
- `missing-cli`: with no executable and no `gh` login, it stopped, did not
  start a login, build from source, or convert through another tool, and
  asked for the release files or an authorized install.
- `boundary-failure`: did not retry the same start row. It recognized that
  every available boundary changes what the user asked to keep and asked the
  user to choose, without running anything.
- `unsupported-request`: predicted the `missing-cached-value` failure for an
  openpyxl workbook, declined to recalculate, and reported `.xls` as
  unsupported. It offered a LibreOffice recalculation without saying its
  values can differ from Excel's, because that caveat was only in
  `recovery.md`, which loads after a failure. The caveat now also sits in
  `SKILL.md`. The changed wording has not been re-run.

## Not covered

- Installation against a real release, and the `gh`-authenticated path.
- Implicit activation and near-miss requests.
- A second model family or client.
- Repeated runs; each case ran once.
