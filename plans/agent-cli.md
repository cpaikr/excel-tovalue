# Agent-facing CLI and consumer skill

Status: implemented and released as 3.0.0 on 2026-11-18. The owner chose the
recommended option for both decisions: JSON as the only encoding, and a
private release channel. The repository later became public under
`cpaikr`, so the private-channel reasoning below is historical.
[PLAN.md](../PLAN.md) owns current state.

## Goal

Let an AI agent convert workbooks with `excel-formula-remover` unattended:
find or install a runnable executable, invoke it, read one structured result,
and recover from failures with the right domain judgment.

The CLI already exists. This plan adapts it and adds a skill; it does not
rewrite the engine or change languages. References:
[darty](https://github.com/cpaikr/darty) (shape of a CLI plus consumer skill)
and the mytech practices for automation contracts, standalone distribution,
installation guides, and CLI consumer skills.

## Gaps in the current CLI

| Area | Today | Target |
| --- | --- | --- |
| Executable | `dotnet excel-formula-remover.dll`, needs .NET 10 runtime | `excel-formula-remover` on `PATH`, no runtime |
| Bare invocation | Usage error, exit 2 | Byte-identical to `--help`, exit 0 |
| Version | None | `--version`, local, prints `excel-formula-remover <version>` |
| Stdout | TOON default; `--json` is multi-line NDJSON; no trailing newline | One newline-terminated document |
| Fatal errors | .NET exception type name | Project-owned `kind` |
| Tests | In-process only | Packaged executable run as an external process |
| Release | Local `publish.ps1`, manual upload | Tagged CI build, smoke-tested per target |
| Install | Unzip by hand | Checksum-verifying installer and README guide |
| Skill | None | `skill/excel-formula-remover/` |

Already aligned and kept: exit codes `0`/`1`/`2`/`130`, stable failure kinds,
no-clobber outputs with the resolved path returned, empty stderr, preflight
before any file is opened, cancellation that reports completed files.

## Design

### 1. Process contract (breaking, release as 3.0.0)

- Keep the flat command: `excel-formula-remover --output <dir> [options]
  <workbook>...`. One operation needs no subcommands or discovery framework.
- Bare invocation prints root help; add `--version` from the embedded
  informational version, stripped of any `+commit` suffix.
- Stdout carries exactly one JSON document plus `\n` for every run that
  attempts a conversion, including usage errors and cancellation:
  - `ok`, `version`;
  - `files[]` in request order: `input`, `status`, `output`, `failure`
    (`kind`, `message`), removed-formula and sheet counts, `warnings[]`;
  - `summary`: `total`, `succeeded`, `failed`, `canceled`;
  - `error` (`kind`: `usage` or `unexpected`, `message`, `help`) when the
    invocation itself fails.
- **Decide 1 — encoding.** Recommended: JSON becomes the only format; remove
  TOON and NDJSON. Reasons: mytech makes JSON the default and wants a concrete
  consumer for a second encoding; darty chose the same; scripts cannot parse
  TOON without a library; the TOON view omits the counts an agent needs to
  notice "0 formulas removed". Alternative: keep TOON as default and only
  collapse `--json` into one document (non-breaking default, two encoders to
  maintain).
- Not adding `inspect` or `--dry-run`: converting into a scratch output
  directory is already non-destructive and returns the same evidence.

Files: `src/ExcelFormulaRemover.Cli/CliApplication.cs`,
`CommandLineParser.cs`, `tests/ExcelFormulaRemover.Cli.Tests/`.

### 2. Standalone executable and release

- Publish the CLI self-contained, single-file, untrimmed, compressed, per
  runtime identifier. The GUI already ships the same engine this way, so
  OpenXML under single-file is proven; trimming and Native AOT are not, and
  are deferred until size matters.
- Targets: `win-x64` and `linux-x64` (cloud agent sandboxes are Linux, and CI
  already has an Ubuntu job to certify it). No macOS claim until a consumer
  needs one.
- Artifacts: `excel-formula-remover-cli-<version>-win-x64.zip`,
  `...-linux-x64.tar.gz`, the unchanged GUI zip, `SHA256SUMS.txt`. The
  framework-dependent "portable" zip is dropped.
- Installers adapted from darty's `scripts/install/`: `install.ps1` to
  `%LOCALAPPDATA%\excel-formula-remover\bin`, `install.sh` to `~/.local/bin`.
  Each verifies SHA-256, runs the new executable's `--version`, leaves an
  existing install untouched on failure, and does not edit `PATH`. Keep
  darty's refusal of redirected destinations and its
  `verify-windows-install.ps1`, because agent shells inside packaged desktop
  apps on Windows see a private `%LOCALAPPDATA%`.
- Artifact contract tests in `Cli.Tests`, enabled by
  `EXCEL_FORMULA_REMOVER_EXE`: bare output equals `--help`, `--version`,
  single-document stdout with trailing newline, empty stderr, each exit code,
  no-clobber naming. `publish.ps1` runs them against the extracted archive.
- `.github/workflows/release.yml` on `v*` tags: `check.ps1`, `publish.ps1`,
  artifact tests on Windows and Ubuntu runners, then create the release.
- No update advisories and no `upgrade` command. mytech: a small tool should
  not acquire an updater, and a private channel cannot be checked without
  credentials. Upgrading means rerunning the installer with a newer release.
- **Decide 2 — channel.** The repository is private, so unattended install
  needs `gh release download` with the user's existing `gh` login.
  Recommended: stay private and document that path. Making the repository
  public would allow plain `curl`/`Invoke-WebRequest` with no login.

Files: `src/ExcelFormulaRemover.Cli/ExcelFormulaRemover.Cli.csproj`,
`scripts/publish.ps1`, `scripts/install/`, `.github/workflows/`.

### 3. Consumer skill

```text
skill/excel-formula-remover/
  SKILL.md
  agents/openai.yaml
  references/
    installation.md
    recovery.md
```

- `SKILL.md` stays small: one `--version` check per context; routing to the
  references; the constraints help text cannot teach; how to read the result;
  what counts as done. Exact syntax comes from `--help`, not a copied table.
- Judgment the skill owns:
  - The tool never calculates. Workbooks written by libraries such as
    openpyxl have no cached values and fail with `missing-cached-value`; the
    fix is a recalculate-and-save in Excel, outside this tool.
  - A manual-calculation warning means values may be stale; report it, do not
    hide it.
  - Read the actual `output` path from the result; names carry a
    `_수식제거` suffix and a counter on collision.
  - `--start-row` and `--exclude` semantics, and that zero removed formulas
    usually means the scope was wrong.
  - Exit 1 with some `ok` files is a partial batch: keep the successes.
- `references/installation.md` loads only when the executable is missing or
  setup is requested, and points at the README installation section as the
  single owner of installer facts.
- `references/recovery.md` loads only on an observed failure: one entry per
  failure kind (`missing-cached-value`, `formula-range-crosses-boundary`,
  `invalid-workbook`, `file-access`, `invalid-input`, `unexpected`).
- Activation: ship explicit-only, with a description that excludes
  recalculation, `.xls`/CSV, and general spreadsheet editing. Enable implicit
  invocation after trigger trials (intended requests and near misses) pass.
- Checks added to `check.ps1`: frontmatter `name` matches the directory, and
  every relative link resolves inside the package.
- Evaluation in fresh contexts, recorded under `evals/skill/`: working CLI,
  missing CLI, `missing-cached-value`, boundary failure, partial batch, and
  near-miss requests. Author and evaluate with `/develop-skills`.

### 4. Documentation and release

- README: installation guide in mytech's order (prerequisites, obtain and
  install, `PATH`, verify, upgrade), the JSON contract, and how to obtain the
  skill separately from the executable.
- `PLAN.md`: milestones, new release artifact contract, validation evidence.
- Release 3.0.0, then verify the documented install sequence from an
  independent terminal, not only from the installing shell.

## Order

1. Process contract and in-process tests.
2. Standalone publish, artifact tests, installers, release workflow.
3. Skill package, checks, evaluation.
4. Documentation, 3.0.0 release, consumer-context install verification.

Steps 1 and 2 touch different files and can proceed in parallel; step 3
depends on the contract from step 1 and the install path from step 2.

## Unverified

- Measured: each compressed single-file CLI archive is about 34 MB. Native
  AOT stays deferred.
- How the skill should be registered with `sjskills`, if at all; the plan
  assumes consumers copy the `skill/excel-formula-remover/` directory.
