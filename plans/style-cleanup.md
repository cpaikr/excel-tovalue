# xlclean: rename and style cleanup

Status: steps 1 to 7 done on 2026-11-26, and step 8's documentation (README,
skill, evals) is prepared; the repository rename, tag, and install
verification of step 8, and the GUI (step 9), remain.
[PLAN.md](../PLAN.md) owns current state; this document owns the design and
order of the work.

## Goal

Extend the tool from one operation (formulas to cached values) to a second:
removing style bloat the way XLStylesTool does, so that a workbook that hits
"Too many different cell formats", opens slowly, or carries thousands of
copied cell styles can be cleaned unattended by a person or an AI agent. The
product is renamed `xlclean` because "formula remover" no longer describes
it.

The safety contract is unchanged and applies to the new operation: never
modify the source, never overwrite an output, transform a per-file temporary
copy, verify it, publish atomically, and fail the whole file with no partial
output when the tool cannot prove the result is equivalent.

## Decisions

Made by the owner on 2026-11-26.

| Topic | Decision |
| --- | --- |
| Cleanup scope | Remove unused custom named styles, remove unused cell formats, and merge identical cell formats. Excess-formatting reset is out of scope. |
| Formats | `.xlsx`, `.xlsm`, and `.xlsb` in the same release; `.xlsb` is its own milestone after the OpenXML formats. |
| CLI shape | Subcommands in one executable, released as a breaking 4.0.0. |
| Name | `xlclean` for the executable, product, skill, release assets, and GitHub repository (`cpaikr/excel-tovalue` becomes `cpaikr/xlclean`). |
| GUI | Unchanged in 4.0.0 apart from the rename; style cleanup reaches the GUI in a later release. |

Defaults chosen within that scope (change any of them before implementation
starts):

- **Built-in styles stay.** Entries with a `builtinId` (Normal, Good, Heading
  1, Comma, ...) are kept even when unused. Excel shows them in the gallery
  regardless, and a customized built-in may be intentional. Only custom styles
  are removed.
- **In-use styles stay, even when they duplicate another style.** Merging
  "Normal 2" into "Normal" keeps the appearance but changes the style name a
  cell reports, which is a visible change. Deferred until a real file needs it.
- **Component tables stay.** `numFmts`, `fonts`, `fills`, `borders`, and
  `dxfs` are not compacted. Pivot tables reference `numFmtId` directly,
  conditional formats and table styles use `dxfs`, and none of them drive the
  format-count limit. Two cell formats count as identical when their
  *resolved* formatting matches, so duplicate fonts or fills do not block a
  merge.
- **VBA warning.** Macros can apply styles by name (`Range.Style = "X"`), and
  the VBA project is not parsed. When a workbook with a VBA project loses any
  custom style, the file result carries a warning naming the count.
- **Over-limit warning.** [MS-XLSB] 2.4.878 caps the non-built-in cell and
  style format records together at 65,430 (`0xFF96`). When the cleaned
  workbook still exceeds that combined count, the result warns instead of
  failing.
- **Unsupported sheet types fail.** Macro sheets (`.xlsx` macrosheet parts,
  `.xlsb` macro sheet parts) hold cell tables with style indices. They are
  rare enough that the file fails with `unsupported-workbook-content` rather
  than gaining a third rewriter.
- **No-op still publishes.** A workbook with nothing to remove succeeds and
  publishes a copy, so every success has an `output` path.
- **No alias for the old name.** Existing `excel-formula-remover` installs
  keep working as 3.x; the README tells users to install `xlclean` and delete
  the old directory. No compatibility shim.

## What the cleanup does

Style indices in a workbook form a small graph. The cleanup computes what is
reachable from cells and rewrites indices to a compact numbering.

1. **Used cell formats.** Collect every `cellXfs` index referenced by cells
   (`c/@s`), rows (`row/@s`), and columns (`col/@style`) in every worksheet,
   plus index 0 (the default format). Everything else is unused.
2. **Merge identical cell formats.** Among the used formats, group those whose
   resolved formatting is identical: number format, font, fill, and border
   *content* (not their ids), alignment, protection, the `apply*` flags,
   `quotePrefix`, `pivotButton`, the parent named style, and any extension
   content. Each group keeps its lowest index. The comparison uses a
   canonical form of the XML that sorts attributes but never fills in
   defaults, so an absent flag and an explicit one differ; a missed merge is
   acceptable, a false one is not. Built-in number formats (ids below 164,
   several of them locale-dependent and absent from `numFmts`) are keyed by
   id; custom formats are keyed by format code. A built-in id is never equal
   to a custom code, even when the text looks the same.
3. **Used named styles.** A `cellStyleXfs` entry is used when a surviving cell
   format references it through `xfId` (an absent `xfId` counts as 0, as Excel
   reads it), when it is index 0, or when a kept
   `cellStyles` entry points to it. A `cellStyles` entry is kept when it is
   built-in (any `builtinId`, including Normal at `builtinId` 0), when its
   `cellStyleXfs` entry is used, or when a table names it: `.xlsx`
   `table/@headerRowCellStyle`, `@dataCellStyle`, `@totalsRowCellStyle` and
   the same attributes on `tableColumn`; `.xlsb` `BrtBeginList` and
   `BrtBeginListCol` style-name fields ([MS-XLSB] 2.4.100, 2.4.101). Every
   other custom style is removed.
4. **Renumber.** Remap `cellXfs/xf/@xfId` and `cellStyles/cellStyle/@xfId` to
   the compacted `cellStyleXfs`, then remap every cell, row, and column style
   index in every worksheet to the compacted `cellXfs`. Update the `count`
   attributes. Nothing else in `styles.xml` changes.

When no referenced cell format moves, only the styles part is rewritten and
worksheets are copied unchanged. Worksheets are rewritten node by node with
only the three style attributes changed, and tab, line feed, and carriage
return characters that a reader would normalize are written as character
references.

### Fail closed

The file fails with no output when the cleanup meets anything that may hold a
style index it does not rewrite:

- legacy shared-workbook revision parts (`revisionHeaders` and revision logs,
  whose cell records carry `s`);
- macro sheets (see Decisions);
- in `.xlsb`, any record inside sheet data whose type is not in the known
  set. The set is built from Excel-saved files before release, so it covers
  the records Excel routinely writes there (alternate-content and
  future-record wrappers, value metadata) as well as cells and rows;
- a style index outside the table it points into, a missing or malformed
  styles part, or a style graph that does not resolve.

New stable failure kind: `unsupported-workbook-content`, for the first two
cases. Malformed styles use the existing `invalid-workbook`.

### Verification before publish

The existing reopen check stays. The cleanup adds a verifier that is
independent of the planner: it does not reuse step 2's canonical key, so a
property the key forgets cannot both cause a wrong merge and pass
verification.

- **Formatting.** For every cell, row, and column, compare the source and
  output formats generically: the whole `xf` subtree, every attribute and
  child except the index fields, then the referenced number format, font,
  fill, and border elements, then the parent `cellStyleXfs` entry and its
  `cellStyles` name. Any difference fails the file.
- **Worksheets.** Stream source and output side by side. Apart from the
  style attributes, each sheet must be identical at the infoset level:
  every element, attribute (including `t`, `cm`, `vm`, `ph`), inline string,
  and extension, in order.
- **Styles part.** Every section other than `cellStyleXfs`, `cellXfs`, and
  `cellStyles` must be unchanged.
- **Other parts.** Content must be byte-identical after decompression.
  `[Content_Types].xml` and relationship parts are compared semantically,
  because the package library may re-serialize them.
- **Schema.** OpenXML validation of the styles part must report no finding
  that the source did not already have, compared as a multiset of (error id,
  part, element name) with no error cap. Positional paths shift after
  compaction, so they are not part of the comparison.

## Formats

### `.xlsx` and `.xlsm`

The styles part is read and rewritten with the Open XML SDK DOM, since it is
small. Worksheets are rewritten with `OpenXmlReader` and `OpenXmlWriter`
streaming, as the formula engine already does, changing only the three style
attributes. The VBA project, calculation chain, and calculation settings are
copied unchanged.

### `.xlsb`

The same algorithm on BIFF12 records, reusing `Biff12Reader` and
`Biff12Writer`. Records involved, per [MS-XLSB]:

- `styles.bin`: `BrtXF` (2.4.878) inside the `BrtBeginCellStyleXFs` and
  `BrtBeginCellXFs` sections (`ixfeParent` links a cell format to its style),
  `BrtStyle` (2.4.811) inside `BrtBeginStyles` (`ixf` links a style name to
  its format), and the count fields of each section's begin record. The
  Styles grammar (2.1.7.50) lets an `XFCOMPLEMENTID` group and other
  future records follow a `BrtXF`, and an `ACUID` precede a `BrtStyle`. These
  are kept or removed together with the record they belong to.
- Worksheets: the 24-bit `iStyleRef` of the `Cell` structure (2.5.10) in
  every cell record (`BrtCellBlank`, `BrtCellRk`, `BrtCellError`,
  `BrtCellBool`, `BrtCellReal`, `BrtCellSt`, `BrtCellIsst`,
  `BrtCellRString`, and the four formula records); the `BrtShort*` cell
  forms, whose layout is not in [MS-XLSB] rev. 23.0; `ixfe` in `BrtRowHdr`
  (2.4.772), which applies only when `fGhostDirty` is set; and `ixfe` in
  `BrtColInfo` (2.4.336). A row's `ixfe` is remapped only when it applies,
  and a stale value there is not a range error.

Confirmed against workbooks saved by Excel 16 (2026-11-26), recorded in
`XlsbStyleRecords`: section begin records carry a u32 count; each `BrtStyle`
is preceded by an alternate-content uid group (37, 3072, 38) that moves with
it; `BrtBeginList` style names follow the name, display name, and comment, and
`BrtBeginListCol` style names are its fourth to sixth strings. Excel wrote no
`BrtShort*` records, so they are not in the accepted sheet-data set and fail
the file. A rewritten worksheet loses its binary index part, which Excel
rebuilds. The `.xlsb` verifier compares records the same way the `.xlsx`
verifier compares elements.

## CLI 4.0.0

```text
xlclean                              root help (byte-identical to --help)
xlclean --version
xlclean values --output <dir> [--start-row <n>] [--exclude <pattern>]...
               <workbook>...
xlclean styles --output <dir> <workbook>...
xlclean <command> --help
```

- `values` is today's behavior with today's options; `styles` takes no scope
  options in this release.
- Output names keep the per-operation suffix convention: `Book_수식제거.xlsx`
  for `values`, `Book_스타일정리.xlsx` for `styles`. Chaining operations means
  running `styles` on a `values` output.
- The JSON envelope gains a top-level `command`. Each file entry keeps `input`,
  `status`, `output`, `failure`, and `warnings`, and moves the
  operation-specific counts into `changes`, which is `null` for a failed file
  because nothing was published:
  - `values`: `removedFormulas`, `removedTableFormulaDefinitions`,
    `processedSheets`, `excludedSheets` (today's fields, moved).
  - `styles`: `cellStyles` and `cellFormats`, each
    `{ "before": n, "after": n }`, plus `mergedCellFormats`.
- An unknown or missing command is a `usage` error with the same JSON shape
  and `command: null`; an option before any command (the 3.x form) gets a
  message that names `xlclean values`. Exit codes are unchanged.

## Code shape

- **Shared pipeline.** Extract from `FormulaRemovalEngine` the parts that are
  not about formulas: input validation, duplicate detection, output name
  allocation, temporary copy, reopen check, atomic publish, cleanup,
  cancellation, progress, and failure mapping. Each operation supplies a
  transform, a verifier, and an output suffix. Both operations then share one
  tested path for the safety guarantees.
- **Style model.** One reader per format produces the same immutable model
  (cell formats, style formats, named styles, resolved keys). The planner
  computes the index maps from that model. Writers per format apply the maps.
  Planner tests need no files.
- **Rename.** Solution, projects, namespaces, and test projects become
  `XlClean.*`; executables `xlclean` and `XlClean.exe`; skill directory
  `skills/xlclean`; release assets `xlclean-cli-<version>-<rid>` and
  `xlclean-gui-<version>-win-x64.zip`; install directory
  `%LOCALAPPDATA%\xlclean\bin`; the artifact-test variable `XLCLEAN_EXE`. The
  rule that `Core` and `Cli` do not depend on Windows APIs is unchanged.

## Tests and fixtures

- Generated workbooks for each rule: unused custom styles, used custom styles,
  styles named only by a table or table column, orphan style formats, unused
  and identical cell formats, formats that differ only in font id, a built-in
  number format next to a custom format with the same code, row and column
  formats, built-in styles, VBA warning, over-limit warning, every
  fail-closed case, cancellation mid-file.
- Property-style test: random style graphs and cell assignments; the
  cleaned workbook passes the verifier and has no unused or duplicate entries.
- Verifier tests that corrupt an output in each way the verifier claims to
  catch (a style index, a value, a cell attribute, one property inside a
  merged format, a dropped style name) and expect failure, so a broken
  verifier cannot pass silently.
- Excel-saved samples in both formats: a workbook bloated by copying sheets
  between workbooks, and at least one real bloated workbook from the owner.
- Excel desktop round trip before release: each cleaned sample opens without
  a repair prompt, looks the same, and shows the expected style gallery.

## Order

1. **Rename to xlclean** with no behavior change: projects, namespaces,
   executables, scripts, installers, release workflow, PLAN, CLAUDE.md, and
   the README's development sections. Merge to `main`; no release. The
   README's installation, usage, and skill sections and the skill itself keep
   describing the published 3.x release until step 8, so `main` never points
   at assets that do not exist. `npx skills add` installs the skill from
   `main`, so the skill directory moves to `skills/xlclean` only at step 8.
2. **Subcommands** with `values` and the new JSON envelope. The skill and its
   evals describe the released CLI, so they move to the new commands and
   fields at step 8.
3. **Shared pipeline extraction.** Existing tests pass unchanged in meaning.
4. **Style model, planner, and verifier** for `.xlsx`/`.xlsm`.
5. **`styles` for `.xlsx`/`.xlsm`**: named styles, then cell formats.
6. **`styles` for `.xlsb`**.
7. **Documentation**: README sections for `styles`, the skill's guidance on
   when to use it and how to read `changes`, PLAN validation.
8. **Release 4.0.0**: Excel round trip, rename the GitHub repository, switch
   the installation docs to `xlclean`, move the skill to `skills/xlclean` with
   its evals, tag, then verify the documented install from an independent
   process.
9. **GUI** (4.1.0): operation selector in the WinForms app.

Steps 1 to 3 change no behavior beyond the CLI shape and can be reviewed
mechanically. Step 6 depends on step 4's model; steps 5 and 6 touch different
files.

## Open questions

- Real bloated workbooks: the owner supplies one or more, or the samples are
  built in Excel only. Real files catch structures that generated ones miss.

## Unverified

- What Excel does when a table names a missing cell style (repair prompt or
  silent fallback). The plan avoids the case, so it only matters for
  diagnosing existing broken files.
- How Excel reads an absent `apply*` flag in `cellXfs`. The plan never merges
  across the difference, so the answer affects only how many merges are
  missed.
- The `BrtShort*` cell layouts and the meaning of `BrtXfComplementId`; both
  are confirmed from Excel-saved files in step 6.
- Whether saving through the Open XML SDK re-serializes
  `[Content_Types].xml` and relationship parts; the verifier compares them
  semantically either way.
