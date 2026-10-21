# Agent edits as a reviewable diff (`ooxml-core/automation`)

An agent's edits can be held back until someone has seen what they change. The agent proposes
edits; the host shows the diff; only an approved proposal is applied. Spreadsheets come first
(issue #156). The envelope is format-neutral so the other formats can follow with their own
edits and diffs.

```ts
import { applyXlsxProposal, proposeXlsxEdits } from 'ooxml-core/automation';

const { proposal, report } = await proposeXlsxEdits(bytes, [
	{ op: 'setCells', sheet: 0, cells: [{ address: 'B2', input: '=SUM(B3:B9)' }] },
	{ op: 'renameSheet', sheet: 0, name: 'Inputs' },
]);
console.log(report); // what would change, nothing is saved
if (approved) bytes = (await applyXlsxProposal(bytes, proposal)).bytes;
```

## Proposing

`proposeXlsxEdits(bytes, edits, { maxDifferences })` applies the edits to a copy of the workbook
through the core edit session (formulas are recalculated) and compares the copy with the original
using `compareWorkbooks` (see `docs/xlsx-compare.md`). It returns:

- `proposal`: `{ format: 'xlsx', base, edits, id }`, plain JSON a host can store or send back.
- `comparison`: the structured `WorkbookComparison`, bounded by `maxDifferences` (default 1000).
- `report`: `comparisonText(comparison)`, the same diff as plain text.
- `warnings`: loader warnings.

The diff is only as complete as the comparison: everything it does not look at (row heights,
comments, drawings and more) is listed in `comparison.notCompared` and at the end of the report.
Cells are matched by address, so inserting a row shows as a change to every cell below it.

## Applying

`applyXlsxProposal(bytes, proposal)` re-runs the proposal's edits and saves the result. It refuses,
changing nothing, when:

- `bytes` are not the workbook the proposal was made against (`base` is the SHA-256 of those
  bytes), so an approval never lands on a file that changed in the meantime;
- the proposal was changed after it was made (`id` is the SHA-256 of the format, base and edits,
  with object keys sorted).

The seal is a consistency check, not a signature: whoever can call `createEditProposal` can seal
other edits. A host that does not trust the agent keeps the proposal itself, between showing the
diff and applying it, rather than taking it back from the agent.

To accept only some of the edits, propose that subset again and review its diff: edits can depend
on each other (a formula reading a cell another edit fills), so a subset can change things the
full diff did not show.

## Edits

Edits are JSON objects run in order, each on the workbook the previous ones left. Sheets are
zero-based indexes, ranges are A1 references, rows are numbered from 1 and columns are letters, as
Excel shows them. At most 1,000 edits and 10,000 cell inputs per proposal.

| `op`                              | Fields                                 | Does                                                     |
| --------------------------------- | -------------------------------------- | -------------------------------------------------------- |
| `setCells`                        | `sheet`, `cells: [{ address, input }]` | Types each input into its cell (`=` starts a formula).   |
| `clear`                           | `sheet`, `range`, `what?`              | Clears `contents` (default), `formats`, `all` and so on. |
| `format`                          | `sheet`, `range`, `style`              | Applies a style patch (`numFmt`, `font`, `fill`, ...).   |
| `merge` / `unmerge`               | `sheet`, `range`, `mode?`              | Merges (`merge`, `center`, `across`) or unmerges.        |
| `insertRows` / `deleteRows`       | `sheet`, `row`, `count`                | Inserts above `row`, or deletes from `row` down.         |
| `insertColumns` / `deleteColumns` | `sheet`, `column`, `count`             | Inserts left of `column`, or deletes from it rightwards. |
| `addSheet`                        | `name?`, `at?`                         | Adds a sheet.                                            |
| `renameSheet` / `deleteSheet`     | `sheet`, `name`                        | Renames or deletes a sheet (one sheet always stays).     |
| `setName` / `deleteName`          | `name`, `formula`, `sheet?`            | Adds, replaces or deletes a defined name.                |

The first invalid edit throws `Edit <n>: <reason>` and nothing is proposed or saved.
`applyXlsxEdits(workbook, edits)` applies the same edits to a loaded workbook, for hosts that
already hold one (an editor showing the proposal live, for example).

## Other formats

`createEditProposal(format, bytes, edits)` and `checkEditProposal(proposal, format, bytes)` are the
format-neutral half. A format adds its edit type, an apply function and a diff. Word already has a
native review flow: `suggestDocxRunText` writes tracked changes the user accepts or rejects in Word
or the suite's editor.

## MCP

The tools belong to `viewers/xlsx/mcp` (see `docs/mcp.md`), which depends on the published core,
so they land after the core release that ships these functions: `xlsx_propose_edits` (read-only:
returns the report and the proposal) and `xlsx_apply_proposal`.
