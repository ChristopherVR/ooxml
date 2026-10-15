# Excel parity review

Reviewed 2026-10-12 against `cd3915465` in the `ooxml-core` checkout. Reviewed 2026-10-15 for the
pivot gap list in the priority table.
This is a focused source review and implementation, not an Excel parity certification.

## Current PivotTable status

The dated source audit below records the initial implementation. Subsequent
native-backed batches support one or two real row fields, up to sixteen
measures, uniform tabular/compact/outline layouts and an optional column axis.
One report filter supports All or one exact cached source value when Values
are on columns. Both row levels retain complete cache item indexes; selection,
source refresh, history and save/reopen reuse the existing indexed engine.

The nested report-filter corpus contains 509 native cached/refreshed cases,
including 114 new cases and 108 owned native layout profiles. Multiple or
dynamic filters, Values-on-rows filters, grouping, slicers, shared caches,
mixed/custom layouts, ambiguous captions and repeated column captions remain
preserved unsupported profiles. Native English caption forms are measured.
See [the current suite evidence](office-parity-review.md) and
`native-pivot-nested-page-evidence.json` for bounds and reproduction scripts.

## Implemented

`src/core/xlsx/formula/functions/text-regex.ts` previously rejected every
`REGEXEXTRACT` return mode except the first match. It now returns all matches
as a vertical spill for mode 1 and the first match's capture groups as a
horizontal spill for mode 2. The existing calculation, editing, grid and save
pipelines handle those results. Leading zeroes remain text. Invalid modes and
patterns report errors; no match reports `#N/A`.

The function catalog now explains the return modes. Insert Function displays
that description and can find the function by "capture groups". Regression
tests cover insertion from the dialog, input edits, shrinking spills, undo
and redo, blocked destinations, displayed cell text and XLSX round trips.

The source-development aliases in `viewers/xlsx/scripts/local-core-aliases.ts`
now resolve `ooxml-ui/pptx/dom`. The Excel demo imports the shared UI root,
which also exports Teams presentation previews using that subpath. Without
the alias, an opted-in source build failed before browser tests could start.

Microsoft documents the three return modes in its
[REGEXEXTRACT reference](https://support.microsoft.com/en-us/excel/functions/regexextract-function).
Its illustrated first example confirms the vertical all-match layout.
The installed Excel version returned `#NAME?` for REGEXEXTRACT, so native
Excel calculation acceptance was unavailable. These functions still use
JavaScript regular expressions: PCRE2-only syntax remains unsupported.

## Remaining gaps and priorities

### Sparkline authoring follow-up

The follow-up adds core edit-session commands to create, format and clear
sparkline groups, with an Insert ribbon group and settings dialogs. Existing
extension XML remains the source of truth; edits retain unknown group fields
and unrelated extensions. Settings apply to whole selected groups and clearing
can target selected hosts or whole groups. Group creation maps data rows or
columns to a one-dimensional location range. It accepts local or sheet-qualified
A1 ranges; named ranges, external/3D references and merged hosts are rejected.
Creation is capped at 100,000 hosts per operation.

The native Excel acceptance script is `scripts/xlsx-sparkline-acceptance.ps1`.
It verified fresh line, column and win/loss groups, custom axis limits, RGB
color and high-point highlighting, clearing all groups, undoing that clear,
and adding a fifth group to the Excel-authored fixture. Excel silently ignores
the extension if its URI uses uppercase `4FD2`; the writer now uses Excel's
exact lowercase `4fd2` spelling. Adding to revision-tagged native groups also
requires a distinct namespaced revision UID, which the commands now provide.

Core preservation/edit/structural round trips, translated dialog/ribbon tests
and the production browser create/format/clear/undo/redo flow pass. Clearing
the last sparkline disabled the focused ribbon menu trigger, leaving focus on
the page body. Ribbon command completion now restores grid focus when focus
is still on the ribbon or has fallen to body; dialogs restore the grid or
their opener. Ctrl+Z and Ctrl+Y work after clearing without a manual click.
Style galleries, grouping/ungrouping, date-axis authoring and
individual source-range editing remain follow-up work. Existing date-axis
metadata is preserved.

| Priority | Gap                                     | Current evidence and next step                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| -------- | --------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1        | Pivot table grouping and cache scope    | Bounded pivot authoring and refresh now exist: `src/core/xlsx/pivot/` (aggregation, layouts, report filter, native cache XML) with `model-pivots.ts`, and the pivot dialogs in `src/ui/src/xlsx`. Number and date grouping (seconds to years), any number of row and column fields and per-field custom subtotals now round trip. Remaining gaps: shared or slicer-connected caches, date groups with report filters or timelines, custom subtotals with Values on rows, and column field subtotal settings. These stay preserved and read-only with explicit warnings (`pivot/admission.ts`). |
| 1        | Workbook import coverage                | `xlsx/load/detect.ts` identifies XLSB so it can be refused; ODS is not a supported format. Implement value imports separately from formula fidelity and expose the limitation on load.                                                                                                                                                                                                                                                                                                                                                                                                         |
| 2        | Financial and statistical function gaps | `viewers/xlsx/docs/features.md` lists missing functions including odd-period bond pricing, Bessel functions and FORECAST.ETS. Native Excel fixtures should establish edge behavior before adding them.                                                                                                                                                                                                                                                                                                                                                                                         |
| 2        | Sparkline authoring                     | Creation, group settings and clearing now work. Style galleries, grouping/ungrouping, date-axis authoring and individual source-range editing remain.                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| 2        | Print workflows                         | The supported print path is browser printing of the used range or print area. Page Layout and Page Break Preview are absent. Shared pagination models should precede these views.                                                                                                                                                                                                                                                                                                                                                                                                              |
| 3        | Complex drawing and chart authoring     | Advanced chart families and complete axes, labels and effects editing remain incomplete; most non-picture drawings use placeholders or cached renderings. Native fixtures are needed for each newly supported family.                                                                                                                                                                                                                                                                                                                                                                          |

External connections use cached values and are not refreshed. Macros and Power
Query are not executed. These are substantial product gaps beyond the focused
calculation improvement in this change.

## Reference review

The subsequent PivotTable batch implements range-based creation, settings,
refresh and deletion with one row field, an optional column field and one
Sum/Count/Average/Min/Max measure. It writes native table/cache parts, supports
undo/redo and exposes ribbon dialogs. The wider core suite passes 6,628 tests;
190 focused UI tests, the save/reopen browser workflow and eight native Excel
acceptance cases pass. Unsupported layouts, custom formats, shared caches,
slicers and unknown options remain preserved with explicit warnings. Paginated
printing and broader pivot field layouts remain work in progress.

The supplied Storytold SheetCraft URL redirects to
[GridCraft](https://github.com/storytold/gridcraft). Its README describes a
Rust engine separated from the UI, a command API shared by UI and automation,
dynamic arrays, pivot authoring and round trips, print previews, and limited
XLSB and ODS value import. These are useful workflow and architectural
references, especially pivots and print preview. README claims were not
independently certified against Excel, and no reference implementation was
copied into this change. The existing OOXML core, editing sessions and shared
editor remain the implementation path.

## Validation

- Strict core TypeScript check passed.
- Core regex, editing/grid integration and dynamic-array save/load tests passed.
- Insert Function dialog tests passed.
- Source-alias tests passed.
- Production-build browser regression passed on port 8133: matching results
  are painted, an input edit updates the result and removes old visible spill cells.

The browser test inspects visible cells because the grid pools hidden nodes
whose text remains in the DOM after their cells disappear.

## Subsequent native-validated workflows

Pivot authoring now supports tabular, compact and outline layouts, one or two
row fields, one value field and an optional column field. Automatic subtotals
can appear above or below compact/outline groups, or be hidden. Outline and
tabular layouts support repeated labels. The native corpus contains 46 cases
with cached/refreshed cell values, types, dimensions and compact indentation
checked. Custom or conflicting layout profiles remain preserved and read-only.
Multi-value layouts, grouping, slicers and shared-cache refresh remain gaps.

Paginated preview and print fields support saved page setup, header/footer
fields and rich formatting: font family/size, RGB color, bold, italic,
underline and strike. Native PDF checks and browser save/reopen tests verify
the supported styles. Theme colors, images, page offsets, double underline,
super/subscript and unknown font styles remain outside preview support.
