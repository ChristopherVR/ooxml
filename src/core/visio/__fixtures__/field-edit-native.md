# Existing local field edits

`scripts/visio-existing-field-acceptance.ps1 -VerifyCore` creates three owned
Visio 16.0 plain local shapes, each with two fields and literal neighbors. The
first fields are numeric Width (picture 2), literal DATE (204) and PAGENAME (37);
the second field is Height (2). Core updates the first field to Width*2 with
two decimals, DOCCREATION with ISO date, or TITLE with text formatting. Native
open/save/reopen accepts all three outputs and retains their Value formulas.
Creation-date observations use E. Australia Standard Time (Australia/Brisbane).

The atomic command retains each fld IX, other Field rows, UI metadata,
Character/Paragraph markers, literal neighbors, unknown attributes/cells/XML
and all untouched package parts, including the native thumbnail. Only Value,
Format and Type caches/formulas/units plus the selected fld cache are changed.
No-op explicit edits return the original bytes. Stale source text/cache/row
proofs, self/Field references, dependent caches, locks, ambiguous markers,
grouped/inherited/foreign targets and nonstandard Calendar/ObjectKind refuse.

Native independent Field formula admission is confined to the measured
literal picture codes 2/37/204, literal integer DATE and existing direct
context functions. Unknown/dynamic pictures and dependencies still refuse.
This dependency proof does not add a formula evaluator or locale formatter.

The shared Field dialog exposes an Insert or edit choice, preloads the selected
formula and requires an explicit supported format when the original internal
picture is not writable. Shared UI and browser regressions use the reachable
ribbon, preserve atomic spans and verify undo/redo, export and reopen. Native
Field-dialog category/format UI fidelity, other internal picture codes and
complex/inherited/grouped field editing are not established by these probes.
New literal DATE authoring, including keeping an old literal DATE formula while
changing only its format, remains outside the insertion formula subset. The
measured date edit selects the existing supported DOCCREATION function instead.
Source edit proofs are omitted above 64 row cells or 2,048 proof characters;
those fields still render, but are not offered as editable targets.
