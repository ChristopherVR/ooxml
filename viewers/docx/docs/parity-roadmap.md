# Word parity milestones

Status reviewed against the current model, shared editor and regression tests on
2026-10-03. Reviewed 2026-10-15 for footnote continuation, nested tables, Update Fields,
chart painting and editable Print Layout pages. Microsoft Word equivalence is the target, not a capability claim for
this release. Import, visual layout, editing, export, preservation and accessibility
must each be demonstrated independently. Keeping XML does not prove that it renders
or recalculates correctly.

## Current foundation

One `<docx-editor>` implements the UI for six lifecycle adapters: React, Vue,
Angular, Svelte, Solid and Vanilla. All document logic belongs in `ooxml-core/docx`
and its layout/load subpaths; `docx-core` is a thin re-export. Shared drawing,
diagram, maths and collaboration logic belongs in the corresponding core areas.
`ole2` owns the legacy binary codecs. Fix a behavior once in the shared editor,
then exercise it through each adapter.

The current implemented subset includes:

| Area                   | Available behavior                                                                                                                                                                             | Evidence and remaining boundary                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| ---------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Text and formatting    | Editing, undo/redo, paragraph and character styles, theme fonts/colors, direct formatting, direction/language metadata, font scale/spacing/position/kerning and ligatures                      | `editor.spec.ts`, `advanced-font-preservation.spec.ts`, `font-advanced-controls.spec.ts`, core style/theme/toggle tests. Browser shaping and Word glyph metrics can differ.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| Lists                  | Bullets, numbering, multilevel definitions, heading-linked outlines, restart rules and marker alignment                                                                                        | `numbering-restart.spec.ts`, `heading-numbering.spec.ts`, `multilevel-list-dialog.spec.ts`, core numbering tests. Picture bullets and some style-linked/custom numbering remain limited.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Tables                 | Rectangular row/column commands, row height/repeat/split properties, default and selected-cell margins, cell borders and Cell/Table solid shading                                              | `table-properties.spec.ts`, `table-borders.spec.ts`, `table-shading.spec.ts`, `table-cell-margins.spec.ts`, core table preservation tests. Merged/nested/complex structural edits stay protected. No per-side border pens or pattern-fill editor.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| Sections and stories   | Per-section page setup and breaks, editable headers/footers/notes, first/even/default variants, Link to Previous, page-number fields                                                           | Header/footer context, variants, link and history browser contracts, core story creation/editing tests. Decorative galleries and full Word story layout remain missing.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| References and review  | Editable field results, TOC/table of figures, captions/cross-references, explicit Update Fields, comments and tracked-change commands                                                          | `fields.spec.ts` and shared-editor/core field, caption, comment and revision tests. Unsupported field switches and automatic recalculation remain limited.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Pictures and objects   | Pictures, supported floating-picture layout, hyperlinks/bookmarks, simple inline text boxes, SmartArt display/insert/node/text/layout/style/colour editing, chart and admitted ChartEx editing | Shared chart/diagram core and product browser tests cover history and save/reopen across six bindings. The Format pane exposes supported effects, shape 3D, text, pattern fill, analysis, markers and axis titles; picture media allocation, wider combo cases and native typography/effect fidelity remain limited. SmartArt uses the neutral interpreter with explicit fallback and conservatively generated 2D caches; Styles/Change Colors retain node overrides. Wider native styles, complex 3D/rich-text authoring and drawings in page stories remain incomplete. Native Word/Excel semantic reopen checks are scoped in the [shared delivery ledger](https://github.com/ChristopherVR/ooxml/blob/main/docs/office-shared-parity-delivery.md). |
| File and preferences   | File Home, Info/properties, New/Open/Save/Save As/Export/Print; display language, theme, review author, spelling/ruler/marks/thumbnails preferences                                            | `file-tab.spec.ts`, `document-properties.spec.ts`, backstage/chrome tests. Preferences apply to this editor session; recent files and account/cloud sharing are missing. Protection for filling in forms is enforced (`form-fields.spec.ts`); other editing restrictions are kept but not enforced.                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| Search and coauthoring | Literal Unicode-aware find/replace, authority-based step synchronization and peer cursors                                                                                                      | Search/controller and collaboration tests plus the local coauthoring demo. Networking, permissions, persistence and shared-save coordination belong to the host.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |

Five display locales are available: English, French, German, Spanish and
Simplified Chinese. This is interface localization, separate from document
language and spelling. Spelling uses the browser service; there is no bundled
Word-equivalent grammar or proofing engine.

Imported inline and display equations now render through native MathML using
`ooxml-core/math`. They remain display-only: the original OMML is preserved when
surrounding text is edited and saved. Unsupported equation previews show a visible
notice. Print Layout draws equations as SVG from the shared `ooxml-core/math` box model
(the `[Equation]` placeholder appears only when no drawn object is available); this does
not reproduce Word's equation layout.

Print Layout already paginates through `ooxml-core/docx/layout`, including
sections/columns, keep rules, table row fragmentation, repeating table headers,
footnotes, headers/footers, page fields and supported floating pictures. The main
editing surface remains continuous, and Print Layout pages are editable too: the caret,
selection and IME run through the editor's ProseMirror input and are painted on the pages
(`print-editing.spec.ts`, `print-cursor.spec.ts`). Headers, footers, footnotes and endnotes
are edited in place on the pages (double-click a header or footer; press a note's text),
and floating pictures are selectable there (see `outstanding-work.md`). Merged table cells
are read-only on pages, nested tables keep their authored grid, and caret placement at line
and page boundaries is heuristic. Pagination is approximate.

Selected-cell vertical alignment is editable with save/reopen and undo coverage
(`table-cell-alignment.spec.ts`). Toolbar/context-menu clipboard commands retain
supported HTML formatting and tables in the body and stories (`formatted-clipboard.spec.ts`);
plain-text-only browser APIs still work. External HTML does not guarantee preservation
of unsupported OOXML properties or objects.

## Next implementation priorities

This is the one live list of open work. [Outstanding work](/outstanding-work) records the
detail and limits of each feature already built.

1. **Editable pages and reliable layout.** Establish stable page/line/run mappings
   for caret placement, selections, scrolling, keyboard movement and IME before
   adding editing directly to pages. Improve footnote continuation, per-column
   notes, text-box layout, complex scripts and font substitution against references.
2. **Document preservation under richer edits.** Add merged-cell/table commands,
   remaining table/cell geometry controls only with imported-package surgery,
   identity-aware undo and save/reopen tests. Extend formatted clipboard coverage
   across paragraphs, cells and document stories.
3. **Drawing and embedded objects.** SmartArt, classic charts and admitted ChartEx editing
   run on the format-neutral core and shared `ooxml-ui` elements. SmartArt style/colour galleries
   and admitted generated 2D caches are implemented. Next are wider native styles, complex
   typography/3D, picture relationship authoring and shapes beyond preset geometry. Add each only after model, parser, layout and
   serializer contracts preserve the source package. Keep UI controls in `ooxml-ui`
   and framework adapters in the viewer.
4. **Word authoring breadth.** Extend style-definition editing, object positioning
   and wrapping, citations/bibliography/index, field recalculation and accessibility
   inspection. Unsupported imported content must remain visible or explicitly
   identified and preserved.
5. **Production review and integration.** Add proofing services, compare/protection
   and application persistence as explicit contracts. Complete cross-browser,
   keyboard/screen-reader, large-document and concurrent-editing coverage before
   advertising replacement-level reliability.

## Acceptance for Word equivalence

Build a checked-in corpus of Word-authored DOCX files with Word-rendered reference
pages and semantic expectations. Include ordinary business documents and difficult
cases: multiple sections, complex tables, linked stories, tracked changes, fields,
floating drawings, SmartArt/charts/equations, RTL/CJK, font substitutions and large
files. Cross-check with another renderer where helpful, but Word is the primary
reference for the stated target.

For each supported feature, require:

- Import diagnostics and expected model/text, including unsupported content.
- UI edit, undo/redo, selection and keyboard behavior in all six bindings.
- Export validation, reopening in Word and this editor, and checks that unrelated
  parts/relationships/content types and opaque source XML stay intact.
- Page counts, line/shape positions and visual comparisons with recorded
  tolerances. A passing serializer test does not establish visual parity.
- Accessibility, clipboard and IME coverage; browser and font versions recorded
  with visual references so differences can be investigated.

The Word reference corpus, how to regenerate it and its current results are in
[verification](/verification).

The existing Word COM checks documented in [outstanding work](/outstanding-work)
validate specific exported properties. They are useful evidence for those cases,
not a complete rendered-document comparison corpus or proof of full parity.

## Release gates

Run the repository CI checks, framework browser contracts and independent packed
consumer checks. Consume a published core release, restore any temporary `file:`
dependencies before committing, and reject copied core/binary logic. No unsupported
feature may disappear silently. Record known approximations in the UI/docs and
require evidence for each capability claim.

[Outstanding work](/outstanding-work) records detailed implemented behavior and its
limits. [Editing text](/editing) and [collaboration](/collaboration) describe the
current integration contracts.
