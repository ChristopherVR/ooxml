# Word parity review and implementation progress

Initial source review: 2026-10-07. Updated 2026-10-08 with the bounded native,
browser and regression evidence below. Current Microsoft 365 Word parity remains
unverified; the installed perpetual Word build is not subscription certification.

## Source state and ownership

The consolidated workspace uses `src/core` (`ooxml-core`) and `src/ui`
(`ooxml-ui`). Keep format logic in `src/core/docx`, format-neutral collaboration
in `src/core/collab`, reusable controls in `src/ui`, and framework lifecycle/event
adapters in the viewer. Core must never import the UI package.

The viewer entry re-exports `ooxml-ui/docx`. Six framework bindings consume one
shared editor. Editing is continuous and paginated Print Layout remains read-only.
Executed browser and native comparisons are recorded by feature below; those
cases do not establish pixel-equivalent output across arbitrary Word documents.

## Principal gaps

| Area                  | Evidence and remaining work                                                                                                                                                                                                                                                                        | Reuse boundary                                                                                                        |
| --------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Editable pages        | Editing remains continuous; Print Layout is read-only. Caret, selection, IME and keyboard mappings need contracts before editable page surfaces.                                                                                                                                                   | Shared docx layout; one Word editor                                                                                   |
| Layout                | Covered continuous-section bands, column balancing, headers and page fields have native/browser references below. Complex typography, floats, boundary transitions and Word font metrics remain incomplete.                                                                                        | Injectable measurer and shared layout engine                                                                          |
| Tables                | Complex/merged structural editing and collaborative table changes remain limited. Contiguous header pagination is fixed; complex structural editing and native table-layout coverage remain incomplete.                                                                                            | Word model, editing and layout in core                                                                                |
| Drawing objects       | The viewer roadmap identifies missing general shape, chart and SmartArt editing. Equations are display-only in the newer editor.                                                                                                                                                                   | Geometry, diagram and future shared DrawingML/chart/math areas, rather than imports of PowerPoint UI internals        |
| Proofing and review   | Browser spelling is not a Word grammar engine. Compare, protection, richer references and automatic field calculation need explicit implementations.                                                                                                                                               | Product logic in docx; host service contracts where appropriate                                                       |
| Collaboration         | Word uses authority-ordered ProseMirror steps and separate transient presence. Shared collab already has Yjs, awareness, WebSocket protocol and external WebRTC/WebSocket adapters. Opt-in Word Yjs binding is now wired through the shared session; multi-story collaboration remains incomplete. | Format-neutral lifecycle/providers in collab; Word schema mapping in docx; cursor DOM and Share/status controls in UI |
| Save and preservation | The model alone does not carry the loaded package's opaque parts or media. CRDT synchronization of model JSON cannot establish preservation or shared export reliability.                                                                                                                          | LoadedDocument/package preservation and assets in core                                                                |

## Implemented first slice

- Table pagination repeats only the contiguous header group at the start of a
  table. A later isolated `tblHeader` flag is ignored, including its implied
  non-splitting treatment; explicit `cantSplit` remains effective. This follows
  [Microsoft's documented tblHeader semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tableheader?view=openxml-3.0.1).
- External Yjs providers decode boolean and object sync events correctly, clear
  sync on WebRTC disconnection, expose explicit disconnect immediately, recognise
  initially connecting sockets, and guard repeated destruction/later actions.
- Sessions recognise providers that completed synchronization before attachment.
  Read-only bindings adopt an already-synced room without publishing their
  bootstrap content.
- Sessions expose `reconnect()` and capability-checked `resync()`. The stock
  transport resync repairs dropped remote and local updates without leaving the
  room. These actions are reusable by every Office format.

Regression coverage exercises real Yjs documents, dropped messages, peer counts,
retained document identity, read-only adoption, external provider event shapes,
and multi-page table fragments. These are core behavior tests, not Word-rendered
visual comparisons or production networking tests.

## Implemented Yjs editor slice

The shared Word editor now exposes `startYjsCollaboration(session, options)`,
`reconnectCollaboration()` and `resyncCollaboration()`. The Word mapping lives
in `src/core/docx/ui`, over the stable Yjs 13 ProseMirror binding; provider
lifecycle, identities, sanitization and asset routing are reused from `collab`.
The shared asset router now handles raw binary payloads without changing its
existing string/version-counter contract. The UI reuses localized cursor DOM
and the existing image cache. No Office logic was added to framework adapters.

Body text, formatting and revision marks synchronize through a Y.XmlFragment;
root page and section attributes use a separate map. Bootstrap requires actual
provider synchronization, matching source package identity and one designated
creator. Read-only peers cannot publish local document transactions. Local undo,
relative cursor positions and history across detach/remount are covered. New
picture parts synchronize separately and survive export and stopping while
unmounted. The host retains ownership of the provider/session.

The browser convergence and export contract passed in all six framework mounts.
Core and UI regressions also cover simultaneous first insertions in empty
paragraphs, formatting, page settings, recovery, loaded opaque parts and new
picture bytes. This is bounded evidence, not a complete M365 comparison. The
matching loaded package remains necessary for styles, notes and
existing assets. Editing outside the body and structural table commands remain
disabled; canonical multi-author saving, authorization and persistence remain
host responsibilities. See `viewers/docx/docs/collaboration.md` for the API and
limitations.

Comment editing commands now live in core and the UI re-exports them. New Yjs
rooms share independent anchor attributes and comment records, with separate
resolution and deletion maps. Adding or deleting a thread and its anchor uses
one local undo operation. Concurrent replies survive synchronization; deleted
roots hide concurrent replies until undo restores the root. Imported overlapping
anchors normalize before bootstrap, and deletion formats only the relevant Yjs
attribute to avoid reasserting another author's concurrently deleted anchor.
Legacy rooms without the comment capability and authority-step rooms keep comment
editing disabled. Tests cover partitioned edits, imported threads, author undo,
read-only peers and detached DOCX export, plus the review pane in six browser
bindings. This does not implement M365 mentions, notifications or assigned tasks.

The Track Changes recording setting now uses document transactions, shares
through Yjs, supports local undo/redo and exports `w:trackRevisions`. Recording
and revision resolution live in core and use the caller's schema. Revision IDs
reuse the client identity generator. Resolution separates legacy IDs by author,
ignores stale ranges, blocks read-only writes and bypasses recording, as do
undo/redo transactions. A two-peer partition/reject regression retains the other
author's insertion. Browser review contracts pass in all six bindings.
Run-formatting recording now retains prior snapshots through supported mark
transactions. Paragraph-mark recording still needs implementation and native evidence.

Move names now reuse shared identity generation during editing, with numeric
names assigned only on export. A fixed-clock regression covers 101 independent
moves. Native Word opened the corrected core export, saved both linked moves
and accepted or rejected the first without resolving the second. Moved-from runs now retain
`w:t`; the earlier `w:delText` export caused Word to reject the package as corrupt.
The four DOCX fixtures and native build/revision report are kept in
`src/core/docx/__fixtures__/review-moves`. This is desktop-build interoperation
evidence, with current M365 subscription validation still required.

## Original formatting display

Shared core projections now read prior run and paragraph properties through the
same restoration logic as Reject Changes, without modifying the source model.
Missing, invalid or unsupported snapshots return a diagnostic and retain current
formatting. Native Word rejected references cover four paragraph cases and two
run cases. The body editor uses paragraph projections in Original mode, including
prior paragraph style inheritance, numbering inputs and removal of current direct
spacing, indentation, direction, borders and shading. Display changes preserve
the authoritative model, revision records, text positions, selection and undo
history. All six browser bindings match the native before document's paragraph
appearance and retain pending revisions on export. The editable body also projects
prior run properties. Direct mark views retain their authored document identity
while deferring appearance to the shared resolved formatting. This removes current
underline, highlight and script positioning without importing display properties
into text edits. Native bold and combined font references match the before
documents across all six bindings. Display modes remain local to each Yjs peer
and do not create shared undo entries. Continuous header/footer/note previews and
their in-place editors now use the same projections and review visibility policy.
Mode changes preserve an active story's selection, source properties and history;
closing it retains the selected display mode. Preview paragraph attributes reuse
the complete core conversion instead of a separate alignment-only mapping.

Inline run conversion now lives in core and takes the caller's schema. Pictures
and page/column breaks retain run properties and revision history through editor
conversion, peer synchronization and export. Review visibility covers inline
atoms as well as text, using the same core policy as layout. Browser checks cover
pictures, breaks, equations and note references across all six bindings. The
visibility rule overrides node-view display styles. Shared editor Review commands
now enumerate and resolve inline-object text revisions, including linked moves,
isolated undo/redo, read-only access and reused IDs from different authors. Yjs
peers converge on acceptance/rejection and local undo. Ten native Word comparisons
cover body picture/note insertion/deletion and page-break deletion. Desktop Word
reopened all ten editor exports with matching body content, object counts and
pagination and zero body revisions. Follow-up note snapshots now let Accept All
and Reject All resolve retained footnote/endnote stories in the same transaction
as the body. Removed references also remove their note entries on export; Word
reopened the ten exports with zero body and footnote revisions. One local or Yjs
undo restores both the reference and note content. Note insertion includes its
content in the reference's history, note edits update the stored snapshot, and
history/provider updates refresh previews. Existing unreferenced notes remain
untouched unless their own revisions are resolved.
Note preview and editing identities include the note kind, so a footnote and an
endnote sharing the same numeric ID retain independent content during editing
and review-mode refreshes. Component and six-binding browser checks cover this.
Picture parsing now carries its run property basis; removing a picture may remove its own opaque properties without
weakening the guard for unsupported retained text or copying picture properties
onto neighboring text. Native page-break insertion also revises paragraph marks.
Shared review commands now enumerate, navigate and resolve inserted/deleted
paragraph boundaries, retain the following paragraph's formatting when merging,
and resolve consecutive boundaries from the end in one undoable transaction.
The merge rule reuses the core model command and follows the documented
[WordprocessingML paragraph-mark behavior](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document).
Native insertion accept/reject exports match Word's paragraph content and page
counts, and all twelve object-review exports reopen with zero body and note
revisions. Yjs peers converge and restore the original structure with one undo.
Six-binding browser checks cover acceptance, rejection, export and undo.
Removal across table/section boundaries or at the final paragraph is guarded;
display projection and recording of paragraph-mark revisions need further work.
Document-wide review now also resolves header/footer snapshots, including tables,
with the same core commands and history. Header/footer-only changes enable the
All actions. A native five-story fixture compares body, header, footer, footnote
and endnote formatting acceptance/rejection, and both editor exports reopen in
Word with matching story text and zero revisions in every inspected story.
Local and Yjs undo restore all five stories together; six browser bindings cover
resolution and preview refresh. Unsupported boundary removal fails before
dispatch, preserving the body and history. Shared review also resolves run-format
history on pictures, page breaks, note references and field markers/codes. The
resolver uses the same complete prior-property restoration and run conversion as
text; it retains independent insertion history, links, comments and note labels.
Native Word fixtures cover four object types and eight accepted/rejected exports.
Word reopens every export with native story text and zero inspected revisions.
Picture media bytes, drawing XML and opaque run properties survive resolution.
Pictures also retain their source drawing when preceding text is split or removed.
Changing the number of pictures sharing one media relationship is guarded until
the model carries a stable identity for each drawing.
Component history, Yjs peer undo and all six browser bindings cover these cases.
Continuous editing views now also project prior properties on inline atoms and
hard breaks through the existing shared run-style resolver. Node decorations
retain note labels, picture dimensions, source positions and undo history.
Six-binding browser checks compare the four native object cases with their
before references and retain the source model through display-mode changes.
Unavailable snapshots report diagnostics and retain current formatting.
Pure supported atom-format changes now record complete prior run properties
through the existing tracking plugin. The shared style-aware toggle command
lives in core and includes atoms in mixed text selections. Canonical atom
attributes retain formatting and history through Yjs, whose element mapping
does not retain ProseMirror marks. Native comparisons cover picture, note,
page-break and field-code recording, acceptance/rejection, earliest snapshots
and full reversion. Word recognizes all twelve recorded/resolved exports and
its Reject All restores the expected direct bold properties with zero revisions.
Peer tests retain the editing author and isolated undo; six-binding ribbon
checks cover mixed text/picture and text/note selections. Hidden field-code
formatting is tested through the command path because editing selections trim
hidden markers. Mixed structural transactions, OMML formatting, all advanced
atom formatting controls and hard-break formatting in Yjs need further work. These checks do not
establish complete native Word object-review parity.

Print Layout now projects prior run and paragraph formatting through the same
core helpers, including table cells, section stories and notes. Native before
references cover bold, combined run properties and combined paragraph properties.
The projection retains text offsets for click-to-cursor mapping and reports
unavailable snapshots through layout diagnostics and print warnings. The shared
layout API now accepts review display options: Original hides insertions and move
destinations; No Markup and Simple Markup hide deletions and move sources. Hidden
runs retain their UTF-16 source length for following fragments, including one
position for hidden inline atoms. Hidden breaks and pictures consume no layout
space, and note numbering uses only visible references. Headers and footers use
the same visibility policy, including floating pictures and watermarks. Native
move references pass across all six browser bindings, including click-to-cursor
and retained revisions on export. Print Layout still retains revised paragraph
marks and reports that display-time merging remains unsupported. Structural review
display still need implementation. These checks establish the covered formatting
semantics, not pixel parity with current Microsoft 365 Word.

## Next implementation sequence

Native bold and combined font-change references now retain the full prior
`w:rPr` snapshot through parsing, text editing, editor model conversion and
export. Core rejection restores the full prior properties, including opaque
XML, and matches Word's rejected documents. Word opens both package-preserving
and standalone rejected exports with the expected formatting and no revisions.
Accepting a run-format revision clears the snapshot. Shared editor Review
commands now navigate, accept and reject imported run-format revisions,
including undo/redo, peer synchronization and export. Supported run-formatting
changes and pure paragraph-format changes now record revisions. Original display
projects prior run and paragraph formatting in the body editor and Print Layout.
Print Layout also filters text and move revisions by review mode. Continuous story
previews and editing views now project prior formatting too. Revised paragraph-mark
merging for display remains unfinished. Shared Review commands
now navigate, accept and reject imported paragraph changes, with undo/redo,
Yjs peer synchronization and package export. Core
paragraph rejection now restores the full prior properties and matches all four
native rejected documents. Imported
paragraph snapshots now survive text editing, editor conversion and both export
paths, with four native references for alignment, spacing, indentation and
combined changes. Paragraph mark revisions also survive text edits. The references
were generated by the same installed desktop Word build described below.

Native runs with unmodeled properties still reject export when an edit changes
their boundaries. Tracked typing can trigger this guard, including in Yjs rooms.
Granular preservation of those properties across split and merged runs remains
required; package-preserving text edits with unchanged run counts are covered.

Continuous-section pagination now shares a physical page after a single-column
section, including changed margins and a new column count. Each section gets
its own column band; overflow adopts the new section's page geometry. Empty
continuous-break markers do not add a printed blank line. Page size and
orientation changes still start a new page, matching native Word measurements.
Print Layout limits column separators to their bands and retains per-section
line numbering and section-page counts.

Fourteen committed native DOCX references record paragraph page numbers and origins
from desktop Word `16.0.20430.20140`, with Arial 12 pt and exact line spacing.
The installed licenses are 2021/2024, so this is explicit desktop Word evidence,
not a current M365 subscription certification or a glyph/raster comparison.
Equal and unequal-width paragraph columns now balance their final page before a continuous
break, reusing the same paragraph flow, widow/orphan and keep-together rules.
The native corpus verifies even and odd distributions, earlier-page overflow,
keep-with-next groups and reversed unequal widths. Floating pictures and
explicit breaks still report unsupported balancing. Vertical alignment changes
and shared-page footnote cases also retain reported approximations. Floats
relative to changed margins and remaining complex layouts need native references.

Six additional native table references cover four, five and 120 exact-height
kept rows, multiline splittable rows and repeated headers before a continuous
break. The table flow now balances those rows on
the final page and retains full capacity on earlier pages. Word's mandatory
paragraph after the table adds its own line after balancing, unlike an empty
break marker after ordinary text. Trial validation checks earlier columns so a
kept row cannot silently overflow a reduced capacity. Core and six-framework
browser tests compare every nonempty row and following paragraph origin.
Native PDFs also establish header repeat counts in each column on both pages.
Trials that cannot fit one table line make progress with a reported oversized-row
fallback rather than looping. Complex row splits, nested tables and merged-cell
layout still need broader reference coverage.

Twelve additional native DOCX/PDF references cover visible headers, footers,
PAGE and SECTIONPAGES values across shared section pages. Page-number restarts
now survive continuous transitions, including odd/even-header parity on both
odd and even physical starting pages. Roman labels use numeric page facts to
select header slots. First-page headers and section page counts match the
recorded references. DOM-free page-field logic lives in core; the shared UI
renders it through compatibility facades. These semantic references do not
certify font metrics, header/footer geometry, all numbering formats, or M365
subscription behavior.

1. Extend collaboration to note content, headers/footers, style and
   numbering definitions, using granular mappings and explicit conflict rules.
   Add granular table transactions before enabling structural table editing.
2. Define full package bootstrap, authenticated provider permissions,
   persistence and canonical export contracts. A client role or document ID
   is not server authorization or proof of matching source bytes.
3. Establish a corpus authored and rendered by current Microsoft 365 Word.
   Record the Office build, fonts, semantic expectations and measured layout
   tolerances. Import/edit/undo/export/save-reopen and all six bindings must be
   checked for every feature.
4. Finish continuous-section balancing and complex page transitions, and develop editable pages with
   caret, selection, IME and keyboard contracts.
5. Fill DrawingML, chart, SmartArt, equations, proofing, references and protection
   gaps through shared core areas and one shared Word UI.

Complex-script font family, size, bold and italic, plus the East Asian font family,
now have shared parsing, style resolution, editor preservation and serialization.
The native paragraph-format fixtures support tracked text splitting and peer
export without dropping these properties. Desktop Word reopened all four exports
with their paragraph-format and insertion revisions, retained Arial for complex
scripts, and rejected both revisions back to the original text. All six browser
bindings exercise tracked and untracked typing. Script shaping, font fallback and
glyph metrics still require separate rendering parity work.

The parser now retains a complete property basis for runs containing unsupported
current formatting. Text splits can reuse that basis, including font hints,
extension properties and unmodeled formatting-history attributes. Each rebuilt
run imports its own basis, so changed run positions do not copy another run's
opaque properties. Known object-valued properties compare by value before XML
rewriting. Core package/standalone export, Yjs peer history and all six browser
bindings cover this path. Missing bases and mixed segmentations remain guarded;
this is not a claim that opaque formatting can be rendered or freely edited.

Text runs now preserve formatting history independently when they are inside
insertion, deletion or move wrappers. Both histories receive valid export IDs;
resolving formatting retains the pending text revision. Modern Word UTC timestamps
are preserved separately from legacy revision dates. Core, editor, Yjs and six
browser bindings cover overlapping tracked typing. Desktop Word reopened four
exports combining run formatting with insertion or deletion and rejected them
back to native baseline text and formatting. Original mode now displays prior
properties in the body editor and Print Layout while retaining both histories.

The shared Track Changes plugin now records supported run-formatting mark edits,
including the explicit-off properties used by UI toggles. It reuses shared mark
conversion and the ordinary OOXML writer for complete prior snapshots, retaining
script fonts, opaque source properties, text revisions, links and comments.
Successive formatting retains the earliest snapshot, reverting to that snapshot
removes the revision, and undo/redo includes both formatting and history. Native
Word action references cover bold, bold followed by italic, reverting bold and
formatting a pending insertion. Desktop Word reopened all eight package and
standalone exports with matching revision types and Reject All results. These
references do not establish current M365 certification, multi-author attribution
rules or mixed structural/formatting transactions.

Word's document-level Track Formatting and Track Moves preferences now parse
and export their negative settings flags. The editor retains them as shared
document attributes and core supplies isolated undoable toggle commands.
Disabled formatting tracking keeps formatting edits untracked while text tracking
continues; disabled move tracking keeps drag and cut/paste revisions as ordinary
insertion/deletion pairs. Native preference references and package/standalone
round trips cover the flags. A localized Tracking options dialog now exposes
Track Formatting and Track Moves through the existing shared dialog shell.
Both changed preferences apply as one undo operation; untouched fields retain
concurrent peer updates. Other advanced tracking and markup options remain unfinished.

Core revision enumeration and resolution now include headers, footers, footnotes,
endnotes and table cells through the shared document-story traversal. Synthetic
package tests verify accept/reject/save/reopen, including headers shared across
sections. Paragraph-mark merges stay within their story or cell. Failed Reject All
operations throw and preserve the input model and pending revisions. These core
commands do not enable collaborative editing of non-body stories or certify native
Word behavior for all story revisions.

Full M365 Word parity remains unfinished and must not be claimed without this
reference evidence.

Line breaks now use the shared core schema with run-property attributes, so Yjs
retains their formatting, language and imported history. Formatting commands,
recording and accept/reject reuse the same code as other supported inline atoms.
Tracked insertion, deletion and linked cut/paste moves retain line-break history;
deleting the author's own pending insertion removes it. Core regressions retain
compatibility with older mark-based hard-break schemas. Peer tests cover exports,
resolution and isolated undo/redo, and the six browser bindings cover recording
and Original display against native before documents. Word 16.0 build 20430
reopens the three recorded exports with the expected bold property and revision
counts and rejects tracked formatting to the baseline. This reference proves
those bounded semantics; current M365, layout and general atom mark parity remain
unfinished.

Direct font properties, highlights, script positioning and Font-dialog advanced
properties now use one core patch command for text and supported inline objects.
Selection readback includes objects instead of substituting the caret's font.
Clear Formatting retains hyperlinks, comments, field identities and tracked
text while clearing modeled direct run properties. Native references establish
size, color, small caps, spacing, scale, baseline position and kerning on pictures,
note references, page/line breaks and field-code runs. Core accepted/rejected
exports match their native references; Word reopens fifteen exports and native
Reject All restores baseline stories and properties. Font-dialog peer tests
verify shared attribution, atomic undo and resolution. Browser tests cover the
dialog and export across the six bindings. Formatting general shapes and OMML,
concurrent review attribution on the same atom, and current M365 certification remain
unfinished.

Independent direct properties on the same inline atom now merge through Yjs,
including removals, without replacing the full imported property JSON. Local
undo retains the other author's properties. The shared schema uses v2 rooms;
codec guards reject mismatched schemas and rooms, while matching legacy schemas
retain v1 support. Migration requires export with the old client and initialization
of a new room; provider persistence is not rewritten. DOM serialization retains
effective properties and note/field metadata. Eight disconnected-peer tests,
five DOM round trips and native export/rejection checks cover this change.
Concurrent review attribution still needs implementation.

Paragraphs now carry their complete source property basis through editor and
Yjs conversion. Rejection preserves current paragraph-mark and section properties
when the prior paragraph snapshot omits them, while restoring the paragraph
formatting that snapshot records. Raw-property regressions cover opaque
attributes, borders and section metadata. Desktop Word reopened all eight native
paragraph rejection exports with zero revisions and the expected paragraph-mark
fonts; all six browser bindings cover Review rejection and undo.

The shared Track Changes plugin now also records pure paragraph-format changes,
including alignment, spacing, indentation and pagination attributes. Snapshots
use the shared paragraph conversion and writer, excluding independently tracked
paragraph-mark and section properties. Successive edits retain the prior snapshot;
full reversion clears history. Empty paragraphs, per-paragraph snapshots, ordinary
undo/redo, Yjs peer history and six browser bindings are covered. Desktop Word
reopened all eight newly recorded package and standalone exports and rejected
them to the native baseline. Mixed text/structural transactions and multi-author
format attribution remain outside this slice.

Style-inherited `pageBreakBefore` now resolves through paragraph style ancestry
and layout, while explicit false cancels the style. The editor uses null for
inheritance and retains false through model, DOM and OOXML conversion. A native
three-paragraph reference establishes page assignments 1, 2, 2; core and all six
browser bindings reproduce them. Word reopens the edited package export with
the same pagination and explicit off override. This reference establishes page
break semantics, not general font-metric or pixel-equivalent pagination.

The same native fixture revealed that ordinary calculated pagination cache
markers blocked edits. Parser and writer now share an empty-cache-marker check;
edited runs invalidate the cache without treating it as an authored break.
Unknown inline extensions stay guarded and untouched package saves preserve
their bytes. Yjs tests cover a tracked explicit off override, export and shared
rejection back to inheritance. Word recognizes that export's paragraph revision,
places all paragraphs on page 1, and restores pages 1, 2, 2 on rejection.

Hyperlink selection and mutation now reuse shared core commands. Imported
picture targets and edits to supported inline elements survive Yjs joining,
retargeting, removal, clipboard conversion and local undo. Six browser bindings
verify peer picture-link export without losing shared media. Equation links and
hyperlink change attribution still need separate implementation and evidence.

Comment navigation now reads imported inline run anchors together with marks,
so pictures, breaks, note references and field markers can be located through
Review navigation. Core tests cover five kinds and overlapping anchor IDs;
mounted view tests cover four kinds. This read-path correction does not yet
establish concurrent creation/deletion of comments anchored only to elements.

V3 Yjs rooms now store element comment ranges independently per thread using
shared relative positions. Five core cases cover pictures, line/page breaks,
note references and field markers, including concurrent creation, deletion,
local undo, adjacent typing and selected-picture deletion. Mounted editor tests
verify immediate model updates as well as saved anchors. Six browser bindings
exercise concurrent picture comments, export and ribbon undo. Appended review
transactions now refresh the public document model, and deleted node selections
fall back to a cursor before the Yjs view updates.

Desktop Word reopened fifteen comment exports (two threads, deleted threads,
and one author's restored thread for each element kind). Counts and restored
authors match. The second overlapping field-marker comment has a zero-width
native scope, so field-range parity remains unfinished. The installed perpetual
Word build does not establish current Microsoft 365 parity. Existing v1/v2 rooms
require export with their matching client and creation of a fresh v3 room.

The field-comment scope gap above is now corrected for balanced complex fields.
Native Word authoring expands a begin-marker selection, a one-character result
selection and a whole-field selection to the complete field. Ordinary and Yjs
comment authoring now share that range expansion. Three native packages retain
both overlapping anchors through a neighboring text edit; rebuilt field exports
reopen with matching full scopes, deletion and local restoration. Five mounted
peer cases now include field results. Six browser bindings cover result selection,
whole-field export and undo. Nested fields, simple fields and M365 subscription
build comparisons still require separate evidence.

Anchor projections use the existing binding mutex to remain read-only during map
notifications. This prevents an older view snapshot from erasing incoming text
marks when a comment spans text and inline elements in the same transaction.

Ordinary comment undo currently removes range anchors but leaves thread metadata
in the pane; Yjs undo restores both. The field browser cases verify anchor undo
only. Ordinary thread metadata history is a separate remaining gap.

Ordinary thread history now corrects that gap: creation and deletion put records
and anchors in one document transaction, and replies/resolution have separate
undo events. Core history tests cover text and pictures; a mounted pane verifies
all four actions, saved restoration and redo. Six field browser cases now verify
that undo removes both the thread and its package part, and redo restores the
pane. Stopping Yjs and remounting retains shared threads for subsequent ordinary
history. Local snapshots are excluded from shared Yjs document attributes, which
continue to use independent thread maps. Modern comment notifications and task
assignment remain unfinished.

Provider transactions that change body content and root document settings now
project settings under the existing binding mutex. This prevents a settings
notification from writing an older body snapshot back before the body observer
runs. Regressions cover both notification orders, editing/read-only peers and
local undo, plus mounted public-model and saved-package checks. The mapping is
split into a small core module; provider lifecycle remains format-neutral.

### Simple-field comment scope and shared result scanning

Comments selected within a simple field's text result expand to the full result
across formatting splits. Whole-field comment edges enclose the simple-field
wrapper on export. Native Word 16.0.20430 expands result-character comments to
include the code (scope 6:29 in the synthetic QUOTE fixture); previously our
export covered only displayed text (23:28). The field update command now reuses
the core result scanner. Adjacent simple fields with identical instructions,
non-text results and nested field behavior still need separate evidence.

Simple-field imports also reuse the paragraph range/revision collector instead
of flattening only direct runs. Internal partial comment scopes and overlapping
outer comments survive parse/save; comment-reference runs no longer become
empty field results. The writer retains those imported partial scopes.

### Adjacent imported simple-field identities

Distinct imported simple fields with identical instructions retain identities
through the model, existing opaque editor marks, direct formatting, Yjs and
export. Comment expansion and field updates respect those boundaries; the first
field's scope matches native Word (6:29) and excludes the second field. This
addresses the imported adjacency gap noted above. Identity is regenerated from
source order on import, not written as an OOXML attribute. Clipboard copies and
new adjacent fields lacking an identity are not established by these cases.

### Field-result clipboard projection

The shared field guard projects copied/pasted result-only field marks to literal
text while retaining formatting and removing boundary identity. Complete complex
field representations keep their markers/results. Native Word formatted-range
transfer records three fields for a whole-field copy between two existing fields,
and two fields for full-result or one-character copies; copied bold survives.
The core helper is shared by both clipboard hooks, with mounted history and Yjs
convergence/export cases. Native range transfer corroborates these semantics;
full Microsoft 365 clipboard formats, whole simple-field copying in the editing
surface and pasting within field results still need separate coverage.

All three clipboard-hook/model-adapter exports reopen in Word with the same
field counts, complete body text and copied bold formatting as their native
range-transfer references. `write-word-field-copy-exports.mjs` and
`check-word-field-copy-exports.ps1` reproduce this bounded comparison.

### Replacing cached field results

Typing and inline text paste now retain the instruction and identity of an
existing simple field when replacing part, all, the first character or the last
character of its result. Pasted text retains its formatting and adopts the target
field identity. Boundary cursors and closed paragraph or non-text slices remain
outside this adoption. The implementation reuses the core field guard and result
scanner; Yjs receives the same authored marks without a separate mapping.

Six browser bindings cover typing, HTML paste, saved field counts and local undo
for simple and complex results. Mounted Yjs cases cover result replacement,
convergence, export and undo with one or two imported simple fields. Eight editor
hook/model-adapter exports reopen in Word with matching body text, field codes,
cached results and bold formatting against four native replacement references.
Word 16.0.20430 converts those imported simple fields to complex fields on save.
Deleting an entire simple result, non-text/structural paste and current Microsoft
365 subscription certification remain unfinished.

Backspace/Delete now retain the instruction when removing the whole simple
result or its final character. An empty result uses the existing complex-field
markers, which native Word also writes when saving the empty imported field.
Direct formatting from deleted result text is excluded from those structural
markers; the empty export reopens with matching native result formatting. Ordinary
history restores the original simple field, and mounted Yjs peers converge and
restore it through local undo. Six bindings cover result deletion, saved code,
undo/redo and subsequent typing. Cut, structural paste, imported explicit empty
simple results and selection spanning multiple simple fields still need coverage.

With Track Changes enabled, deleted result text stays inside one complex field,
instead of retaining a simple field beside an empty duplicate. The shared revision
recorder recognizes this structural replacement, and native deleted result text
retains its field tag on import. Accept/reject preserve the instruction; deleting
the author's own pending result insertion removes it outright. Core and six-binding
browser cases cover resolution and undo. The tracked export reopens in Word with
matching body text, field codes/results, bold and revision count. Multi-author
result replacements, partial pending insertions and cut remain separate work.

### Shared field-aware cut

Keyboard cut and the asynchronous ribbon/context-menu clipboard fallback now
reuse `deleteFieldSelection` from the core field guard. Whole simple-result cuts
retain an empty instruction, and result clipboard text remains literal. Successful
clipboard writes precede deletion; denied fallback writes and read-only cut
events preserve the document. Plain selections and complete complex-field
selections retain their ordinary host behavior.

Six bindings cover keyboard and forced fallback cut, clipboard text, saved field
codes and undo for simple/complex fields. Mounted Yjs cases cover clipboard
projection, convergence, saved empty codes and local restoration with adjacent
simple fields. Two cut/paste hook exports match native formatted-range move
references on reopen: field counts, complete text and copied bold. The native
range operation does not certify Windows clipboard formats. Legacy cut events
without clipboard data, multi-author tracked cut/paste and structural selections
still require separate evidence.

### Core field integrity follow-up

Clipboard projection removes unmatched field instructions and structural markers,
including invalid marker order, while preserving independently complete nested
fields and the formatting of literal result text. Mounted editor and six-binding
browser cases cover both boundary slices, paste into a document with an existing
field, saved marker counts and undo. Native formatted-range transfer produced
inconsistent boundary behavior and is not evidence of interactive clipboard
parity; these checks establish document integrity only.

Explicitly empty simple-field cache runs now import without invented `[Field]`
text. They reuse complex markers to retain the instruction through editor typing,
save/reparse and undo. A missing cache keeps its existing display fallback;
supported nontext cached runs remain intact. Empty result formatting survives
direct model saving, but the editor does not retain zero-length run properties.

Backspace at the end or Delete at the start of a single-grapheme simple result
preserves an empty instruction, including emoji, combining accents, flags and ZWJ
sequences split across formatting runs. Interior positions and multi-grapheme
results keep the existing host deletion behavior. Core tests cover adjacency and
undo; native subscription-build keyboard equivalence remains unverified.

Nested complex-field parsing uses independent instruction/result stack frames.
Inner cached text, including deleted text, keeps its own metadata; the outer
result resumes after the inner end. Tests cover split codes, nested instructions,
PAGE display recalculation and edited save/reparse. Tracker lifetime remains
paragraph-local, so fields spanning paragraphs require separate implementation.

### Core style resolution follow-up

Logical paragraph `start` and `end` alignment now follows the final resolved
direction when direct or derived styles switch between RTL and LTR. Explicit
physical alignment edits remain authoritative. XML import, text edit/save/reload
and unchanged style-part checks cover both logical values and both directions.

Packed transitional `tblLook` values now select the existing conditional table
styles when no named look attribute is present. Named attributes suppress the
entire packed value, including explicit off, following Microsoft's documented
precedence. Tests cover each flag, first-row/banding selection, malformed input
and cell text edit/save/reload without rewriting the source packed flags. These
are core semantic fixes, not native layout or current M365 certification.
