# Word parity review and first implementation slice

Reviewed 2026-10-07. This is a source review and a bounded first implementation,
not a claim of Microsoft Word equivalence or a completed browser assessment.

## Source state and ownership

The consolidated workspace uses `src/core` (`ooxml-core`) and `src/ui`
(`ooxml-ui`). Keep format logic in `src/core/docx`, format-neutral collaboration
in `src/core/collab`, reusable controls in `src/ui`, and framework lifecycle/event
adapters in the viewer. Core must never import the UI package.

The reviewed viewer entry re-exports `ooxml-ui/docx`. The initial shared checkout
was detached at an older commit and lacked the Word editor and its exports.
On preparing this slice for publication, current `origin/main` was found to
already contain the consolidated sources in `src/ui/src/docx` and
`src/core/docx/ui`, with the viewer under `viewers/docx`. These existing sources
are the integration target. The slice is landed from an isolated checkout of
current main, without publishing the older workspace migration. No source has
been extracted in this slice.

The viewer's parity roadmap and browser contracts describe six bindings over one
editor, continuous editing, read-only paginated Print Layout, formatting, lists,
tables, sections/stories, fields, comments and tracked changes. Those contracts
are useful feature evidence, but were inspected rather than executed in this
review. They do not establish pixel-equivalent Word output.

## Principal gaps

| Area                  | Evidence and remaining work                                                                                                                                                                                                                                                                        | Reuse boundary                                                                                                        |
| --------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Editable pages        | Editing remains continuous; Print Layout is read-only. Caret, selection, IME and keyboard mappings need contracts before editable page surfaces.                                                                                                                                                   | Shared docx layout; one Word editor                                                                                   |
| Layout                | Continuous section breaks are explicitly approximated as page breaks in `docx/layout/page-flow.ts`. Complex typography and Word font metrics need reference documents.                                                                                                                             | Injectable measurer and shared layout engine                                                                          |
| Tables                | Complex/merged structural editing and collaborative table changes remain limited. Header pagination was incorrectly repeating noncontiguous marked rows.                                                                                                                                           | Word model, editing and layout in core                                                                                |
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

## Next implementation sequence

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

Full M365 Word parity remains unfinished and must not be claimed without this
reference evidence.
