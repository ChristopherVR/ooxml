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

| Area                  | Evidence and remaining work                                                                                                                                                                                                    | Reuse boundary                                                                                                        |
| --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------- |
| Editable pages        | Editing remains continuous; Print Layout is read-only. Caret, selection, IME and keyboard mappings need contracts before editable page surfaces.                                                                               | Shared docx layout; one Word editor                                                                                   |
| Layout                | Continuous section breaks are explicitly approximated as page breaks in `docx/layout/page-flow.ts`. Complex typography and Word font metrics need reference documents.                                                         | Injectable measurer and shared layout engine                                                                          |
| Tables                | Complex/merged structural editing and collaborative table changes remain limited. Header pagination was incorrectly repeating noncontiguous marked rows.                                                                       | Word model, editing and layout in core                                                                                |
| Drawing objects       | The viewer roadmap identifies missing general shape, chart and SmartArt editing. Equations are display-only in the newer editor.                                                                                               | Geometry, diagram and future shared DrawingML/chart/math areas, rather than imports of PowerPoint UI internals        |
| Proofing and review   | Browser spelling is not a Word grammar engine. Compare, protection, richer references and automatic field calculation need explicit implementations.                                                                           | Product logic in docx; host service contracts where appropriate                                                       |
| Collaboration         | Word uses authority-ordered ProseMirror steps and separate transient presence. Shared collab already has Yjs, awareness, WebSocket protocol and external WebRTC/WebSocket adapters. No Word CRDT editor binding is wired here. | Format-neutral lifecycle/providers in collab; Word schema mapping in docx; cursor DOM and Share/status controls in UI |
| Save and preservation | The model alone does not carry the loaded package's opaque parts or media. CRDT synchronization of model JSON cannot establish preservation or shared export reliability.                                                      | LoadedDocument/package preservation and assets in core                                                                |

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

## Next implementation sequence

1. Use the consolidated editor in `src/ui/src/docx` and headless Word helpers in
   `src/core/docx/ui`, preserving existing APIs and the six thin bindings.
2. Add opt-in Yjs collaboration to that single editor using
   [y-prosemirror](https://github.com/yjs/y-prosemirror), the existing Word schema,
   and shared CollabSession/provider lifecycle. Keep the authority-step mode
   available and prevent both engines from controlling one editor state.
3. Use Yjs relative positions for selections and collaborative undo restricted to
   local changes. Seed only after initial sync, define late-join and file-load
   policy, and test concurrent text/formatting edits and offline reconnects before
   exposing richer table commands. Include headers, footnotes and comments in the
   schema/ownership design rather than synchronizing only the body silently.
4. Define package/media bootstrap, authenticated provider permissions,
   persistence and canonical export contracts. The current client role is
   advisory; neither provider sync nor awareness proves server authorization or
   durable persistence.
5. Establish a Word-authored corpus with Word-rendered pages and semantic
   expectations. Gate each feature on import diagnostics, edit/undo, export,
   save/reopen preservation, measured layout tolerances and all six bindings.

End-to-end Word Yjs editing, editable pages and 1:1 Word parity remain unfinished.
The first slice improves the shared foundation without advertising them as done.
