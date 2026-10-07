# OpenTeams parity review and implementation sequence

Reviewed 7 October 2026 against `src/core/teams`, `src/ui/src/teams`,
`viewers/teams`, the reference server, and the Teams browser tests.

The target is Microsoft Teams user-workflow parity. This is an early workspace,
not an implementation of the Microsoft Teams service or its app platform.
No feature is considered equivalent solely because a control is present.

## Current coverage

| Area                            | Evidence in the implementation                                                                               | Remaining work                                                                                                     |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| Channels and posts              | Yjs channel/message model, replies, reactions, author-side edit/delete checks                                | Teams hierarchy, private/shared channels, membership enforcement, threaded side pane, moderation                   |
| Direct and group chats          | The channel model has a `direct` kind                                                                        | No participant-scoped chat workflow or server-enforced privacy                                                     |
| Search and unread               | Message search, attachment metadata search, per-channel local read markers                                   | Indexed file contents, filters, mentions, activity feed, notifications, shared read receipts                       |
| Presence                        | Awareness, availability and typing                                                                           | Authenticated identity, idle state, richer status and privacy controls                                             |
| Meetings                        | Prejoin, microphone, camera, screen share, raised hand, mesh WebRTC                                          | Scheduling, invitations, SFU, lobby, host roles, recording, captions, backgrounds and large calls                  |
| File sharing                    | Direct uploads with retry, workbook creation, Files views, unique storage names, signed download links       | Permissions, versions, folders, durable local-mode sharing, byte progress and cancellation                         |
| Office content                  | Native Word, Excel and Visio previews; static PowerPoint preview; XLSX local editing and channel save copies | Full PowerPoint rendering/playback, coediting, write-back/version conflict handling and fidelity acceptance corpus |
| Markdown                        | Safe block and flat inline subset added in this change                                                       | Full CommonMark/GFM, tables, task lists, nested structure and relative links                                       |
| Sites and web apps              | Sandboxed HTML/site previews and shared file/website channel tabs                                            | App permissions, approved origins, app messaging and authentication                                                |
| Accounts and administration     | Reference server has optional shared token and origin allowlist                                              | User accounts, SSO, tenant/team/channel ACLs, guests, audit, retention and policy enforcement                      |
| Bindings                        | Six lifecycle bindings share `TeamsProps` and the same app                                                   | Framework-by-framework browser acceptance for the new embedding prop                                               |
| Accessibility and visual parity | Existing Lit controls and token styles                                                                       | Keyboard/focus review, screen-reader acceptance, responsive/mobile workflow coverage, reference screenshots        |

## First implemented slice: content previews

Attachment Open actions now retain the workspace and display a preview pane.
The cancelable `teams-open-file` event and per-kind `openers` still take precedence.
Closing returns to the current workspace view. Switching channel or workspace
clears the preview; superseded fetches are aborted and late responses ignored.

- Word and Excel use the existing native editor elements in read-only mode.
- Visio uses the existing viewer. Its local viewer interactions are not shared
  edits and are not written back to file storage.
- Markdown supports ATX headings, bullets, quotes, fenced code, bold, italic,
  inline code and absolute/origin-relative web links. Raw HTML remains text.
  This is a subset, not CommonMark or GFM parity.
- `.txt`, `.csv` and `.json` have literal text previews.
- HTML attachments and the Files view's **Preview website** action use sandboxed
  frames. Scripts and forms are allowed; same-origin, popups and top navigation
  are not. Sites may refuse framing or require functionality the sandbox blocks.
  An external-open link remains available; iframe load is not proof of successful
  site rendering. There is no TeamsJS or Microsoft app manifest compatibility.
- PowerPoint has a default static preview with navigation, slide text and notes.
  A host embedding page via `embeds.pptx` can override it.

All native byte reads omit ambient credentials and referrers. Text is capped at
2 MiB and Office bytes at 32 MiB, including streamed responses without a size
header. Cross-origin files need appropriate CORS. Signed URLs are preserved;
an expired link must be reopened from the attachment to request a fresh link.
Local-mode name-only attachments still have no bytes to preview.

### Integration contract

All six component bindings accept `embeds` through `TeamsProps`. Angular exposes
an input; Vue declares and watches the prop; Solid snapshots it; React, Svelte
and vanilla forward the common props contract.

```ts
const embeds = {
	pptx: ({ url }: { url: string | undefined }) =>
		`https://viewer.example.com/embed?source=${encodeURIComponent(url ?? '')}`,
};
```

The host must provide a real viewer page that understands `source` and supports
the sandbox and CORS. A normal demo URL does not automatically load an attachment.
Do not configure the raw `.pptx` URL as an embed page. The callback is trusted
host configuration, never peer-provided code. Only web URLs are accepted.

The custom element also exposes `previewContent({ attachment, url })` for
host-provided content. It is a local view action and does not publish a tab or
message to other users. For a website use `kind: 'other', mime: 'text/html'`.

## Second implemented slice: shared tabs and XLSX workflows

Channel members can pin an uploaded file from Files or add a named website tab.
Tab definitions use stable IDs in the workspace CRDT and persist with its snapshot;
selection remains local. Creator-side rename/remove checks are advisory, just like
message ownership, and do not provide authenticated server authorization. Peer data
is validated before rendering. Name-only attachments cannot become file tabs.

Excel opens in viewing mode and offers **Edit workbook** and **Download workbook
copy**, including without channel storage. With server storage or a host upload
adapter configured, it also offers **Save copy to channel**. Channel saving serializes
the native workbook, uploads under a unique storage name, and posts a new attachment
in the channel captured when the file opened. It does not overwrite the original,
switch the tab to the saved copy, or coedit with other users. Macro workbooks keep
their `.xlsm` extension; macro execution is not supported.

Failed uploads leave the workbook dirty and available for retry. Edits made during
upload remain dirty after the older snapshot is shared. Navigation asks before
discarding edits and is blocked while saving. If another client removes an open
tab, its local pane remains available until closed. Unrelated presence updates
do not reload open content. These guards do not protect against a host forcibly
replacing the component, identity or workspace configuration.

Evidence: core tab merge/snapshot and upload tests; UI lifecycle tests; browser
tests with two same-browser clients adding, renaming and removing a website tab;
and an actual XLSX fixture edited, saved, parsed and reopened, including upload
failure/retry and editing during upload. The eight workflow specs now run against
each of vanilla, React, Vue, Angular, Svelte and Solid with
`bun run test:browser:bindings` in `viewers/teams`. Each run forces fresh Vite
dependency optimization to cover first-open behavior, including the lazy
PowerPoint engine and Word metafile converter.
The suite uses local mode and test upload adapters; it is not evidence of
authenticated remote coediting or renderer fidelity across all files.

Microsoft Teams supports editing and coediting files from channel tabs and uses
SharePoint-backed channel folders. Save copies are an incremental OpenTeams
workflow, not evidence of matching those collaboration or storage semantics.

## Third implemented slice: creating, uploading and finding files

The channel and workspace Files views now offer **Upload files**, **New Excel
workbook** and file search by name, author, channel and type. New files target
the channel named beside the controls. Workbook creation delegates to the
SpreadsheetML engine, stores a real blank XLSX, posts it in the captured channel
and opens it locally. Existing same-name files keep their own storage URLs.

Uploads and new workbooks require configured storage. Failures retain the selected
files for retry, with the original destination named in the error. Uploaded batches
are posted only when every upload succeeds; a failed batch can leave unreferenced
storage objects and currently has no server cleanup contract. The size limit is
32 MiB per file, with at most 20 files in a batch. This implementation reports a
busy state, not byte-level progress or cancellation. File search covers loaded
attachment metadata, not indexed document contents or folders.

Message attachments now also keep their original channel and reply target while
uploads are pending, without clearing a newer reply started in another channel.
Browser evidence covers workbook creation, same-name upload preservation, upload
failure/retry, file search and opening uploaded Markdown bytes.

## Fourth implemented slice: static PowerPoint previews

PPTX attachments now open without host embed configuration. The pane delegates
parsing and SVG generation to the existing core PowerPoint engine, displays the
SVG as an isolated image, and offers previous/next navigation, speaker notes,
top-level slide text and parser compatibility warnings. Replacing or closing
a preview disposes the parser and revokes slide image URLs. External images
are disabled and expanded package bytes are limited to 128 MiB.

Visual inspection of a real ten-slide deck exposed unwrapped paragraphs in
the SVG exporter. Core export now wraps horizontal text, retains run formatting
and supports optional host font measurement. The Teams pane supplies canvas
metrics for the fonts the browser actually renders; headless callers use the
existing core font tables and fallback metrics. Explicit no-wrap text is honored.

This is a static baseline, not PowerPoint renderer parity. Font substitution,
paragraph geometry, vertical text, autofit and complex layouts can differ.
Animations, transitions, media playback and editing are unavailable. The pane
states these limits and exposes text so layout differences do not hide all
readable content. Browser acceptance checks actual deck bytes, visible slide
images, navigation, text and cleanup; it does not compare Microsoft reference
screenshots or prove layout equivalence.

## Fifth implemented slice: local workbook copies

The latest XLSX increment adds local copy downloads. Downloads serialize through
the existing workbook engine and do not upload or post a new attachment. The
explicit copy-download action reports that the download started and keeps its local dirty state;
browser delivery does not prove that the user retained the file. Invalid pending
cell edits block both downloads and channel saves instead of exporting the old
cell value. Switching viewing/editing mode also commits or rejects pending edits.
Browser acceptance covers a real stop-validation rule, downloading and parsing
edited bytes, source preservation, absence of a shared attachment, and navigation
confirmation. This remains a copy workflow, with the engine's existing format
support and preservation limitations.

## Next releasable slices

1. Shared tab permissions: authenticated membership and server enforcement,
   approved web-app origins and a supported app messaging contract.
2. Full PowerPoint adapter: expose the existing framework-neutral renderer
   through the UI package, then consume it here without copying format or render
   logic. Test actual slide content, navigation, media and teardown.
3. Files workflow: permission-aware storage IDs, progress/cancel/retry, folders,
   version metadata, and save-back contracts. Add coediting only after identity,
   access control and revision handling are enforceable.
4. Authenticated direct/group chats and activity: participant-scoped rooms,
   mentions, notification preferences, read receipts and message-level search.
5. Meetings and enterprise integration: scheduling and invitations, SFU/lobby
   and roles, captions/recording, then SSO/admin/retention/guest policy.

For each slice: record the Microsoft reference workflow, implement in the owning
layer, add protocol/unit coverage and browser acceptance across the six bindings,
and document observable differences. Measure rendering fidelity with real files
and Microsoft-generated reference output, rather than claiming lossless parity.

## Microsoft reference workflows

- [File and app tabs in channels and chats](https://support.microsoft.com/en-us/teams/teams-channels/use-a-tab-in-a-channel-or-chat-in-microsoft-teams)
- [File collaboration and channel storage](https://support.microsoft.com/en-us/teams/files/collaborate-on-files-in-microsoft-teams)
- [Web app tabs and their embedding model](https://learn.microsoft.com/microsoftteams/platform/tabs/what-are-tabs)
- [Website tab behavior differs from Teams web apps](https://devblogs.microsoft.com/microsoft365dev/upcoming-updates-to-loading-websites-in-teams-tabs/)

These sources define the first content-workspace scope. They are not evidence
that the remaining features have been matched or exhaustively reviewed.
