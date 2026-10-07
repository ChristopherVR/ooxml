# OpenTeams parity review and implementation sequence

Reviewed 7 October 2026 against `src/core/teams`, `src/ui/src/teams`,
`viewers/teams`, the reference server, and the Teams browser tests.

The target is Microsoft Teams user-workflow parity. This is an early workspace,
not an implementation of the Microsoft Teams service or its app platform.
No feature is considered equivalent solely because a control is present.

## Current coverage

| Area                            | Evidence in the implementation                                                                                 | Remaining work                                                                                                           |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| Channels and posts              | Yjs posts and live thread panes, personal followed threads, replies, reactions, author-side edit/delete checks | Teams hierarchy, private/shared channels, membership enforcement, cross-device follows, thread notifications, moderation |
| Direct and group chats          | The channel model has a `direct` kind                                                                          | No participant-scoped chat workflow or server-enforced privacy                                                           |
| Search and unread               | Message search, attachment metadata search, per-channel local read markers                                     | Indexed file contents, filters, mentions, activity feed, notifications, shared read receipts                             |
| Presence                        | Awareness, availability and typing                                                                             | Authenticated identity, idle state, richer status and privacy controls                                                   |
| Meetings                        | Prejoin, microphone, camera, screen share, raised hand, mesh WebRTC                                            | Scheduling, invitations, SFU, lobby, host roles, recording, captions, backgrounds and large calls                        |
| File sharing                    | Uploads with cancel/retry and file progress, workbook creation, unique storage names, signed download links    | Permissions, versions, folders, durable local-mode sharing, byte progress and storage cleanup                            |
| Office content                  | Native Word, Excel and Visio previews; static PowerPoint preview; XLSX local editing and channel save copies   | Full PowerPoint rendering/playback, coediting, write-back/version conflict handling and fidelity acceptance corpus       |
| Markdown                        | Safe blocks and flat inline formatting, task lists, pipe tables and relative web links                         | Full CommonMark/GFM, nested structure, reference links and inline precedence                                             |
| Sites and web apps              | Sandboxed HTML/site previews and shared file/website channel tabs                                              | App permissions, approved origins, app messaging and authentication                                                      |
| Accounts and administration     | Reference server has optional shared token and origin allowlist                                                | User accounts, SSO, tenant/team/channel ACLs, guests, audit, retention and policy enforcement                            |
| Bindings                        | Six lifecycle bindings share `TeamsProps` and the same app                                                     | Framework-by-framework browser acceptance for the new embedding prop                                                     |
| Accessibility and visual parity | Existing Lit controls and token styles                                                                         | Keyboard/focus review, screen-reader acceptance, responsive/mobile workflow coverage, reference screenshots              |

## First implemented slice: content previews

Attachment Open actions now retain the workspace and display a preview pane.
The cancelable `teams-open-file` event and per-kind `openers` still take precedence.
Closing returns to the current workspace view. Switching channel or workspace
clears the preview; superseded fetches are aborted and late responses ignored.

- Word and Excel use the existing native editor elements in read-only mode.
- Visio uses the existing viewer. Its local viewer interactions are not shared
  edits and are not written back to file storage.
- Markdown supports ATX headings, bullets, quotes, fenced code, bold, italic,
  inline code, flat task lists, top-level pipe tables and file-relative web links. Raw HTML remains text.
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
completed-file progress and cancellation, not byte-level progress. File search covers loaded
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

## Sixth implemented slice: file transfer control

The Files controls now cancel uploads and workbook creation, retain their captured
destination for retry, and show completed-file progress. Core copy-save actions
also accept cancellation and progress callbacks. The workbook save-copy toolbar
offers cancellation during serialization or upload, retains local edits, and allows
retry. It does not yet display byte or file-count progress. Closing the Files panel or replacing its client
aborts its active operation. Cancellation stops waiting even when a host storage
adapter ignores its optional abort signal, and guards prevent late attachment
publication. Already written objects can remain in storage because there is no
cleanup contract. Browser tests exercise pending cancellation, late completion,
retry, and progress between files across all six bindings. Workbook browser acceptance
also checks canceled save copies, ignored late storage writes, retained dirty edits,
and a successful retry without publishing the canceled copy.

## Seventh implemented slice: structured Markdown previews

Markdown previews now render top-level pipe tables with column headers, alignment,
escaped pipes, missing-cell padding and excess-cell truncation. Flat unordered task
items show disabled checked/unchecked controls: a preview does not mutate the source.
Relative links resolve against the attachment URL, while unsafe schemes, credentials
and protocol-relative links remain literal text. Signed source queries are not copied
to sibling resources. Tables have keyboard-focusable horizontal scrolling and use
the same safe inline text bindings as the rest of the preview. No HTML or automatic
image loading is introduced. Nested containers, full inline precedence, reference
links and CommonMark/GFM conformance remain unsupported. The syntax reference is
the [GFM specification](https://github.github.com/gfm/#tables-extension-);
this subset does not establish Microsoft Teams rendering equivalence.
Tables are limited to 128 columns and 16,384 generated cells (including headers).
Wider tables and rows beyond that limit remain readable as ordinary source text.

## Eighth implemented slice: channel thread panes

Channel conversations now show root posts with live reply counts and open a focused
thread pane. Nested reply chains resolve to their original post; deleted parents
remain visible, and missing-parent/cyclic input retains a deterministic visible root.
Only non-deleted replies contribute to counts. Thread selection stays local to each
client and resets on a channel change. Replies, reactions, edits, deletion and file
cards reuse existing core actions. Search results open the corresponding thread.
The pane replaces the main conversation at narrow widths. Browser acceptance uses
two real local clients and checks live replies, nested replies, deleted parents,
search navigation and responsive layout. Thread notifications,
cross-pane draft retention, authenticated membership and server enforcement remain
unsupported. The reference workflow is Microsoft's
[thread pane and followed threads](https://support.microsoft.com/en-us/teams/teams-channels/follow-threads-in-microsoft-teams).

## Ninth implemented slice: personal followed threads

The thread pane supports follow/unfollow. A personal Followed threads view lists
channel roots, live non-deleted reply counts and latest activity, and opens the
correct channel and thread. Preferences persist locally per user and workspace;
they do not enter the shared channel document. Restored preferences tolerate
missing messages until hydration, late parents, duplicate aliases and unavailable
storage. Archived channels are excluded. Unit coverage checks restoration and user
isolation; browser acceptance checks follow, live counts, reload, navigation and
unfollow across six bindings. This is a device-local list: cross-device preference
sync, automatic following and thread notifications remain unsupported.

## Tenth implemented slice: unread followed threads

Followed threads carry separate personal read markers. Opening a thread marks
its current messages read; rendering its pane in a visible browser document also
marks newly arrived replies read. Viewing channel roots alone does not mark the
collapsed threads read. The list exposes unread counts, mark-read/unread actions
and an unread-only filter. New non-deleted messages from other authors contribute
to counts; edits and reactions do not. Explicit unread marks persist locally.
Tests cover foreign/self/deleted messages, no-op read updates, restoration, filtering
and reopening. This is not a shared read receipt or notification delivery system;
cross-device state, viewport-level reading detection, activity notifications and
automatic following on mentions remain outstanding. These controls follow Microsoft's
[followed-thread inbox workflow](https://support.microsoft.com/en-us/teams/teams-channels/follow-threads-in-microsoft-teams).

## Eleventh implemented slice: automatic following preferences

Successful new posts (including channel file posts) and replies through the client
actions automatically follow their thread root. Personal settings independently
disable following started threads or replied-to threads and persist per user and
workspace. Defaults enable both. Edits and rejected empty sends do not create
follows, and existing follow choices are not changed retrospectively when a setting
changes. Low-level writes through `workspace.chat` remain host-controlled and do
not infer personal follow intent. Unit tests check both settings, restoration,
author action results and defensive preference copies; browser acceptance checks
post following, opting out, reload and independently enabling reply following.
Automatic following on mentions and following all content in a channel remain
unsupported, alongside notifications and cross-device preference sync.

## Twelfth implemented slice: personal message drafts

Posts, thread replies and edits retain separate drafts through channel, thread and
view navigation. A Drafts view lists unsent messages and resumes their original
compose context. Drafts stay outside Yjs and persist per user and workspace in
local storage. Text is restored after reload; file bytes stay in memory during a
session. After reload, restored attachment names require reattaching or explicit
discard before sending. The UI does not silently send a draft without those files.
Successful sends clear their captured draft without clearing newer text typed
during an upload; rejected edits retain the draft. Local storage is bounded to
512 KiB of serialized text/metadata, 500 restored contexts and twenty attachments per
draft; blocked storage or larger snapshots preserve only session state. Cross-device
draft sync, durable attachment bytes, unavailable-message recovery and shared read
receipts remain outstanding. Unit tests cover context isolation, defensive copies,
restoration, blocked storage and send races. Browser acceptance covers post/thread
navigation, retained attachment bytes, central resume, reload, missing-attachment
guard and sending. Microsoft's current
[Drafts quick view](https://support.microsoft.com/en-us/teams/platform/what-s-new-in-microsoft-teams)
is the workflow reference; this local implementation does not establish full parity.

## Thirteenth implemented slice: reliable chat attachment transfers

Message attachments use configured storage, retain their display names and receive
fresh storage names on every attempt. A post or thread reply is published only
after every file has a valid URL. Missing storage, invalid files and partial upload
failures retain the complete draft instead of publishing filename-only cards.
Channel transfer status shows completed file counts and supports cancellation,
including adapters that ignore abort signals. Failures and cancellations recover
the original text and attachment bytes in Drafts; newer text typed during the send
is preserved as a separate draft. Resuming and sending uses the captured channel
and thread context. Unit and browser coverage exercises partial failure, retry,
cancellation, late adapter completion, preserved bytes and concurrent new text.
Progress counts completed files rather than transferred bytes. Uploaded storage
objects are not automatically removed after failure or cancellation because the
storage contract has no deletion operation. Durable attachment recovery, shared
storage permissions, folders and version/save-back contracts remain outstanding.

## Fourteenth implemented slice: Shared and Files presentation

Channel files use the current Shared tab label. The Files view presents a New menu,
Upload command and search field above a semantic table of filenames, sharing
dates, authors and channel locations. Clicking a filename opens the existing
preview; pinning retains the existing channel-tab workflow. The native file picker
is hidden behind Upload, and the prominent website URL form has been removed from
Files. Websites remain available through Add tab and host content previews.
Unavailable file bytes disable opening instead of pretending a filename is a file.
Desktop and narrow-screen browser acceptance covers upload, retry, workbook
creation, search, filename opening and table overflow. This is an incremental UI
alignment with Microsoft's [Shared file workflow](https://support.microsoft.com/en-us/teams/files/collaborate-on-files-in-microsoft-teams),
not pixel-perfect parity. Folder navigation, file selection/bulk actions, views,
sorting and permission-aware sharing menus still need implementation.

## Fifteenth implemented slice: Settings navigation and personal appearance

The Settings shell now separates General, Appearance and accessibility,
Notifications and activity, Files and links, and Connection. Light/dark/system
choices apply within OpenTeams and persist per user/workspace on this device;
they do not change the host document theme. Followed-thread preferences share the
existing core personal settings. Keyboard category navigation and mobile layouts
are covered in browser acceptance. Connection settings preserve validation and
explicit Apply/Cancel behavior; cancel discards unapplied edits. Replacing the
client flushes pending local document snapshots so reconnect does not discard
recent posts or tabs. Files and links initially explains the built-in preview;
the next slice adds the browser default. Desktop defaults, density, notification delivery, account/privacy controls
and device selection still require implementation. Microsoft's
[settings categories](https://support.microsoft.com/en-us/accessibility/teams/customize-your-teams-chat-interface-with-chat-density-settings)
and [notification preferences](https://support.microsoft.com/en-us/teams/notifications-settings/manage-notifications-in-microsoft-teams)
are references, not evidence of complete visual or feature parity.

## Sixteenth implemented slice: file actions and browser viewer preference

Shared file rows expose a More actions popover with Open in OpenTeams, Open in
browser, Download, Copy link and Pin as tab. Downloads reuse bounded content reads,
request a fresh signed URL and preserve original bytes and filenames. Clipboard
failure retains a selectable link. Menus remain open through unrelated presence
updates, and a changed file/client cancels the pending action. Settings > Files
and links saves the Office open preference per user and workspace on this device.
Browser opens the same native content renderer in a separate tab, not a raw Office
URL. The viewer carries its bounded, validated payload in the fragment, does not
join the collaboration room, and has no opener. Local XLSX editing and download
remain available there; channel save-back and host callback adapters are available
in the main workspace only. Explicit Open in OpenTeams bypasses the personal
browser default; existing host open handlers retain precedence. Browser acceptance
checks actual workbook rendering in both contexts, exact downloaded bytes, link
copying, preference restoration and pin/save-copy behavior. Signed links can expire;
copying does not grant or manage recipient permissions. Desktop defaults, folder
and bulk actions, version/save-back contracts and permission-aware sharing remain
outstanding. Microsoft's [file-open preferences](https://support.microsoft.com/en-us/office/collab-files/open-file-links-directly-in-microsoft-365-desktop-apps-from-teams-and-classic-outlook)
and [Shared file downloads](https://support.microsoft.com/en-us/teams/files/download-a-file-from-microsoft-teams)
are the workflow references.

## Next releasable slices

1. UI parity: match the current Teams shell, Shared/Files commands, Settings
   navigation, typography, spacing and responsive layouts against reference
   screenshots. Keep implemented preferences functional and expose unsupported
   settings honestly. Visual similarity alone does not establish workflow parity.
2. Shared tab permissions: authenticated membership and server enforcement,
   approved web-app origins and a supported app messaging contract.
3. Full PowerPoint adapter: expose the existing framework-neutral renderer
   through the UI package, then consume it here without copying format or render
   logic. Test actual slide content, navigation, media and teardown.
4. Files workflow: permission-aware storage IDs, progress/cancel/retry, folders,
   version metadata, and save-back contracts. Add coediting only after identity,
   access control and revision handling are enforceable.
5. Authenticated direct/group chats and activity: participant-scoped rooms,
   mentions, notification preferences, read receipts and message-level search.
6. Meetings and enterprise integration: scheduling and invitations, SFU/lobby
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
