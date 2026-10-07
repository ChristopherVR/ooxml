# OpenTeams parity review and implementation sequence

Reviewed 7 October 2026 against `src/core/teams`, `src/ui/src/teams`,
`viewers/teams`, the reference server, and the Teams browser tests.

The target is Microsoft Teams user-workflow parity. This is an early workspace,
not an implementation of the Microsoft Teams service or its app platform.
No feature is considered equivalent solely because a control is present.

## Current coverage

| Area                            | Evidence in the implementation                                                    | Remaining work                                                                                              |
| ------------------------------- | --------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| Channels and posts              | Yjs channel/message model, replies, reactions, author-side edit/delete checks     | Teams hierarchy, private/shared channels, membership enforcement, threaded side pane, moderation            |
| Direct and group chats          | The channel model has a `direct` kind                                             | No participant-scoped chat workflow or server-enforced privacy                                              |
| Search and unread               | Message search, per-channel local read markers                                    | File search, filters, mentions, activity feed, notifications, shared read receipts                          |
| Presence                        | Awareness, availability and typing                                                | Authenticated identity, idle state, richer status and privacy controls                                      |
| Meetings                        | Prejoin, microphone, camera, screen share, raised hand, mesh WebRTC               | Scheduling, invitations, SFU, lobby, host roles, recording, captions, backgrounds and large calls           |
| File sharing                    | Upload adapter, file cards, Files views, signed download links                    | Permissions, versions, folders, durable local-mode sharing and upload progress                              |
| Office content                  | Native Word, Excel and Visio previews; XLSX local editing and channel save copies | PowerPoint default renderer, coediting, write-back/version conflict handling and fidelity acceptance corpus |
| Markdown                        | Safe block and flat inline subset added in this change                            | Full CommonMark/GFM, tables, task lists, nested structure and relative links                                |
| Sites and web apps              | Sandboxed HTML/site previews and shared file/website channel tabs                 | App permissions, approved origins, app messaging and authentication                                         |
| Accounts and administration     | Reference server has optional shared token and origin allowlist                   | User accounts, SSO, tenant/team/channel ACLs, guests, audit, retention and policy enforcement               |
| Bindings                        | Six lifecycle bindings share `TeamsProps` and the same app                        | Framework-by-framework browser acceptance for the new embedding prop                                        |
| Accessibility and visual parity | Existing Lit controls and token styles                                            | Keyboard/focus review, screen-reader acceptance, responsive/mobile workflow coverage, reference screenshots |

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
- PowerPoint can use a host embedding page via `embeds.pptx`. Without one, the
  pane explicitly reports that no viewer is configured and offers external open.

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

With server storage or a host upload adapter configured, Excel opens in viewing
mode and offers **Edit workbook** and **Save copy to channel**. Saving serializes
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
failure/retry and editing during upload. Browser coverage currently uses vanilla;
the six bindings share the app and are built/typechecked, but are not each tested
against this workflow in a browser.

Microsoft Teams supports editing and coediting files from channel tabs and uses
SharePoint-backed channel folders. Save copies are an incremental OpenTeams
workflow, not evidence of matching those collaboration or storage semantics.

## Next releasable slices

1. Shared tab permissions: authenticated membership and server enforcement,
   approved web-app origins and a supported app messaging contract.
2. Default PowerPoint adapter: expose the existing framework-neutral renderer
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
