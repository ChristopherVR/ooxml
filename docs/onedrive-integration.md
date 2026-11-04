# OneDrive storage and sharing integration

Word, Excel, PowerPoint and Visio use one optional OneDrive connection from
`ooxml-core/embedding/onedrive`. It resolves an authorized sharing link, downloads the file,
and saves edited native Office bytes to the same drive item. OpenTeams can share a link to that
item; it is not another Office document editor.

This is explicit file storage and sharing. It does not connect Microsoft Office coauthoring to
the viewers' collaboration rooms, merge independently edited Office packages, or continuously
refresh an open document when another person saves it.

New to this? The [OneDrive quick start](onedrive-quickstart.md) walks through app registration,
sign-in, choosing a file and a complete React PowerPoint example, with screenshots.

## Package reference

Install `ooxml-core` as a direct dependency of a host that imports the adapter. Do not rely on
an editor's transitive dependencies being accessible to your application.

| Package type                 | Package or entry                                                                                                                           | Integration point                                                                                                                                        |
| ---------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Shared headless connection   | `ooxml-core/embedding/onedrive`                                                                                                            | `createOneDriveClient`, `createOneDriveSession`, typed file identities and conflict errors.                                                              |
| Shared storage contract      | `ooxml-core/embedding`                                                                                                                     | `DocumentStorage`; the same storage object is accepted by the product editors.                                                                           |
| Headless format engines      | `ooxml-core/docx`, `ooxml-core/xlsx`, `ooxml-core/pptx`, `ooxml-core/visio`                                                                | Read/write bytes with the format's existing model API, then call the shared session.                                                                     |
| Compatibility model packages | `docx-core`, `@christophervr/xlsx-core`, `pptx-viewer-core`, `visio-core`                                                                  | Model facades; import the OneDrive connection from `ooxml-core`, not these facades.                                                                      |
| Product elements             | `ooxml-ui/docx`, `ooxml-ui/xlsx`, `ooxml-ui/pptx`, `ooxml-ui/visio`                                                                        | Existing load/export methods and `storage` property/option.                                                                                              |
| Word bindings                | `docx-{react,vue,angular,svelte,solid,vanilla}-viewer`                                                                                     | One `<docx-editor>` storage contract, forwarded by every binding.                                                                                        |
| Excel bindings               | `@christophervr/xlsx-react-viewer`; `xlsx-{vue,angular,svelte,solid,vanilla}-viewer`                                                       | One `<xlsx-editor>` storage contract, forwarded by every binding.                                                                                        |
| PowerPoint bindings          | `pptx-{react,vue,angular,svelte,vanilla}-viewer`                                                                                           | Shared `storage` prop/input/option; each binding has its existing byte API.                                                                              |
| Visio bindings               | `visio-{react,vue,angular,svelte,solid,vanilla}-viewer`                                                                                    | One `<visio-viewer>` storage contract, forwarded by every binding.                                                                                       |
| OpenTeams bindings           | `openteams-{react,vue,angular,svelte,solid,vanilla}-viewer`                                                                                | Share the authorized OneDrive URL through the host's existing chat/file actions.                                                                         |
| CLI/MCP/server packages      | `@christophervr/pptx-viewer`, `ooxml-mcp`, `pptx-viewer-mcp`, `docx-viewer-mcp`, `xlsx-viewer-mcp`, `visio-viewer-mcp`, `openteams-server` | Do not acquire Microsoft credentials or enable OneDrive automatically. An authorized host may use the headless adapter and pass bytes to existing tools. |
| Integrated Office app        | Private `ooxml-office` (`apps/office-suite`)                                                                                               | Configure the host token provider before bootstrap; all four document editors use one connection flow.                                                   |

The earlier `ooxml-ui/pptx/onedrive` path is a compatibility surface. New hosts should use the
format-neutral core subpath. There is no shipped simulated service or simulated OneDrive demo.

## Authentication and item permissions

The application embedding the editors owns Microsoft sign-in. Register an application in
Microsoft Entra, configure its supported account types and redirect URIs, and use an appropriate
Microsoft authentication library or an existing authorized host session. A browser application
uses the authorization code flow with PKCE and must not embed a client secret.

```ts
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';

const drive = createOneDriveClient({
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
});

const opened = await drive.open(oneDriveSharingLink);
const session = createOneDriveSession(drive, opened);
const canEdit = opened.file.access === 'write';
// Give opened.bytes and canEdit to the editor, and session.storage to native Save.
```

`microsoftSession`, `permissionService` and `oneDriveSharingLink` are host-owned integration
points in this example, not bundled services. The adapter requests a fresh token when a Graph
operation needs it; the host handles expiry and interactive sign-in. No Microsoft application ID,
sign-in popup, account cache or credential persistence is supplied by the adapter.

Microsoft currently documents delegated `Files.ReadWrite` as the least privileged scope for
sharing-link resolution, upload sessions and link creation. Request only the scopes your actual
endpoints require, and respect tenant consent and Conditional Access. See
[shared-item access](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0),
[upload sessions](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0)
and [link creation](https://learn.microsoft.com/en-us/graph/api/driveitem-createlink?view=graph-rest-1.0).

An access token's scopes do not prove that the user can edit a particular file. A trusted host
permission resolver supplies `canWrite`; unknown access remains read-only. Set the editor's
host read-only restriction before loading or exposing the file. Microsoft still authorizes
every service request, including permissions that change during an editing session.

Local Office profiles and OpenTeams display identities do not constitute Microsoft authentication.
Tokens, input sharing links, temporary download URLs and upload-session URLs must not be placed
in document metadata, analytics, logs, query strings or durable application preferences.

### Personal Microsoft accounts

Personal OneDrive is an intended account type. Configure the host registration to allow personal
Microsoft accounts, or both organizational and personal accounts. Use the Microsoft identity
platform's `consumers` authority for personal-only sign-in, or `common` when the registration
accepts both. The host requests a Microsoft Graph delegated token with `Files.ReadWrite` for
this storage flow. Microsoft's [delegated authorization guide](https://learn.microsoft.com/en-us/graph/auth-v2-user)
describes these authority, client and redirect parameters; the
[permission reference](https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadwrite)
lists personal-account availability.

A login URL from a Microsoft account settings page belongs to that Microsoft application. Its
`client_id`, redirect target and requested resource scope do not authorize this application to
read or write Graph files. Do not reuse another application's client ID or redirect URI. Have
the host authentication library generate a request using its own registered values and PKCE:

```text
https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize
  ?client_id=YOUR_REGISTERED_APPLICATION_ID
  &response_type=code
  &redirect_uri=YOUR_REGISTERED_AND_URL_ENCODED_REDIRECT_URI
  &scope=Files.ReadWrite%20openid%20profile
  &code_challenge=YOUR_GENERATED_PKCE_CHALLENGE
  &code_challenge_method=S256
  &state=YOUR_GENERATED_CSRF_STATE
```

Line breaks are for explanation; this is an anatomy reference, not a URL to paste and run.
The library also retains and checks the generated state and PKCE verifier. An existing personal
Microsoft account does not eliminate the need for an authorized application. For a first API
check without creating a registration, follow the [Graph Explorer test path](onedrive-testing.md#what-can-be-tested-without-our-own-app-registration).

For a personal-account test restricted to the user's own drive, the host can deliberately allow
edits only when the returned item belongs to that verified owned drive:

```ts
const getAccessToken = () => microsoftSession.getGraphAccessToken();
const response = await fetch('https://graph.microsoft.com/v1.0/me/drive?$select=id', {
	headers: { Authorization: `Bearer ${await getAccessToken()}` },
	credentials: 'omit',
	redirect: 'error',
});
if (!response.ok) throw new Error('Could not identify the signed-in user drive.');
const ownedDrive = await response.json();
if (typeof ownedDrive.id !== 'string') throw new Error('Missing owned drive identity.');

const drive = createOneDriveClient({
	getAccessToken,
	canWrite: (file) => file.driveId === ownedDrive.id,
});
```

This is an explicit host policy for testing files on the verified personal drive, not discovery
of effective access for arbitrary shared files. External/shared drive items remain read-only
under this policy. Microsoft still enforces permission changes, locks and restrictions on every
write. A broader host needs an appropriate item-specific permission policy; do not infer one
from an "edit" link or the presence of a write scope alone.

## Open, save and close

The supported native types are `.docx`, `.xlsx`, `.pptx` and `.vsdx`. The client selects the
type from the returned file metadata; it does not parse or convert Office content. Route the
returned bytes to the matching editor and reject a mismatched type before replacing an open file.

| Method                                           | Purpose                                                                                  |
| ------------------------------------------------ | ---------------------------------------------------------------------------------------- |
| `drive.open(sharingLink, { signal })`            | Resolve an HTTPS OneDrive/SharePoint sharing link and download its file.                 |
| `drive.read(file, { signal })`                   | Read the latest file by drive/item identity without retaining its original sharing link. |
| `drive.save(file, bytes, { signal })`            | Conditional update of an existing item; callers retain the returned revision.            |
| `drive.saveCopy(file, bytes, name, { signal })`  | Create a separate sibling file rather than replacing the original.                       |
| `drive.share(file, { type, scope }, { signal })` | Create or reuse an explicitly selected view/edit sharing link.                           |
| `session.save(bytes)`                            | Serialize saves and retain the loaded baseline for meaningful conflict handling.         |
| `session.reload()`                               | Explicitly adopt the latest cloud revision; reload the returned bytes into the editor.   |
| `session.saveCopy(bytes, name)`                  | Save a separate cloud copy and continue the session against that copy.                   |
| `session.getSnapshot()`                          | Read the current file identity, baseline and any retained conflict.                      |
| `session.storage`                                | `DocumentStorage` adapter for native Save and supported native Save As.                  |

Defaults are a 32 MiB byte bound and a 60 second operation timeout. `maxBytes`, `timeoutMs`
and `fetch` can be supplied to the client. An `AbortSignal` cancels client operations. Cancel
or disregard obsolete results when switching documents, signing out or disposing a host.

Download uses Graph's preauthenticated download URL without forwarding the Graph bearer token.
See [Microsoft's browser download guidance](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0).
Persist safe drive/item/revision metadata and any required local baseline bytes, rather than a
temporary transfer URL. Baselines are document content and must follow the host's storage policy.

Native Save only succeeds after the storage promise resolves. Mark the editor clean and show
"Saved to OneDrive" only after that success. Serializing bytes, saving a local checkpoint or
starting an upload is not evidence of a successful cloud save. Keep edits made during an upload
dirty until those newer edits are saved too. Serialize cloud operations per document.

PowerPoint's native save methods report failures through the binding's error callback and notice
stack; some methods resolve rather than reject. Observe that callback as well as storage failures.
Word, Excel and Visio raise a cancelable `file-command` before their normal storage behavior. A
host that prevents it must perform the save itself; it must not also expect `storage.save` to run.

## Word (DOCX)

The plain element accepts `storage`, `readOnly`, `fileName` and `load(bytes)`. Its `saveBytes()`
exports DOCX without invoking storage. Every framework handle has `load(bytes)`, `save()` returning
a `Blob`, `markClean()` and `dirty`; use `new Uint8Array(await blob.arrayBuffer())` when exporting
through a handle. `save()` does not mark the document clean.

```ts
editor.storage = session.storage;
editor.readOnly = !canEdit;
editor.fileName = opened.file.name;
await editor.load(opened.bytes);
// A custom host Save button:
await session.save(await editor.saveBytes());
// Call markClean() only if no edit occurred after the exported snapshot.
```

Do not feed a OneDrive sharing-page URL to a generic `src` loader. Resolve the link first.
The Word connection is documented beside the [Word viewer](../viewers/docx/docs/onedrive.md).

## Excel (XLSX)

The element and framework handles expose `load(bytes, fileName?)`, `saveBytes('xlsx')` and
`markClean()`. Pass the same `storage` object and a host `readOnly` restriction in every binding.
Flush an active cell edit through the editor's normal save path before exporting. CSV is a
separate export, not bytes that may replace an XLSX item.

```ts
editor.storage = session.storage;
editor.readOnly = !canEdit;
await editor.load(opened.bytes, opened.file.name);
await session.save(await editor.saveBytes('xlsx'));
// Call markClean() only if no edit occurred after the exported snapshot.
```

See the [Excel integration reference](../viewers/xlsx/docs/onedrive.md). External workbook links,
Power Query, macro behavior and formula/rendering limitations are unchanged by cloud storage.

## PowerPoint (PPTX)

| Binding | Load                                  | Restrict editing              | Export without downloading        |
| ------- | ------------------------------------- | ----------------------------- | --------------------------------- |
| React   | `content={opened.bytes}`              | `canEdit={canEdit}`           | `await ref.current.getContent()`  |
| Vue     | `:content="opened.bytes"`             | `:can-edit="canEdit"`         | `await viewer.value.getContent()` |
| Angular | `[content]="opened.bytes"`            | `[canEdit]="canEdit"`         | `await viewer.getContent()`       |
| Svelte  | `source={opened.bytes}`               | `canEdit={canEdit}`           | `await viewer.getContent()`       |
| Vanilla | `await viewer.loadFile(opened.bytes)` | `viewer.setEditable(canEdit)` | `await viewer.save()`             |

Pass `storage: session.storage` through the binding's existing prop/input/option. `onOpenFile`
can open the host's OneDrive link chooser. Native Save now goes through the same storage object
as a custom host Save button. Avoid reassigning the source after each save, which would reset
the editing session. PPT/PPTM/PPSX exports must remain separate files and must not replace the
connected PPTX. See [PowerPoint integration](../viewers/pptx/docs/guide/onedrive.md).

## Visio (VSDX)

The element and framework handles expose `load(...)` and `exportVsdx()`, whose result includes
`bytes` and the writer's diagnostics. Native File Save delegates to `storage`. Forward `canEdit`,
`readOnly` and `storage` through the binding's options/props, then load the returned bytes using the existing
Visio load contract. For a custom Save button, call `await session.save((await viewer.exportVsdx()).bytes)`.
Review export diagnostics before claiming that unsupported diagram content was preserved.
Legacy `.vsd` previews are not supported native cloud editing targets.

See [Visio integration](../viewers/visio/docs/onedrive.md).

## Automatically combine concurrent edits during Save

Prefer `createOneDriveSession` over a bare `drive.save` callback. The suite registers a
format-aware merger for every Office product. When a revision conflicts, it reads the latest
OneDrive file and combines changes from the opened baseline, local edits and remote version.
Independent changes are saved together against the latest revision, and the combined document
is loaded back into the editor. Copy/version recovery is used only when changes cannot be combined.

Embedded package consumers opt into the same behavior through the session's `merge` callback:

```ts
import { mergeDocxSnapshots } from 'ooxml-core/docx';

const session = createOneDriveSession(drive, opened, {
	merge: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
	maxMergeAttempts: 3,
});
```

| Product    | Merger                   | Import                    |
| ---------- | ------------------------ | ------------------------- |
| Word       | `mergeDocxSnapshots`     | `ooxml-core/docx`         |
| Excel      | `mergeWorkbookSnapshots` | `ooxml-core/xlsx/collab`  |
| PowerPoint | `mergePptxSnapshots`     | `ooxml-core/pptx`         |
| Visio      | `mergeDrawingSnapshots`  | `ooxml-core/visio/collab` |

Each merger takes `(baseBytes, localBytes, remoteBytes)` and returns either
`{ status: 'merged', bytes }` or `{ status: 'unresolved', issues }`. The storage client keeps
Office engines out of its own bundle; hosts load the current format's merger as needed.
See [the suite's format router](../apps/office-suite/src/onedrive-merge.js).

After success, `session.getSnapshot().baseBytes` contains the actual saved document, including
remote edits. An embedding host must adopt those bytes, rather than leave its older local
snapshot visible. Reload only if the editor has no edits newer than the exported snapshot;
otherwise rebase those newer edits too. The suite locks the editor during Save and reloads the
combined result before enabling edits. A bare `session.storage` has no access to the editor:
hosts using native Save must wrap its save callback to adopt the acknowledged bytes safely.

Word and PowerPoint combine independent source text/property edits and stable identified nodes;
same-position text insertions retain both strings in a deterministic order. Ambiguous run/list
restructuring, incompatible replacements, colliding new relationship IDs and conflicting opaque
parts stay unresolved. Excel and Visio reuse their collaboration mappings, with additional
three-way checks for overlapping fields and unsupported content. These files contain snapshots,
so this process does not reconstruct a missing real-time CRDT operation history.

Remote changes during a merge trigger a fresh rebase, bounded by `maxMergeAttempts` (default 3).
If the content cannot be combined, the session retains the baseline, local bytes and latest
remote bytes in a structured `OneDriveConflictError`, including optional diagnostic `issues`.
If no merger is supplied, only unchanged-content revision refresh is automatic.
If the remote bytes already equal the local save snapshot, the session acknowledges the existing
cloud result instead of creating a needless conflict, for example after a lost save response.
An HTTP 423 lock is reported as `OneDriveError` with code `locked`; it is not retried as a content conflict.

```ts
import { OneDriveConflictError } from 'ooxml-core/embedding/onedrive';

try {
	await session.save(await exportCurrentBytes());
} catch (error) {
	if (!(error instanceof OneDriveConflictError)) throw error;
	// All three are retained for a host preview or comparison:
	const { base, local, remote } = error.conflict;
	showConflictChoices({ base, local, remote });
}
```

`exportCurrentBytes` and `showConflictChoices` are host UI hooks. A useful recovery dialog offers:

1. **Save my edits as a OneDrive copy:** call `session.saveCopy(local, copyName)`, update the
   editor's file name and continue on the separate file. Do not overwrite the original.
2. **Keep local copy and open latest:** persist the local edited bytes as a recoverable file first,
   explicitly confirm replacing the open editor, then call `session.reload()` and load its result.
3. **Download my edits:** export the retained local bytes without changing the cloud item.
4. **Keep editing:** leave the original local editor open and retain the conflict.

The suite also offers **Download latest OneDrive version**, with a distinct filename, so the
user can inspect both files before choosing a recovery action. Downloading either version keeps
the recovery dialog open. Read-only cloud files show that state and disable the suite Save button.

Do not clear dirty state on a failed save, remove `If-Match`, or retry an unmerged local snapshot
against a newer revision. Unsupported or overlapping edits require explicit recovery; the format
mergers do not claim that arbitrary independently changed Office files always merge safely.

The revision condition is checked when Graph creates the upload session. It is not an exclusive
lock through final commit, and a lost response can leave the outcome uncertain. Check the latest
file before deliberately retrying an interrupted upload. The current adapter does not resume
abandoned upload sessions. These limits follow the implemented request flow, not a guarantee of
Microsoft Office coauthoring. See the [upload-session reference](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0).

## Browse files

`createOneDriveFileProvider` (`ooxml-core/embedding/onedrive`) implements `CloudFileProvider`, so File > Open can browse the signed-in user's files instead of asking for a sharing link. The host still owns sign-in and tokens:

```ts
const provider = createOneDriveFileProvider({
	getAccessToken, // as for the client; called once per operation
	canWrite, // the host's write policy, unchanged
	signIn: async () => ({ name, email }), // show your sign-in UI; resolve when getAccessToken works
	signOut, // optional
	account, // optional: already signed in
	mergers: { docx: mergeDocxSnapshots, xlsx: mergeXlsxSnapshots }, // optional, by kind
});
```

- `list(folderId?)`, `recent()` and `search(query)` return `CloudListing` pages; `next` is a Graph `@odata.nextLink` that the client accepts only if it is an `https://graph.microsoft.com/` URL (the bearer token goes nowhere else). Requests use `$select` and `$top` (default 100, maximum 200).
- Entry ids are `${driveId}!${itemId}` with any `!` inside a part percent-encoded, so SharePoint drive ids and personal item ids round-trip. Recent items that arrive as `remoteItem` resolve to the target's drive and item id. OneNote packages and non-Office files are listed without a `kind`.
- `open(entry)` makes one metadata request and one download (no sharing link); `openLink(link)` keeps the sharing-link flow. Both return a `storage` backed by a `createOneDriveSession`, so Save, Save As and conflict merging behave as before. `canEdit` is true only when `canWrite` approved the file. A kind without a merger reports conflicts instead of merging; core never imports a format to supply one.
- While signed out, every call rejects with `OneDriveError('authentication')`. `subscribe` notifies on sign-in and sign-out.
- Tokens, download URLs and sharing links are never logged or stored by the provider.

### Performance notes

- The write policy (`canWrite`) runs while the bytes download, not before. Metadata requests ask only for the fields needed (`$select`), including `@microsoft.graph.downloadUrl`.
- The sharing-link request sends `Prefer: redeemSharingLink`, so a plain `Files.Read` token can open a link without a separate redemption call. Redemption can add the item to the user's "Shared with me", which is why it is limited to links the user opens explicitly.
- `getAccessToken` is called once per client operation; the pending promise is shared by that operation's requests and is never kept afterwards.
- The body is read into one buffer, preallocated from `Content-Length` when present (grown geometrically otherwise), with `onProgress(loaded, total)` per chunk.
- `client.metadata(file)` returns the current eTag and identity without downloading the content. `share()` never downloaded the file; only `read()` does, because a conflict needs the remote bytes.
- Not measured against live Graph; the tests assert request counts, ordering and concurrency, not wall-clock time.

## Sharing and OpenTeams

Opening a file does not grant additional people access. To create a link, make the scope and role
an explicit user choice, then call, for example:

```ts
const link = await drive.share(session.getSnapshot().file, {
	type: 'view',
	scope: 'organization',
});
```

Organization links require an account and tenant that support them. Anonymous links grant access
to anyone possessing the URL and may be disabled by policy. Never silently fall back to anonymous
sharing when another scope is rejected. A file's `webUrl` opens the existing item for people who
already have access; copying it is different from creating a new permission-bearing sharing link.
The direct client recognizes Graph's `users` scope, but it does not implement recipient invitation
or a recipient picker. Do not present that scope alone as a complete named-person sharing flow.

OpenTeams local workspace links, OpenTeams server attachments and OneDrive links have different
storage identities. Posting a OneDrive link leaves the original on OneDrive; uploading bytes to
an OpenTeams server creates an independent attachment. The suite's connected-file Share flow uses
the cloud item's link and does not imply synchronization of an uploaded Teams copy.
See [OpenTeams integration](../viewers/teams/docs/onedrive.md).

## Integrated Office Suite

For custom providers, authentication/transport overrides, per-format merge policies,
sharing choices and host conflict dialogs, see [Custom storage integrations](storage-integrations.md).

Before the suite entry runs, supply the same client options through the host bootstrap:

```ts
globalThis.ooxmlOneDriveOptions = {
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
};
// Load the Office Suite entry after this configuration is installed.
```

The suite shows **Open from OneDrive** and explains when sign-in has not been configured.
It routes DOCX, XLSX, PPTX and VSDX through the matching real editor. Save first checkpoints
the current bytes locally, then uploads the connected item. A cloud failure retains the checkpoint
and a pending-cloud state, including after a reload. The pending label is **Saved on this device;
OneDrive save pending**. Local document commits derive pending state from content versus the
cloud baseline, including edits from assistant operations and version restoration. Changed cloud
content cannot be committed as editable content while its host permission is read-only.

The suite's native Save As downloads a separate copy; explicit cloud-copy recovery is available
from the conflict dialog. That dialog preserves both versions and offers the recovery choices
described above. Connected-file Share can copy an existing item URL without granting access or
request an explicit organization/anonymous view/edit link. A named-recipient (`users`) sharing
flow is not implemented in the suite. The suite's local profile does not replace Microsoft sign-in.
Native Open offers the host's device/OneDrive choice; native Word Share enters the suite's
sharing flow. The suite hides PowerPoint's separate collaboration Share entry so it cannot be
mistaken for OneDrive sharing. Standalone viewer packages retain their own collaboration APIs.

## Evidence and acceptance limits

Use [OneDrive integration testing](onedrive-testing.md) for automated tests and real-account
acceptance. Test transports belong only in test fixtures; there is no checked-in simulation UI
or public mock-service export. Live account acceptance requires authorized Microsoft access.
Cloud integration does not remove any product's format fidelity limitations or prove lossless export.
