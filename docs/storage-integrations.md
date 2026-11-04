# Custom storage integrations

Every Word, Excel, PowerPoint and Visio editor accepts the same host-owned
`DocumentStorage` contract from `ooxml-core/embedding`. It works with all framework
bindings and plain elements. Authentication, file picking, permission checks, recovery
storage, sharing UI and provider calls stay under the application's control.

Use the smallest layer that fits:

| Layer                                            | Use it for                                             | Host supplies                                                            |
| ------------------------------------------------ | ------------------------------------------------------ | ------------------------------------------------------------------------ |
| `DocumentStorage`                                | Existing save/open flow, local disk, REST backend      | `save`, optionally `open` and recovery methods                           |
| `createStorageSession`                           | Revision-aware storage with optional three-way merging | Provider `read`, conditional `save`, `isConflict`; optionally `saveCopy` |
| `createStorageMerger`                            | Per-format merge policy                                | Format handlers and optional fallback                                    |
| `createOneDriveClient` + `createOneDriveSession` | Built-in Microsoft Graph transport                     | Token acquisition, verified write policy and optional transport settings |
| Suite bootstrap                                  | Replace the integrated app's connection and dialogs    | `ooxmlCloudOptions` or `ooxmlOneDriveOptions`                            |

There are no built-in Google Drive, Dropbox, S3 or WebDAV clients. An application
can implement these against its provider SDK or backend using the same contracts.
Credentials and provider authorization are never inferred from a file link.

## A provider with its own identity and revision

The following adapter targets an application-owned backend. These example routes
are an API the host must implement; they are not endpoints shipped by this repository.
The backend authenticates the caller, enforces permissions, and atomically checks
`If-Match` against its document revision before accepting a write.

```ts
import {
	createStorageSession,
	createStorageMerger,
	type VersionedStorageProvider,
	type StoredDocument,
} from 'ooxml-core/embedding';
import { mergeDocxSnapshots } from 'ooxml-core/docx';

interface File {
	key: string;
	revision: string;
	name: string;
	kind: 'docx';
	access: 'read' | 'write';
}
class StaleRevision extends Error {}

const provider: VersionedStorageProvider<File> = {
	async read(file, options) {
		const response = await fetch(`/api/files/${encodeURIComponent(file.key)}`, {
			...(options?.signal ? { signal: options.signal } : {}),
		});
		if (!response.ok) throw new Error('Read failed');
		const revision = response.headers.get('ETag');
		if (!revision) throw new Error('Missing revision');
		return { file: { ...file, revision }, bytes: new Uint8Array(await response.arrayBuffer()) };
	},
	async save(file, bytes, options) {
		if (file.access !== 'write') throw new Error('Read-only');
		const response = await fetch(`/api/files/${encodeURIComponent(file.key)}`, {
			method: 'PUT',
			headers: { 'If-Match': file.revision },
			body: new Blob([bytes.slice().buffer]),
			...(options?.signal ? { signal: options.signal } : {}),
		});
		if (response.status === 412) throw new StaleRevision();
		if (!response.ok) throw new Error('Save failed');
		const revision = response.headers.get('ETag');
		if (!revision) throw new Error('Missing acknowledged revision');
		return { ...file, revision };
	},
	isConflict: (error) => error instanceof StaleRevision,
};

const merge = createStorageMerger<File>({
	docx: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
});
// Metadata comes from the host's authenticated picker or API, including current permissions.
const opened: StoredDocument<File> = await provider.read(selectedFile);
const session = createStorageSession(provider, opened, {
	merge,
	maxMergeAttempts: 3,
	validateSave(meta, file) {
		if (!meta.fileName.toLowerCase().endsWith(`.${file.kind}`))
			throw new Error('Choose a separate export destination for this format');
	},
});
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
await editor.load(opened.bytes);
```

`selectedFile` and `editor` are supplied by the application. For other products, use
the load/export methods in the product references below. Request limits, timeouts,
authentication refresh and transport retries belong in the provider. Pass cancellation
through its requests. Classify only stale-revision failures as conflicts: an authentication
failure, lock, network failure or permission denial must not trigger a merge.

`saveCopy` is optional and creates a new identity. Without it, `session.saveCopy`
and native Save As reject; they never overwrite the original as a substitute.
Open and share methods can live beside the provider. A minimal session does not require
either, and never grants access automatically.

## Customize merge policy

Each handler receives `{ kind, base, local, remote, signal }`. Each version includes
its host-owned metadata; `local` is the exported byte array. Return
`{ status: 'merged', bytes }` or `{ status: 'unresolved', issues }`.
Use the existing semantic mergers or replace individual handlers:

| Kind   | Export                   | Import                    |
| ------ | ------------------------ | ------------------------- |
| `docx` | `mergeDocxSnapshots`     | `ooxml-core/docx`         |
| `xlsx` | `mergeWorkbookSnapshots` | `ooxml-core/xlsx/collab`  |
| `pptx` | `mergePptxSnapshots`     | `ooxml-core/pptx`         |
| `vsdx` | `mergeDrawingSnapshots`  | `ooxml-core/visio/collab` |

Kinds are exact strings, so a host may also register its own document types.
Missing handlers remain unresolved. An explicit `undefined` handler disables that
kind even when a fallback is present. No Office engines are imported by the router
or session; handlers can dynamically import them or call an authorized backend merger.
The registry is copied on creation; create another router to change policy.

Omit `merge` to disable automatic content merging. `maxMergeAttempts` is an integer
from 1 to 10, default 3 with a merger and 1 without. It bounds fresh remote rebases,
not network retry counts. Same-content revision refresh still works without a merger.
`conflictError` customizes the reported error; `getSnapshot().conflict` retains the
three versions regardless of the error type. `cloneFile` supports nonstandard metadata;
its default is `structuredClone`, and a custom implementation must isolate mutable fields.

The session serializes writes and retains acknowledged bytes. It also rebases queued
local snapshots created before a preceding remote merge. A failed merge never retries
the unmerged bytes against a fresh revision. `StorageConflictError` exposes the three
versions and diagnostics. A host may offer a comparison, copy, download or manual resolution.

After every save, adopt `session.getSnapshot().baseBytes`: they may include remote
changes absent from the submitted snapshot. Lock editing through export/save/adoption,
or track the editor revision and rebase newer local edits before replacing its content.
A bare `DocumentStorage.save` result cannot reload an editor. Wrap native Save or handle
the editor's cancelable file command so this adoption happens before marking it clean.
The suite already locks and adopts the merged result.

This is snapshot merging during Save. Live collaboration can use the separate
`ooxml-core/collab` transport and product mappings; neither a storage provider nor a
sharing URL supplies Microsoft Office's live coauthoring protocol.

## Integrated suite configuration

Set bootstrap configuration before loading the suite entry. Existing
`ooxmlOneDriveOptions` remains compatible and accepts the existing token, permission,
fetch, timeout and size options plus these host extensions:

| Option             | Behavior                                                                                           |
| ------------------ | -------------------------------------------------------------------------------------------------- |
| `client`           | Supply a compatible OneDrive client instead of creating the built-in client                        |
| `merge`            | Supply a complete merger, or `false` to disable merging                                            |
| `mergers`          | Override individual default format handlers; `undefined` disables a kind                           |
| `maxMergeAttempts` | Bound remote rebases                                                                               |
| `createSession`    | Supply a compatible custom session implementation                                                  |
| `sessionOptions`   | Additional session options; top-level merge policy takes precedence                                |
| `resolveConflict`  | Async host dialog returning `copy`, `latest` or `cancel`                                           |
| `shareChoices`     | Explicit `{ label, options }` entries passed to `client.share`; an empty array disables submission |

The conflict callback receives `{ doc, conflict }`; display it in the host's UI.
Returning `latest` preserves a local copy before reloading; `copy` requires provider
copy support. Returning `cancel` retains pending edits. Sharing choices are explicit
permission grants, so use understandable labels and enforce them on the backend.

For another service, set `ooxmlCloudOptions` instead:

```ts
globalThis.ooxmlCloudOptions = {
	id: 'company-documents',
	label: 'Company Documents',
	client: companyProvider,
	keyOf: (file) => file.key,
	serializeFile: ({ key, revision, name, kind, access }) => ({ key, revision, name, kind, access }),
	mergers: { xlsx: undefined }, // Disable only Excel auto-merge.
	maxMergeAttempts: 2,
	shareChoices: [
		{ label: 'Existing link, keep permissions', options: { type: 'existing' } },
		{ label: 'Project members can view', options: { audience: 'project', role: 'reader' } },
	],
};
```

`companyProvider` implements `VersionedStorageProvider` plus `open(reference, options?)`
returning `{ file, bytes }`. Its files use native suite `kind` values (`docx`, `xlsx`,
`pptx`, `vsdx`) and `access: 'read' | 'write' | 'unknown'`. Only `write` enables editing.
To support Share, provide `read` with a current `file.webUrl`; optionally add
`share(file, options)` to create a link. The suite gives custom providers only an
existing-link choice by default, unless the host explicitly supplies other choices.
Copy recovery is disabled when `client.saveCopy` is absent.

`id`, `label`, `client`, `serializeFile` and `keyOf` are required for custom providers.
Use a stable, unique provider/account namespace for `id`; `keyOf` must uniquely identify
a file within that namespace and return a string. `serializeFile` must retain identity,
revision, name, kind and access while excluding credentials, temporary transfer URLs
and capability links. Its result is persisted in local document storage.
The ID `onedrive` is reserved for the built-in/legacy schema; use `ooxmlOneDriveOptions`
to replace that connection's client without changing its metadata schema.

The suite persists the provider identity and label alongside the baseline. It refuses
to save or share a document through another configured provider. Legacy cloud records
without a provider ID belong to OneDrive. Changing a connection does not migrate files.
One provider is mounted per suite instance; a host needing a multi-account/provider
picker can supply a routing client/session and its own conflict and sharing UI.

## Product and package references

Each reference includes the package/framework table and the native editor methods:

- [Word](../viewers/docx/docs/onedrive.md): `saveBytes()` on the element; `save()` returns a Blob on framework handles.
- [Excel](../viewers/xlsx/docs/onedrive.md): flush the active cell and use `saveBytes('xlsx')`.
- [PowerPoint](../viewers/pptx/docs/guide/onedrive.md): set the shared `storage` prop and observe its save error callback.
- [Visio](../viewers/visio/docs/onedrive.md): `load` and `exportVsdx()` returning bytes with writer diagnostics.
- [OneDrive authentication, permissions and testing](onedrive-integration.md).

## Test a host integration

Run `bun run --cwd src/core test -- embedding` for provider/session tests and
`bun run --cwd apps/office-suite test:onedrive` for suite orchestration and configuration.
Build the core first so the suite's public package imports resolve.
The tests exercise a custom non-OneDrive identity/revision, concurrent saves, cancellation,
unknown merge kinds, format overrides, copy capability and cross-provider isolation.
The built-suite OneDrive browser check is `bun run test:suite:onedrive` against a
preview configured by `SUITE_URL`. It uses test-only transports, not live accounts.

For your own backend, test two clients from the same baseline: edit different content,
save one, then the other, reopen from storage and verify both changes. Also test overlap,
permission revocation, a remote writer racing a merge, canceled requests and a failed
response after the backend accepted a write. A successful save must mean the backend
acknowledged both the bytes and their revision. Keep unsupported structural changes
unresolved rather than reporting a successful merge that drops content.
