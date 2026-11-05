# OneDrive integration

## File > Open > OneDrive

Pass a provider as `fileProviders` and File > Open shows a OneDrive place next to "This device":

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your sign-in owns the token
		signIn,
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

The place lists recent files and My files, searches, and opens a sharing link. A picked workbook
loads like any other, saves back to the same item with Save (a copy with Save As), and is marked
clean only after the upload resolved; a failed upload keeps it unsaved. A file `canWrite` rejects
(another drive, a shared file) opens read-only. `fileProviders` is a prop of every framework
binding (`fileProviders` in React, Vue, Solid, Svelte, Angular and Vanilla). The demo app
(`demos/xlsx`) passes the demo OneDrive provider by default; configure it with
`VITE_ONEDRIVE_CLIENT_ID` or paste a Graph Explorer token in its sign-in dialog. The lower-level
`storage` wiring below is for hosts that build their own Open UI.

The same storage contract also supports custom providers. See
[provider-neutral storage and merge configuration](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)
for `createStorageSession`, per-format merge overrides and suite bootstrap settings.
Use these extension points with every Excel framework package.

Excel uses the shared `ooxml-core/embedding/onedrive` connection. Resolve the sharing link in the
host, verify item-specific permissions and give the existing editor the downloaded XLSX bytes:

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
await editor.load(opened.bytes, opened.file.name);
```

The element and framework handles expose `saveBytes('xlsx')`, `load(bytes, fileName?)` and
`markClean()`. A custom host save calls `session.save(await editor.saveBytes('xlsx'))`; mark clean
only after persistence succeeds with no newer edits. Use the normal editor save path to commit
an active cell edit. CSV is a separate export and must never replace the connected XLSX item.
Native Save delegates to `storage` unless the host prevents the cancelable `file-command` event.

The same `storage` and `readOnly` contract reaches `@christophervr/xlsx-react-viewer`,
`xlsx-vue-viewer`, `xlsx-angular-viewer`, `xlsx-svelte-viewer`, `xlsx-solid-viewer` and
`xlsx-vanilla-viewer`, as well as the plain `ooxml-ui/xlsx` element. `@christophervr/xlsx-core`
is the model facade, not an authentication service.

See the [Office OneDrive package reference](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#excel-xlsx)
for authentication, sharing and recovery that retains both local and remote versions. The
[testing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) explains
credential-free test fixtures and live account acceptance; no simulation UI or service is shipped.

Register `mergeWorkbookSnapshots` from `ooxml-core/xlsx/collab` as the session's `merge` callback
to combine supported independent workbook edits during Save. The suite enables it automatically;
overlapping cell edits and unsupported changes still require recovery. This does not connect
workbook collaboration sessions to Microsoft Excel coauthoring, refresh external data or change
the [feature limitations](/features).

## Open a cloud file from your own UI

File > Open is only reachable once a document is open. To open a file from a start page or your own
picker, get a `CloudOpenedFile` from any `CloudFileProvider` (the shared `<office-ui-cloud-files>`
browser emits it in `office-cloud-file-open`) and pass it to `editor.openCloudFile(file)` on `<xlsx-editor>` (every binding's handle exposes it too). It runs the
same path as File > Open > OneDrive: the file's name, read-only when `canEdit` is false, and Save routed
to the file's own storage until another document loads.
