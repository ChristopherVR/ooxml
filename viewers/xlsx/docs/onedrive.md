# OneDrive integration

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
