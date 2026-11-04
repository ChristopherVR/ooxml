# OneDrive integration

For other services or custom OneDrive flows, use the same `storage` property with
[`createStorageSession` and the provider-neutral integration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md).
Hosts can replace authentication, transport, merge handlers, conflict UI and sharing policy
without changing any Word framework binding.

All Word packages use the shared `ooxml-core/embedding/onedrive` connection. The host resolves
the sharing link, supplies Microsoft authentication and verified write permission, and connects
the resulting session to the existing editor storage contract.

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
editor.fileName = opened.file.name;
await editor.load(opened.bytes);
```

The plain `<docx-editor>` in `ooxml-ui/docx` exposes `saveBytes()`. Every framework handle exposes
`save()` returning a `Blob`, `load(bytes)`, `markClean()` and `dirty`. A custom Save button can
convert the Blob to bytes and call `session.save(bytes)`; mark clean only when persistence
succeeds and no newer edit intervened. Native Save goes through `storage` automatically unless
the host prevents the cancelable `file-command` event.

This applies to `docx-react-viewer`, `docx-vue-viewer`, `docx-angular-viewer`, `docx-svelte-viewer`,
`docx-solid-viewer` and `docx-vanilla-viewer`. Each forwards the same `storage` and `readOnly`
props/options. `docx-core` is a model facade and does not implement Microsoft sign-in.

The [Office OneDrive reference](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#word-docx)
contains the package matrix, auth setup, sharing scopes and actionable conflict recovery. The
[testing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) separates
test-local transports from live Microsoft acceptance. No simulation UI or service is shipped.

Register `mergeDocxSnapshots` from `ooxml-core/docx` as the session's `merge` callback to combine
independent text and property edits during Save. The suite enables this automatically. Ambiguous
structure or incompatible replacements retain both versions for recovery. This does not connect
Word collaboration rooms to Microsoft Office coauthoring. Existing [preservation limits](/preservation) still apply.
