# OneDrive integration

Visio uses the shared `ooxml-core/embedding/onedrive` connection for native VSDX files. The host
owns Microsoft sign-in, link resolution and item-specific permission verification. Give the
downloaded bytes to the existing viewer, with a host read-only restriction when access is unknown:

```ts
// Set these through the element or binding's existing props/options.
const canEdit = opened.file.access === 'write';
const options = { storage: session.storage, canEdit, readOnly: !canEdit };
// After mounting with options:
await viewer.load(opened.bytes);
```

`exportVsdx()` returns an object containing the native `bytes` and writer diagnostics. A custom
Save button can call `await session.save((await viewer.exportVsdx()).bytes)`. Native File Save
uses the same storage contract unless the host prevents the cancelable `file-command` event.
Review diagnostics rather than promising lossless export. Binary `.vsd` previews are not native
OneDrive editing targets.

This applies to the `ooxml-ui/visio` element and `visio-react-viewer`, `visio-vue-viewer`,
`visio-angular-viewer`, `visio-svelte-viewer`, `visio-solid-viewer` and `visio-vanilla-viewer`.
`visio-core` is the model facade. The connection does not add Microsoft authentication to it.

See the [Office OneDrive reference](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#visio-vsdx)
for package wiring, auth, permission-aware sharing and conflict recovery that preserves both
versions. See [testing](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)
for test-local transports and live service acceptance. No simulation UI or service is shipped.

Register `mergeDrawingSnapshots` from `ooxml-core/visio/collab` as the session's `merge` callback
to combine supported independent shape and property edits during Save. The suite does so automatically.
Overlapping or ambiguous changes still require recovery. This is not Microsoft coauthoring or a file lock.
Existing [Visio limits](/parity) and export diagnostics remain applicable.
