# OneDrive integration

For a step-by-step setup (Microsoft Entra registration, sign-in and a complete example), start with
the [OneDrive quick start](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

For another service or a custom OneDrive implementation, see
[provider-neutral storage integrations](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md).
`createStorageSession` connects to the same `storage` property across all Visio packages,
with host-owned revision checks, merge handlers, conflict UI and sharing policies.

## File > Open with cloud places

Pass `fileProviders` (a property of the element and a prop of every binding) to add cloud places
to File > Open. When the list is not empty, Open shows **This device** (the usual Browse tile) and
one place per provider, which renders the shared `<office-ui-cloud-files>` browser for `.vsdx`
drawings: sign in, Recent, My files, search and open from a sharing link. The words come from
Visio's display language in every locale.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDrawingSnapshots } from 'ooxml-core/visio/collab';

const oneDrive = createOneDriveFileProvider({
	getAccessToken, // the host's own sign-in
	signIn, // opens the host's sign-in UI and resolves with the account
	canWrite: async (file) => file.driveId === (await myDriveId()),
	mergers: {
		vsdx: ({ base, local, remote }) => mergeDrawingSnapshots(base.bytes, local, remote.bytes),
	},
});
// <VisioViewer fileProviders={[oneDrive]} /> (React); [fileProviders] in Angular; :file-providers in Vue.
```

A drawing chosen there loads through the normal `.vsdx` path and takes the file's name, then:

- It is read-only when the file is (`canEdit` false). A host that already made the viewer read-only
  (`canEdit`, `readOnly`, a collaboration `viewer` role or `capabilities.edit`) stays read-only
  for an editable file: the stricter of the two wins.
- File > Save and Save As use that file's own storage instead of the host `storage`, until another
  drawing loads (Open, New, `load` or `document`). Save resolves only after the provider stored the
  bytes: the drawing reads clean then, and stays unsaved with a `document-error` when the upload fails.
- The cancelable `file-command` event still fires first (`open`, `save`); cancelling `open`
  leaves the current drawing in place. The backstage closes once the drawing is open.
- Export diagnostics are reported as for any other save. Legacy `.vsd` is not a cloud editing target.

The demos pass a OneDrive place by default, signed in through the demo-only dialog (a client ID, or
a Graph Explorer token pasted into **Access token**). That dialog is demo configuration, not
product UI; see `demos/onedrive`. To skip the dialog, set `VITE_ONEDRIVE_CLIENT_ID` (and optionally
`VITE_ONEDRIVE_AUTHORITY`) in `demos/visio/.env.local`, run `bun run dev` in `viewers/visio`, and
register the address it prints (with a trailing slash) as a single-page application redirect URI.

## Account > Connected services

The same `fileProviders` also fill the **Account** page of the File view with a "Connected services"
list (the shared `<office-ui-connected-services>`): one row per provider with the signed-in account
and Connect or Disconnect, which call the provider's own `signIn` and `signOut`. A provider without
`signOut` shows that the host manages its session. The list follows the providers' `subscribe`, so
signing in from File > Open updates it, and its words come from the Visio dictionaries in every
language. Without `fileProviders` the Account page shows only the profile and product information.

## Using a session directly

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

## Open a cloud file from your own UI

File > Open is only reachable once a document is open. To open a file from a start page or your own
picker, get a `CloudOpenedFile` from any `CloudFileProvider` (the shared `<office-ui-cloud-files>`
browser emits it in `office-cloud-file-open`) and pass it to `viewer.openCloudFile(file)` on `<visio-viewer>` (every binding's handle exposes it too). It runs the
same path as File > Open > OneDrive: the file's name, read-only when `canEdit` is false, and Save routed
to the file's own storage until another document loads.
