# OneDrive integration

For a step-by-step setup (Microsoft Entra registration, sign-in and a complete example), start with
the [OneDrive quick start](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

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

## Open from OneDrive in File > Open

Set `fileProviders` (a property of the element and a prop of every binding) to a list of
`CloudFileProvider`s and File > Open shows places: "This device" (the Browse button) and one place
per provider, each rendering the shared `<office-ui-cloud-files>` browser (Recent, My files, search,
open from a sharing link) filtered to `.docx`. Its labels come from the Word dictionaries in every
display language.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDocxSnapshots } from 'ooxml-core/docx';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your Microsoft authentication
		signIn, // opens your sign-in UI and resolves with the account
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			docx: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

A file chosen in a place loads like any other document, takes its file name, and is read-only when
the provider reports `canEdit: false` (the stricter of the host's `readOnly` and the file wins).
File > Save and Save As go to that file's own storage instead of the host `storage` until another
document loads, and the title bar reads saved only after the upload resolved; a failed upload
stays unsaved and raises `document-error`. Opening from a place does not fire `file-command`
(the host chose the providers); Save and Save As still do, and a host that cancels them handles
them itself. A host that removes the `open` capability hides the page and the places with it.

The Word demo passes `createDemoOneDriveProvider` from `demos/onedrive` by default. Sign in from
File > Open > OneDrive: set `VITE_ONEDRIVE_CLIENT_ID` (and optionally `VITE_ONEDRIVE_AUTHORITY`) in
`demos/docx/demo-vanilla/.env.local` to use your own Entra app registration through MSAL, enter a client ID in the
demo dialog, or paste a Graph Explorer token into its "Access token" field. Register the demo's origin
(with a trailing slash) as a single-page application redirect URI. The demo edits only files on the
signed-in user's own drive; other files open read-only.

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

## Account > Connected services

The same `fileProviders` also fill the **Account** page of the File view with a "Connected services"
list (the shared `<office-ui-connected-services>`): one row per provider with the signed-in account
and Connect or Disconnect, which call the provider's own `signIn` and `signOut`. A provider without
`signOut` shows that the host manages its session. The list follows the providers' `subscribe`, so
signing in from File > Open updates it, and its words come from the Word dictionaries in every
language. Without `fileProviders` there is no Account page.

## Open a cloud file from your own UI

File > Open is only reachable once a document is open. To open a file from a start page or your own
picker, get a `CloudOpenedFile` from any `CloudFileProvider` (the shared `<office-ui-cloud-files>`
browser emits it in `office-cloud-file-open`) and pass it to `editor.openCloudFile(file)` on `<docx-editor>` (every binding's handle exposes it too). It runs the
same path as File > Open > OneDrive: the file's name, read-only when `canEdit` is false, and Save routed
to the file's own storage until another document loads.
