---
title: OneDrive integration
description: Connect every PowerPoint viewer binding to shared Office OneDrive storage, sharing and conflict recovery.
---

# OneDrive integration

Every PowerPoint framework package also accepts a custom storage provider through its
existing `storage` prop. See [provider-neutral integration and configuration](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)
for `createStorageSession`, authentication/transport overrides, per-format merge handlers,
host conflict dialogs and explicit sharing choices.

PowerPoint uses the same optional `ooxml-core/embedding/onedrive` connection as Word, Excel and
Visio. The host supplies Microsoft authentication and verified item permissions. There is no
bundled Microsoft app registration, simulation service or simulated OneDrive demo.

For a step-by-step setup with screenshots (app registration, sign-in, a sharing-link box and a
complete React example), start with the [OneDrive quick start](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).
Every framework demo also has a OneDrive mode (`?onedrive=1`) for trying it against your own
OneDrive; the quick start's "Try it in the PowerPoint demos" section explains the setup.

The [complete Office integration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)
documents authentication, session storage, safe cloud copies, sharing roles/scopes, conflict
recovery and the reference for every package type.

```ts
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';

const drive = createOneDriveClient({
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
});
const opened = await drive.open(sharingLink);
const session = createOneDriveSession(drive, opened);
// Pass session.storage and opened.file.access === 'write' to the viewer.
```

The session and permission services in this sample are host integrations, not bundled objects.
Import the adapter as a direct `ooxml-core` dependency. `ooxml-ui/pptx/onedrive` remains only a
compatibility surface for the earlier PowerPoint-specific import.

## All five bindings

| Binding/package                | Open bytes                            | Restrict editing              | Export current PPTX               |
| ------------------------------ | ------------------------------------- | ----------------------------- | --------------------------------- |
| React, `pptx-react-viewer`     | `content={opened.bytes}`              | `canEdit={canEdit}`           | `await ref.current.getContent()`  |
| Vue, `pptx-vue-viewer`         | `:content="opened.bytes"`             | `:can-edit="canEdit"`         | `await viewer.value.getContent()` |
| Angular, `pptx-angular-viewer` | `[content]="opened.bytes"`            | `[canEdit]="canEdit"`         | `await viewer.getContent()`       |
| Svelte, `pptx-svelte-viewer`   | `source={opened.bytes}`               | `canEdit={canEdit}`           | `await viewer.getContent()`       |
| Vanilla, `pptx-vanilla-viewer` | `await viewer.loadFile(opened.bytes)` | `viewer.setEditable(canEdit)` | `await viewer.save()`             |

`canEdit` means `opened.file.access === 'write'`. Unknown access remains read-only. Every binding
accepts `storage` as its existing prop/input/option. Assign `session.storage` to connect native
File Save to OneDrive. `onOpenFile` can open the host's sharing-link chooser; do not pass the
Microsoft sharing-page URL to a generic document source loader.

The viewer marks the deck saved only after `storage.save` resolves, and only when nothing was
edited while the upload was in flight; a rejected save keeps it unsaved. For a host-owned Save
button, export the current PPTX and call `session.save(bytes)`. Keep the
editor mounted and preserve undo history. Track edits made during upload separately from the
uploaded snapshot. Do not treat a native save method resolving as proof of success: some methods
report failures through `onError` and the notice stack. See [host-owned storage](/guide/host-owned-storage).

## Conflicts and sharing

A revision-only change can be retried once when the downloaded cloud bytes still equal the
loaded baseline. A genuine content conflict retains baseline, local and remote versions in
`OneDriveConflictError`. Offer **Save my edits as a OneDrive copy**, **Keep local copy and open
latest**, **Download my edits**, and **Keep editing**. Do not silently overwrite or reload.
Register `mergePptxSnapshots` as the session's `merge` callback to combine supported independent
slide edits automatically. The suite does this by default. Unsupported overlaps still require
recovery; the session does not lock the file through the final upload commit.

Save As to a compatible native PPTX creates a separate cloud copy through session storage.
Exports to PPT/PPTM/PPSX or another format must not replace the original PPTX; keep them as an
explicit separate export workflow. A Share operation creates a link only after the user chooses
its view/edit role and scope. Sharing a OneDrive link does not connect a PowerPoint collaboration
room to Microsoft PowerPoint coauthoring.

See [integration testing](/guide/onedrive-testing) and the existing [format limitations](/guide/limitations).
