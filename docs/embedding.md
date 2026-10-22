# Embedding the editors

This guide is for hosts that put an Office editor inside their own application: what to set, what
the editor emits, and which side owns each decision. The contract is shared: the types live in
`ooxml-core/embedding` (DOM-free) and `ooxml-ui` mounts regions and menus. Each product's own
`api.md` (for PowerPoint, `viewers/pptx/docs/guide/customization.md`) has the exact field list; this
page is the map across products.

## What do you want to control?

| You want to                                            | Word                         | Excel                        | PowerPoint                                                 | Visio                            |
| ------------------------------------------------------ | ---------------------------- | ---------------------------- | ---------------------------------------------------------- | -------------------------------- |
| Take features away (print, export, copy, ...)          | `customization.capabilities` | `customization.capabilities` | `customization.capabilities`                               | `customization.capabilities`     |
| Make the document read-only                            | `readOnly`                   | `readOnly`                   | `readOnly` (plus `canEdit` or `editable` to allow editing) | `readOnly` (or `canEdit: false`) |
| Add a File tab page                                    | `backstage.extraPages`       | `backstage.extraPages`       | `backstage.extraPages`                                     | `backstage.extraPages`           |
| Add right-click commands                               | `contextMenu.extraCommands`  | `contextMenu.extraCommands`  | `contextMenu.extraElementCommands`, `extraCanvasCommands`  | `contextMenu.extraCommands`      |
| Replace a pane, dialog body or File page               | not offered                  | not offered                  | `customization.regions`                                    | not offered                      |
| Hide individual ribbon controls                        | `hiddenActions`              | `hiddenActions`              | `customization.ribbon.*`, `hiddenActions`                  | `hiddenActions`                  |
| Add your own ribbon tab                                | `ribbonAddIns`               | `ribbonAddIns`               | `ribbonAddIns`                                             | `ribbonAddIns`                   |
| Own where files are saved and opened                   | `storage`, `file-command`    | `storage`, `file-command`    | `storage`, `onOpenFile`                                    | `storage`, `file-command`        |
| Share a live document                                  | `collaboration`              | `collaboration`              | `collaboration`                                            | `collaboration`                  |
| Let your own assistant read and edit the open document | live tools                   | live tools                   | live tools                                                 | live tools                       |
| Compose only some parts of the UI                      | building blocks              | building blocks              | building blocks                                            | building blocks                  |

Teams (`<teams-app>`) is a workspace, not a document editor: it has no `customization`, capabilities
or ribbon. It is configured through its server configuration and the `createTeamsClient` store (see
`docs/teams-area.md`); only collaboration roles and theming concern it.

Every binding forwards `customization` as one prop (React, Vue, Solid, Svelte), one input (Angular)
or an option (vanilla). The object is applied when its identity changes, so keep it stable (memoise
it); `null` restores the defaults.

## Capabilities

`customization.capabilities` is one switch per capability. `false` removes every UI path to it
(ribbon controls, File tab pages, menu entries, shortcuts, Tell me); `true` or an omitted key leaves
the editor's own rules in charge. A product honours a subset of the shared ids
(`EDITOR_CAPABILITY_IDS`), and its capability type is a narrow union, so naming an id the product does
not have is a compile error.

| Id         | Word | Excel | PowerPoint | Visio | Removes                                   |
| ---------- | ---- | ----- | ---------- | ----- | ----------------------------------------- |
| `edit`     | yes  | yes   | yes        | yes   | Editing: the document opens read-only     |
| `export`   | yes  | yes   | yes        | yes   | Save As, Export, Download                 |
| `print`    | yes  | yes   | yes        | yes   | Print commands and Ctrl+P                 |
| `share`    | yes  | yes   | yes        | yes   | Share pages and buttons                   |
| `copy`     | yes  | yes   | yes        | yes   | Copy, Cut and their shortcuts             |
| `comment`  | yes  | yes   | yes        | yes   | Adding and editing comments               |
| `review`   | yes  | no    | no         | no    | Tracked changes: tracking, accept, reject |
| `find`     | yes  | yes   | no         | yes   | Find and Replace                          |
| `open`     | yes  | yes   | no         | yes   | New and Open                              |
| `save`     | yes  | yes   | no         | yes   | Save                                      |
| `settings` | yes  | yes   | no         | yes   | File > Options and Customize Ribbon       |
| `present`  | no   | no    | yes        | no    | PowerPoint only: the Slide Show tab, F5   |

`hiddenActions` keeps working beside capabilities; the two are unioned. A hidden command also stays
out of the ribbon search (Tell me) and of Customize Ribbon, and its action refuses to run if
something still calls it.

`edit: false` and `readOnly` are different. `readOnly` is a prop the host can flip at any time,
mirrored by `readonly-change` in Word and Excel. `edit: false` is a policy ceiling of its own: it
holds editing off whatever `readOnly` says, none of the locks lifts another, and it never shows up as
`hostReadOnly`, so a parent that mirrors that field never writes the capability back. See
[Read-only and roles](#read-only-and-roles).

Capabilities are UI policy, not document protection. A user with the browser's developer tools still
has the bytes you handed the editor; if a document must not be copied or exported, do not send it to
the browser.

```ts
import type { WordCustomization } from 'docx-core';

editor.customization = {
	capabilities: { print: false, export: false, copy: false },
} satisfies WordCustomization;
```

## Host File tab pages

`backstage.extraPages` adds pages to the File tab, after the built-in pages of their group. A page is
`{ id, label, group?, order?, render }`: `group: 'footer'` places it next to Account and Options,
`order` sorts host pages within a group, and an empty or repeated `id` is ignored. `label` is shown as
given, so localise it yourself. The body is a host region (`backstage.host.<id>`) and `render` takes
either form of a region renderer:

- a mount function `(container, context) => void | cleanup | { update, destroy }`, or
- a custom element, as its tag name or `{ element, properties }`. The editor creates it, assigns
  `properties` once and keeps its `context` property current.

The page mounts the first time it is shown and is torn down when the page is removed or the editor
disconnects. `context` carries `region`, `locale`, `canEdit` and `close()`, plus product fields (Word
`editor`, `fileName`; Excel `editor`, `fileName`, `activeSheet`; Visio `viewer`, `pageIndex`,
`pageCount`).

```ts
const editor = document.querySelector('xlsx-editor')!;
editor.customization = {
	backstage: {
		extraPages: [
			{
				id: 'policies',
				label: 'Policies',
				render: (container, { fileName, close }) => {
					container.innerHTML = `<h2>${fileName}</h2><button id="ok">Done</button>`;
					container.querySelector('#ok')!.addEventListener('click', close);
				},
			},
		],
	},
};
```

PowerPoint also has `addBackstagePage(page)` and `removeBackstagePage(id)` on the viewer handle.

## Host context-menu commands

`contextMenu` adds commands to the right-click menus. A command is
`{ id, label, icon?, group?, when?, disabled?, onSelect }`: `group` is `top` (default, above the
built-in entries) or `bottom`; `when` hides it, `disabled` greys it out (a boolean or a predicate),
and `onSelect` runs after the menu closes with an immutable snapshot of what the menu was opened on.
Host ids cannot collide with built-in entries, a repeated id keeps the first, and menu separators are
repaired around the commands. Labels are shown as given.

| Product    | Field                                                     | Menus                                      | `onSelect` context                                            |
| ---------- | --------------------------------------------------------- | ------------------------------------------ | ------------------------------------------------------------- |
| Word       | `contextMenu.extraCommands`                               | The document menu                          | `selectedText`, `from`, `to`, `readOnly`, `inTable`, `onLink` |
| Excel      | `contextMenu.extraCommands`                               | Cell, row header, column header, sheet tab | `menu`, `sheet`, `sheetName`, `ref`, `active`, `readOnly`     |
| PowerPoint | `contextMenu.extraElementCommands`, `extraCanvasCommands` | Element menu, empty-canvas menu            | `{ slideIndex, elementIds }` and `{ slideIndex }`             |
| Visio      | `contextMenu.extraCommands`                               | Shape and page menus                       | `target`, `pageIndex`, `pageId`, `shapes`                     |

Use `when` to offer a command in some menus only, for example in Excel:

```ts
editor.customization = {
	contextMenu: {
		extraCommands: [
			{
				id: 'send',
				label: 'Send range',
				when: ({ menu }) => menu === 'cell',
				onSelect: ({ sheetName, ref }) => send(sheetName, ref),
			},
		],
	},
};
```

## Regions

A region is a place in the layout that the host fills with its own UI: the editor keeps the frame
(size, resize handle, collapse state, dialog backdrop and close button) and hands over an empty
container. The renderer forms are the same as for File tab pages.

**Only PowerPoint offers regions today.** `customization.regions` replaces the inspector, slide rail,
notes, status bar, the bodies of the Options, Share, Broadcast and Print dialogs, and each File tab
page (`backstage.<page>`); hiding wins over replacing. Word, Excel and Visio do not support region
replacement: their panes and dialogs can be hidden through capabilities and `hiddenActions`, and you
can add File tab pages, ribbon tabs and menu commands, but you cannot swap a built-in body.

```ts
viewer.setCustomization({ regions: { 'dialog.options': 'acme-settings' } });
```

Mounting a React, Vue or Angular component into a region is covered in the PowerPoint guide
(`viewers/pptx/docs/guide/customization.md`, "Mounting a framework component").

## Add-in ribbon tabs

`ribbonAddIns` adds tabs after the product's own, as an Office add-in does. The descriptor
(`RibbonAddInTab` from `ooxml-core/ribbon`, re-exported by `ooxml-ui` and the framework packages) is
plain data, so one tab works in every product. A tab is `{ id, label, keytip?, groups }`, a group is
`{ label, commands }`, and a command is `{ id, label, icon?, size?, title?, disabled?, items?, run? }`.

A chosen command runs its `run` callback and then dispatches a bubbling `office-ribbon-add-in` event
with `detail: { tab, command }`, so one listener hears every host command. A tab that reuses a
built-in tab id, or has an empty or repeated id, is left out. Leaving the prop out keeps the current
tabs; an empty array removes them. Passing an equal descriptor again keeps the controls and only
replaces the callbacks.

```tsx
const tabs: RibbonAddInTab[] = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [{ label: 'Export', commands: [{ id: 'run', label: 'Run', run: () => runReport() }] }],
	},
];

<DocxEditor ribbonAddIns={tabs} />;
```

PowerPoint's ribbon has no drop-down for host commands (a command with `items` is drawn as its
items), and host tabs show in the desktop ribbon only. `hiddenActions` does not apply to host tabs.

## Host-owned storage

Every document editor accepts a `storage` (a property of the element, a prop or input of each
binding): the `DocumentStorage` contract from `ooxml-core/embedding`. Only `save` is required.

```ts
import type { DocumentStorage } from 'ooxml-core/embedding';

const storage: DocumentStorage = {
	async save(bytes, { fileName, mimeType, reason }) {
		await fetch(`/files/${encodeURIComponent(fileName)}`, {
			method: 'PUT',
			headers: { 'content-type': mimeType },
			body: bytes,
		});
		return { fileName }; // optional: the name the document now has
	},
	open: () => myDrive.pick(), // optional: { bytes, fileName }, or undefined when canceled
};
```

- **Save.** File > Save, Save As and Ctrl+S call `storage.save(bytes, { fileName, mimeType, reason })`
  instead of downloading, with `reason` `'save'` or `'saveAs'`. When it resolves the document is
  marked saved, and a returned `{ fileName }` renames it.
- **Open.** File > Open calls `storage.open()` when it exists; without it the browser file picker is
  used.
- **Recovery.** `readRecovery`, `writeRecovery` and `clearRecovery` replace the library's IndexedDB
  AutoRecover copies, and are used only when all three exist. `listRecovery` is optional and lets
  lists (Recent, Version History) and the retention sweep see the host's copies. Only PowerPoint and
  Visio have AutoRecover; Word and Excel ignore the recovery methods.
- **Concurrency.** A second Save while one is pending is ignored, not queued. If the user edits while
  the host save runs, the document stays dirty, since the saved bytes do not contain the edit.
- **Failure.** A rejected `save` surfaces as the product's error event (`document-error` in Word and
  Visio, `workbook-error` in Excel, `error` in PowerPoint) and the document stays dirty.
- **`file-command`.** Word, Excel and Visio raise a cancelable `file-command` event before Save, Open
  and New (Word and Excel also before the exports and Print). It fires first: `preventDefault()`
  (in Visio's `events` map, return `false`) handles the command yourself and `storage` is not called.
  PowerPoint has no `file-command`: use `storage` for saving and `onOpenFile` for File > Open.

Exports and Print still use the browser. The per-product notes are in each `api.md` and in
`viewers/pptx/docs/guide/host-owned-storage.md`.

## Read-only and roles

Two things decide whether a user can change a document, and they are deliberately separate.

- **The host's permission: `readOnly`.** Every full editor and viewer element and its framework
  bindings take `readOnly` (boolean, default `false`). `true` makes the document read-only; `false`
  is the same as not setting it. Word and Excel also emit `readonly-change` when it flips.
- **The collaboration role: `CollaborationOptions.user.role`** (`src/core/collab/identity.ts`, see
  `docs/collab-area.md`). A `viewer` role makes the editor read-only for as long as the session is
  joined, and edits from a `viewer` are never published.

There is no separate role prop. Both restrictions apply together: the editor is writable only when
the host allows it and the role allows it. In every product the role lock is tracked apart from the host
switch, and the effective value drives the UI and `readonly-change`. Turning `readOnly` off never lifts a `viewer` role, and
leaving the room restores whatever the host last set.

| Product    | Prop                                                               | Inverse alias (kept)                             |
| ---------- | ------------------------------------------------------------------ | ------------------------------------------------ |
| Word       | `readOnly`                                                         | none                                             |
| Excel      | `readOnly`                                                         | none                                             |
| Visio      | `readOnly`                                                         | `canEdit` (`canEdit: false` is `readOnly: true`) |
| PowerPoint | `readOnly` (`editable` in Svelte and vanilla, `canEdit` elsewhere) | the edit switch                                  |

If both spellings are set, **`readOnly: true` wins**: a deck or drawing is editable only when the
edit switch allows it and `readOnly` is not `true`. Setting `readOnly: false` does not grant edit
permission by itself. PowerPoint's edit switch defaults to off, so a PowerPoint host still passes
`canEdit` (or `editable`) to allow editing, and uses `readOnly` to take it away again, for example
per user.

Building blocks (the individual toolbar, canvas and pane elements) keep `canEdit`, read-only unless
it is `true`; they are not full editors and take no `readOnly`.

`customization.capabilities.edit` in PowerPoint is a third, independent restriction (a host policy
ceiling); like the role, it can only take editing away.

### Mirroring `readOnly` from the editor

Word and Excel emit `readonly-change` whenever read-only changes: the user toggled Editing or
Viewing, the host changed `readOnly`, or a collaboration role was joined or left. The detail is
`{ readOnly, locked, hostReadOnly }`:

- `readOnly` is the effective value (the host's switch or a `viewer` role).
- `locked` is true while a `viewer` role holds the editor read-only.
- `hostReadOnly` is the host's own switch.

A parent that keeps `readOnly` in its own state must store `hostReadOnly`, never `readOnly`;
otherwise a role lock would be written into the host's setting and leaving the room would not
restore it. The two-way channels do this for you: Angular's `[(readOnly)]` (`readOnlyChange`) and
Vue's `v-model:readOnly` (`update:readOnly`) carry the host switch only and stay quiet when only a
role lock changed. The callbacks (`onReadOnlyChange`, Vue's `readonly-change`) receive the
effective value first and the detail second. Word's `readonly-change` detail was a boolean before
and is now this object. Visio and PowerPoint have no read-only change event, so nothing there can
write a role lock back.

## Collaboration

Every editor takes the same `collaboration` property (`CollaborationOptions`: `roomId`, `serverUrl`,
`provider`, `webrtc`, `session`, `user` with `name`, `color`, `avatar` and `role`, and `intent`) and
reports the same `collaboration-change` state, so a host integrates it once. A `viewer` role makes the
editor read-only (see above) and is advisory on the client: a server must enforce permissions. What
merges concurrently differs per product. See `docs/collab-area.md` ("Embedding contract" and "Honest
limits") and each product's collaboration page.

```ts
editor.collaboration = {
	roomId: 'plan-2026',
	serverUrl: 'wss://sync.example.com',
	user: { name: 'Ada', color: '#2b6cb0', role: 'editor' },
};
```

## Letting your assistant drive the open editor

A host's own assistant (a chat panel, an agent behind your backend) can read and edit a document
that is open on screen with the same tools the MCP servers and the suite's assistant use. Register
each mounted editor with a `LiveToolHost` from `ooxml-core/automation`, then call tools by name:

```ts
import { automationTools, createLiveToolHost, liveEditorFromElement } from 'ooxml-core/automation';
import { pptxTools } from 'ooxml-core/pptx/automation';

const host = createLiveToolHost([...Object.values(automationTools).flat(), ...pptxTools], {
	// Optional: show a diff from `before` and `after` and return false to discard the edit.
	review: async ({ tool, before, after }) => confirmWithUser(tool, before, after),
});
const forget = host.register('budget', liveEditorFromElement('xlsx', sheetElement, 'Budget.xlsx'));

host.describe(); // function-calling specs for the open formats, each requiring a `target`
await host.call('xlsx_set_cells', {
	target: 'budget',
	params: { sheetIndex: 0, cells: [{ address: 'B2', input: '=SUM(B3:B9)' }] },
});
```

- **Every call names its document.** `target` is required and there is no "current document", so one
  assistant can work across several open editors without editing the wrong one. `host.targets()`
  lists what is registered, with each document's format and read-only state, for the model's prompt.
- **Reads see unsaved work.** A read runs on what the editor would save now, so it includes edits the
  user has not saved.
- **Edits respect the editor.** A read-only editor is never edited. An edit is refused (nothing is
  loaded) when the user changed the document while the tool or your `review` ran, and edits to one
  document run one at a time. `review` is where a host asks the user to accept a change.
- **Create tools are not offered**: they have no document to target. Build the file with the tool
  and open it as a new document yourself.
- **Not a transaction.** The tools work on bytes, so an accepted edit is loaded back as a document
  load. Expect the editor to behave as after opening the file (its selection and undo history are not
  carried over); applying tool edits in place is not implemented.
- Like capabilities, this is a client-side convenience, not access control: enforce what an agent
  may change on your server.

For a headless backend with no editor on screen, use the tools on bytes directly (`docs/mcp.md`).

## Building blocks vs the full editor

The full editors (`<docx-editor>`, `<xlsx-editor>`, `<visio-viewer>`, the PowerPoint viewer) bring the
chrome, and `customization` trims it. To compose your own layout, each product also exports its
toolbar, canvas and panes separately (React: `useViewerBuildingBlocks` plus the element components).
Building blocks are read-only until the host passes `canEdit: true`, and take no `readOnly`,
`customization` or `storage`: you own the chrome, so you own those decisions. See
`viewers/<product>/docs/building-blocks.md`.

## Known limits

- No region replacement outside PowerPoint: Word, Excel and Visio cannot swap their panes, dialogs or
  built-in File tab pages (host File tab pages, ribbon tabs and menu commands work).
- PowerPoint has no `file-command` event; use `storage` and `onOpenFile`.
- Only PowerPoint and Visio have AutoRecover, so the recovery methods of `storage` do nothing in Word
  and Excel.
- Exports (CSV, text, image) and Print still use the browser; `storage` covers Save, Save As and Open.
- React's headless `useViewerBuildingBlocks` takes no `storage` (nor `customization`); use the full
  component for those.
- Teams has no chrome or feature switches.
- Vanilla PowerPoint takes `customization` as an initial option only; change it at runtime with
  `setCustomization`.
- Capabilities, roles and `readOnly` are UI policy on the client, not protection of the bytes.
- Host labels (menu commands, File tab pages, ribbon tabs) are not translated; localise them yourself.
