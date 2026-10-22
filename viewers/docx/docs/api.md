# Element API and events

Every framework adapter wraps the same `<docx-editor>` custom element. The adapters forward these properties and re-emit these events in their framework's idiom (see [framework bindings](/bindings)).

## Attributes and properties

| Property         | Attribute         | Description                                                                                    |
| ---------------- | ----------------- | ---------------------------------------------------------------------------------------------- |
| `documentModel`  |                   | The `DocumentModel` shown. Assigning it is an external replacement, not an edit event.         |
| `readOnly`       | `read-only`       | Blocks local edits. Remote collaboration batches still apply.                                  |
| `locale`         | `locale`          | Interface language: `en`, `fr`, `de`, `es` or `zh-CN`. Document content is never translated.   |
| `fileName`       | `file-name`       | Name of the open document (default `Document1.docx`).                                          |
| `reviewAuthor`   | `review-author`   | Author recorded on tracked changes and comments (default `Author`).                            |
| `theme`          | `theme`           | `light`, `dark` or `auto`. See [theming](/theming).                                            |
| `themeColors`    |                   | Partial token overrides applied as `--dve-*` custom properties.                                |
| `showThumbnails` | `show-thumbnails` | Left rail of page thumbnails (Print Layout only).                                              |
| `showToolbar`    | `show-toolbar`    | `false` hides the ribbon; the title bar and status bar remain.                                 |
| `hiddenActions`  |                   | Ribbon controls to hide by stable id; see the table in [bindings](/bindings).                  |
| `ribbonAddIns`   |                   | Tabs the host adds after the built-in ones; see [bindings](/bindings#ribbon-add-in-tabs).      |
| `customization`  |                   | Capabilities, host File tab pages and host menu commands; see [Customization](#customization). |
| `dirty` (read)   |                   | True after an edit, false after save, load or `markClean()`.                                   |

## Customization

`customization` takes one `WordCustomization` object (the shared contract of `ooxml-core/embedding`, also used by Excel, PowerPoint and Visio). Every binding forwards it as the `customization` prop or input. It is applied when the object identity changes, so keep it stable; `null` restores the defaults.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: {
		extraPages: [
			{
				id: 'policies',
				label: 'Policies',
				render: (container, context) => {
					container.textContent = `Policies for ${context.fileName}`;
				},
			},
		],
	},
	contextMenu: {
		extraCommands: [
			{ id: 'send', label: 'Send to chat', onSelect: ({ selectedText }) => send(selectedText) },
		],
	},
};
```

| Field                       | Effect                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| --------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `capabilities`              | `false` removes every path to the capability (`WORD_CAPABILITY_IDS`): `edit` (read-only whatever `readOnly` says), `print`, `export` (Save As, Export), `save`, `open` (New, Open), `share`, `copy` (Copy, Cut, their keys), `comment`, `review` (tracking, accept, reject), `find`, `settings` (File > Options). Ribbon controls, File tab pages, Quick Access Save, context-menu entries, shortcuts and `file-command` all follow. `hiddenActions` is unchanged; the two are unioned. |
| `backstage.extraPages`      | Host File tab pages `{ id, label, group?, order?, render }` after the built-in pages of their group (`footer`: next to Options). `render` is a mount function `(container, context) => cleanup \| { update, destroy }`, a custom element tag, or `{ element, properties }`. The body mounts the first time the page is shown, receives a `WordRegionContext` (`editor`, `fileName`, `locale`, `canEdit`, `close`) and is torn down when the page is removed or the editor disconnects.  |
| `contextMenu.extraCommands` | Host commands `{ id, label, icon?, group?, when?, disabled?, onSelect }` above (`top`, default) or below (`bottom`) the built-in entries. `onSelect` runs after the menu closes, with a `WordMenuContext` (`selectedText`, `from`, `to`, `readOnly`, `inTable`, `onLink`). Labels are shown as given (localise them yourself).                                                                                                                                                          |

`edit: false` is a read-only lock of its own beside the host's `readOnly` and a collaboration `viewer` role: none of them lifts another, and leaving a room keeps it. `readonly-change` reports it in `readOnly` with `locked: false` (only a `viewer` role sets `locked`) and leaves `hostReadOnly` alone, so a binding that mirrors `hostReadOnly` never writes the capability back. Capabilities are UI policy, not document protection. Region replacement (the Options dialog, panes) is not offered for Word yet.

## Methods

| Method                                                | Description                                                                    |
| ----------------------------------------------------- | ------------------------------------------------------------------------------ |
| `load(bytes)`                                         | Load DOCX or legacy DOC bytes. Rejects and emits `document-error` on failure.  |
| `save()`                                              | Resolves to a `Blob` of the serialized document.                               |
| `saveBytes()`                                         | Resolves to the raw `Uint8Array`.                                              |
| `download(fileName?)`                                 | Saves and starts a browser download.                                           |
| `markClean()`                                         | Clears the dirty flag after the host persisted the document.                   |
| `startCollaboration(config)`                          | Join a collaboration session; see [collaboration](/collaboration).             |
| `startYjsCollaboration(session, options)`             | Join a synchronized Yjs room with its matching source package.                 |
| `reconnectCollaboration()`, `resyncCollaboration()`   | Recover the Yjs connection or request provider state exchange.                 |
| `getPendingCollaboration()`                           | The pending step batch, or `null`.                                             |
| `receiveCollaboration(batch)`                         | Apply an accepted batch from the authority.                                    |
| `stopCollaboration(discardPending?)`                  | Leave the session; pending edits must be acknowledged or explicitly discarded. |
| `publishPresence`, `receivePresence`, `leavePresence` | Transient presence channel; see [collaboration](/collaboration#presence).      |

## Events

All events bubble and are composed, so they cross the shadow boundary.

| Event                  | Detail                   | Meaning                                                                                               |
| ---------------------- | ------------------------ | ----------------------------------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`          | The document changed through editing or review actions.                                               |
| `document-error`       | `Error`                  | Loading, saving, printing or an edit failed.                                                          |
| `document-warning`     | `string`                 | Non-fatal notice.                                                                                     |
| `readonly-change`      | `boolean`                | The user toggled Editing/Viewing in the chrome.                                                       |
| `file-command`         | `{ command, fileName? }` | Cancelable. `command` is `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` or `share`. |
| `dirty-change`         | `boolean`                | Unsaved-changes state flipped.                                                                        |
| `page-change`          | `{ page, pageCount }`    | Print Layout page or page count changed (approximate pagination).                                     |
| `ribbon-action`        | `RibbonAction`           | A ribbon control was activated.                                                                       |
| `ribbon-customize`     | `readonly string[]`      | The user changed which ribbon commands are shown; persist the ids if wanted.                          |
| `office-ribbon-add-in` | `{ tab, command }`       | A command of a host tab (`ribbonAddIns`) was chosen.                                                  |
| `presence-send`        | `PresenceMessage`        | Local presence changed; transport it to other clients.                                                |
| `collaboration-send`   | `StepBatch`              | Local steps are pending; transport them to the authority.                                             |

The package exports the typed event map (`DocxEditorEventMap`) and the runtime list `DOCX_EDITOR_EVENTS`.

The framework bindings forward the props `fileName`, `reviewAuthor` and `themeColors`, and the events `file-command`, `document-warning`, `readonly-change`, `ribbon-action` and `ribbon-customize`, under their framework names (for example `onFileCommand`, `fileCommand` and `onfilecommand`). The table is in [framework bindings](/bindings#binding-props-and-events).

## Host-owned storage

`storage` (a property of the element and a prop of every binding) takes a `DocumentStorage` from
`ooxml-core/embedding`. When it is set:

- File > Save, Save As and Ctrl+S hand the saved bytes to `storage.save(bytes, meta)` instead of
  downloading them. `meta` is `{ fileName, mimeType, reason }`, with `reason` `'save'` or
  `'saveAs'`. Once it resolves the document is marked saved; if it resolves with `{ fileName }`
  the title bar takes that name. A rejection is reported as `document-error` and the document
  stays dirty.
- File > Open calls `storage.open()` when the storage has one and opens the `{ bytes, fileName }`
  it resolves with (`undefined` means the user canceled). Without `open` the file picker is used.
- One save runs at a time: Save while a save is pending is ignored, not queued. An edit made while
  the host save runs keeps the document dirty, since the saved bytes do not contain it.
- `file-command` still fires first. A host that calls `preventDefault()` handles the command
  itself and `storage` is not called. Save As announces `saveAs` and then `save`.
- A host that removed Save (`customization.capabilities.save: false`) gets neither the `save`
  event nor `storage.save`; Save As (the `export` capability) then only downloads.

Export, Export as text and Print still use the browser. Word has no AutoRecover, so the recovery
methods of `DocumentStorage` are not used.

```ts
editor.storage = {
	async save(bytes, { fileName, mimeType }) {
		await fetch(`/files/${encodeURIComponent(fileName)}`, {
			method: 'PUT',
			headers: { 'content-type': mimeType },
			body: bytes,
		});
	},
};
```
