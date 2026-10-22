# Element API and events

Every framework adapter wraps the same `<docx-editor>` custom element. The adapters forward these properties and re-emit these events in their framework's idiom (see [framework bindings](/bindings)).

## Attributes and properties

| Property         | Attribute         | Description                                                                                  |
| ---------------- | ----------------- | -------------------------------------------------------------------------------------------- |
| `documentModel`  |                   | The `DocumentModel` shown. Assigning it is an external replacement, not an edit event.       |
| `readOnly`       | `read-only`       | Blocks local edits. Remote collaboration batches still apply.                                |
| `locale`         | `locale`          | Interface language: `en`, `fr`, `de`, `es` or `zh-CN`. Document content is never translated. |
| `fileName`       | `file-name`       | Name of the open document (default `Document1.docx`).                                        |
| `reviewAuthor`   | `review-author`   | Author recorded on tracked changes and comments (default `Author`).                          |
| `theme`          | `theme`           | `light`, `dark` or `auto`. See [theming](/theming).                                          |
| `themeColors`    |                   | Partial token overrides applied as `--dve-*` custom properties.                              |
| `showThumbnails` | `show-thumbnails` | Left rail of page thumbnails (Print Layout only).                                            |
| `showToolbar`    | `show-toolbar`    | `false` hides the ribbon; the title bar and status bar remain.                               |
| `hiddenActions`  |                   | Ribbon controls to hide by stable id; see the table in [bindings](/bindings).                |
| `ribbonAddIns`   |                   | Tabs the host adds after the built-in ones; see [bindings](/bindings#ribbon-add-in-tabs).    |
| `dirty` (read)   |                   | True after an edit, false after save, load or `markClean()`.                                 |

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
- `file-command` still fires first. A host that calls `preventDefault()` handles the command
  itself and `storage` is not called. Save As announces `saveAs` and then `save`.

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
