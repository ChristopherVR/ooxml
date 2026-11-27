# Element API and events

Every framework adapter wraps the same `<docx-editor>` custom element. The adapters forward these properties and re-emit these events in their framework's idiom (see [framework bindings](/bindings)).

## Attributes and properties

| Property                    | Attribute         | Description                                                                                                                                                                                                                                                                                                    |
| --------------------------- | ----------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `documentModel`             |                   | The `DocumentModel` shown. Assigning it is an external replacement, not an edit event.                                                                                                                                                                                                                         |
| `readOnly`                  | `read-only`       | Blocks local edits. Remote collaboration batches still apply.                                                                                                                                                                                                                                                  |
| `locale`                    | `locale`          | Interface language: `en`, `fr`, `de`, `es` or `zh-CN`. Document content is never translated.                                                                                                                                                                                                                   |
| `fileName`                  | `file-name`       | Name of the open document (default `Document1.docx`).                                                                                                                                                                                                                                                          |
| `reviewAuthor`              | `review-author`   | Author recorded on tracked changes and comments (default `Author`).                                                                                                                                                                                                                                            |
| `theme`                     | `theme`           | `light`, `dark` or `auto`. See [theming](/theming).                                                                                                                                                                                                                                                            |
| `themeColors`               |                   | Partial token overrides applied as `--dve-*` custom properties.                                                                                                                                                                                                                                                |
| `wordDesignPreferences`     |                   | Optional host `read()` and `write(defaults)` callbacks for Design > Set as Default and new blank documents. Clear the property to disable use of the preference store.                                                                                                                                         |
| `onlineVideoHosts`          |                   | Host names of the video sites Insert > Online Video may embed (a name covers its sub-domains). The command stays hidden while the list is empty; the editor never loads the address.                                                                                                                           |
| `onlineVideoSearch`         |                   | A host video search (`{ name, search(query, { signal, locale }) }`) for the Online Video search box; none by default. Only the words the user searches for are sent to it, and results whose host `onlineVideoHosts` lacks are not listed. `createYouTubeVideoSearch({ apiKey })` adapts the YouTube Data API. |
| `buildingBlocksStore`       |                   | A host store (`read()`, `write(value)`) for building blocks shared by every document, as Word's Building Blocks.dotx. Save Selection to Quick Part Gallery and the Building Blocks Organizer then offer them next to the document's own.                                                                       |
| `documentSigner`            |                   | The host's signing key (`{ certificate(), sign(bytes), label? }`; `webCryptoPackageSigner(key, certificate)` builds one). Enables Sign on signature lines: the document gets an XML digital signature (XAdES) and reopens read-only, with Edit Anyway. The editor never reads a certificate store.             |
| `showThumbnails`            | `show-thumbnails` | Left rail of page thumbnails (Print Layout only).                                                                                                                                                                                                                                                              |
| `readAloud`                 |                   | Read Aloud options. `stopOnEdit` (default `false`) also stops playback when the document is edited; replacing, loading or unmounting always stops it.                                                                                                                                                          |
| `showToolbar`               | `show-toolbar`    | `false` hides the ribbon; the title bar and status bar remain.                                                                                                                                                                                                                                                 |
| `hiddenActions`             |                   | Ribbon controls to hide by stable id; see the table in [bindings](/bindings).                                                                                                                                                                                                                                  |
| `ribbonAddIns`              |                   | Tabs the host adds after the built-in ones; see [bindings](/bindings#ribbon-add-in-tabs).                                                                                                                                                                                                                      |
| `collaboration`             |                   | Declarative collaboration (`DocxCollaborationOptions` or `null`); see [collaboration](/collaboration#the-collaboration-property).                                                                                                                                                                              |
| `customization`             |                   | Capabilities, hidden ribbon tabs (File included), host File tab pages and host menu commands; see [Customization](#customization).                                                                                                                                                                             |
| `surfaceOnly`               |                   | Mounts only the document and its editing surface. Choose it before the element connects.                                                                                                                                                                                                                       |
| `surfaceActions`            |                   | Host dialogs for surface mode: `onOpenHyperlinkDialog`, `onOpenFindReplace` and `onOpenImagePicker`.                                                                                                                                                                                                           |
| `googleFonts`               |                   | Webfonts for document fonts that are neither installed nor embedded. Unset follows File > Options > Trust Center (off until turned on); `false` turns them off; `{ baseUrl }` sets the CSS2 endpoint.                                                                                                          |
| `storage`                   |                   | Host-owned storage for Save, Save As and Open; see [Host-owned storage](#host-owned-storage).                                                                                                                                                                                                                  |
| `fileProviders`             |                   | Cloud places for File > Open; see [Cloud places in File > Open](#cloud-places-in-file-open).                                                                                                                                                                                                                   |
| `loading` (read)            |                   | True while a document loads.                                                                                                                                                                                                                                                                                   |
| `collaborationState` (read) |                   | The collaboration room: whether it is active, its room id, status, sync state and people.                                                                                                                                                                                                                      |
| `dirty` (read)              |                   | True after an edit, false after save, load or `markClean()`.                                                                                                                                                                                                                                                   |     |

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

| Field                       | Effect                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `capabilities`              | `false` removes every path to the capability (`WORD_CAPABILITY_IDS`): `edit` (read-only whatever `readOnly` says), `print`, `export` (Save As, Export), `save`, `open` (New, Open), `share`, `copy` (Copy, Cut, their keys), `comment`, `review` (tracking, accept, reject), `find`, `settings` (File > Options). Ribbon controls, File tab pages, Quick Access Save, context-menu entries, shortcuts and `file-command` all follow. `hiddenActions` is unchanged; the two are unioned.                                                                                                                                                                                                                                               |
| `ribbon.hiddenTabs`         | Native ribbon tabs to remove (`WORD_RIBBON_TAB_IDS`), contextual tabs included. `file` removes the File tab and its backstage at every width: no key tip (Alt, F) or other UI path opens it, the status bar compatibility notes go with it, and an open backstage closes and returns focus to the ribbon. Save, Ctrl+S, Ctrl+P, Quick Access Save, `storage` and `file-command` keep working unless a capability removes them. A hidden selected tab hands over to the first remaining one; host tabs (`ribbonAddIns`) are not affected. The signature badge and message bar keep showing the verification state as plain text, and Tell me leaves out commands that live only on hidden tabs. Changes apply at runtime per instance. |
| `backstage.extraPages`      | Host File tab pages `{ id, label, group?, order?, render }` after the built-in pages of their group (`footer`: next to Options). `render` is a mount function `(container, context) => cleanup \| { update, destroy }`, a custom element tag, or `{ element, properties }`. The body mounts the first time the page is shown, receives a `WordRegionContext` (`editor`, `fileName`, `locale`, `canEdit`, `close`) and is torn down when the page is removed or the editor disconnects.                                                                                                                                                                                                                                                |
| `contextMenu.extraCommands` | Host commands `{ id, label, icon?, group?, when?, disabled?, onSelect }` above (`top`, default) or below (`bottom`) the built-in entries. `onSelect` runs after the menu closes, with a `WordMenuContext` (`selectedText`, `from`, `to`, `readOnly`, `inTable`, `onLink`). Labels are shown as given (localise them yourself).                                                                                                                                                                                                                                                                                                                                                                                                        |

`edit: false` is a read-only lock of its own beside the host's `readOnly` and a collaboration `viewer` role: none of them lifts another, and leaving a room keeps it. `readonly-change` reports it in `readOnly` with `locked: false` (only a `viewer` role sets `locked`) and leaves `hostReadOnly` alone, so a binding that mirrors `hostReadOnly` never writes the capability back. Capabilities are UI policy, not document protection. Region replacement (the Options dialog, panes) is not offered for Word yet. All products side by side are mapped in the [embedding guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Methods

| Method                                                          | Description                                                                                                                                                                    |
| --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `load(bytes)`                                                   | Load DOCX or legacy DOC bytes, or a PDF, which is converted into an editable document (see [PDF import](#pdf-import)). Rejects and emits `document-error` on failure.          |
| `save()`                                                        | Resolves to a `Blob` of the serialized document.                                                                                                                               |
| `saveBytes()`                                                   | Resolves to the raw `Uint8Array`.                                                                                                                                              |
| `download(fileName?)`                                           | Saves and starts a browser download.                                                                                                                                           |
| `exportPdf(options?)`                                           | Resolves to `{ bytes, blob, report }`: a PDF with real, selectable text and vector borders, paginated as Print Layout shows it, with no dialog. See [PDF export](#pdf-export). |
| `markClean()`                                                   | Clears the dirty flag after the host persisted the document.                                                                                                                   |
| `setLoadedDocument(session)`                                    | Mounts a `LoadedDocument` already parsed by `loadDocument`.                                                                                                                    |
| `openCloudFile(file)`                                           | Opens a file from a `CloudFileProvider`, as File > Open > OneDrive does.                                                                                                       |
| `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` | State and callbacks for host-built toolbar, page and status controls.                                                                                                          |
| `subscribeEditorState(listener)`                                | Calls `listener` on selection and formatting changes, including ones that do not dirty the document. Returns an unsubscribe function.                                          |
| `startCollaboration(config)`                                    | Join a collaboration session; see [collaboration](/collaboration).                                                                                                             |
| `startYjsCollaboration(session, options)`                       | Join a synchronized Yjs room with its matching source package.                                                                                                                 |
| `reconnectCollaboration()`, `resyncCollaboration()`             | Recover the Yjs connection or request provider state exchange.                                                                                                                 |
| `getPendingCollaboration()`                                     | The pending step batch, or `null`.                                                                                                                                             |
| `receiveCollaboration(batch)`                                   | Apply an accepted batch from the authority.                                                                                                                                    |
| `stopCollaboration(discardPending?)`                            | Leave the session; pending edits must be acknowledged or explicitly discarded.                                                                                                 |
| `publishPresence`, `receivePresence`, `leavePresence`           | Transient presence channel; see [collaboration](/collaboration#presence).                                                                                                      |

### PDF import {#pdf-import}

`load()` and File > Open accept a PDF and convert it the way Word's PDF Reflow does: the text becomes flowing paragraphs with their fonts, sizes, bold, italic, underline, colours and super- and subscripts, alignment, indents and spacing; headings come from the PDF's structure tags and bookmarks when it has them, else larger short lines become headings; text repeated at the top or bottom of the pages becomes the header and footer, with page numbers as PAGE fields; notes at the foot of a page that a mark in the text refers to become footnotes; links inside the document (a table of contents, cross-references) go to bookmarks at their targets, and dot leaders become tab stops; rotated text goes into text boxes turned like the original; pictures (also inline and fax-coded ones, with their alternate text) and external links stay. Ruled tables become Word tables (merged cells, borders and cell shading kept; a table continued on the next page is one table, its repeated header row a header row), and text aligned in columns over three or more rows a borderless table. Typed bullets and numbers become real lists, with levels from their indents. Two- and three-column regions are read column by column and become a section with that many columns. The document then saves as `.docx` (its name changes to `.docx`) and the PDF is not changed. A PDF that needs a password to open is refused; one that only restricts permissions opens. `importPdf(bytes, options)` from `ooxml-core/docx/pdf-import` does the same outside the editor and returns a report per page.

Vector drawings (shapes, lines and charts the PDF draws as paths) become floating Word drawings behind the text at the place the PDF drew them: rectangles, ellipses and lines as preset shapes, other paths as custom geometry, with their fill, outline, opacity and dashes; touching pieces stay one shape or group, and their text stays ordinary text. Shadings (gradients), pattern fills, shapes a clipping path cuts and shapes over a page's limit (500 shapes) are left out and reported.

This version does not rebuild borderless tables whose cells wrap or lists inside table cells, JPEG 2000 and JBIG2 pictures are decoded and kept as PNG (a picture whose data cannot be decoded is left out and reported). Scanned pages stay pictures unless `importPdf` is given an `ocr` provider.

### PDF export {#pdf-export}

`exportPdf()` (and File > Export > Create a PDF, which first shows an Options step like Word's "Publish as PDF" Options dialog, then downloads the file) writes the document the way Word's Save As PDF does: real text you can select, search and copy, vector table borders, shading and rules, charts and drawing shapes as vector paths with real text, text boxes, watermarks, page borders, line numbers, floating objects (in the body, headers, footers and table cells), and pictures at their own resolution, on the pages Print Layout shows (it paints the same layout result). No page is turned into a picture and no print dialog opens. Outside the editor, `exportDocxPdf(model, options)` from `ooxml-ui/docx` does the same for any `DocumentModel` (pass `media` for its pictures, and a `measurer` to run without a canvas).

| Option               | Default                             | Effect                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| -------------------- | ----------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `pages`              | `'all'`                             | `{ from, to }` (1-based, inclusive) writes a range. In the editor, `'current'` writes the page Print Layout shows (only in Print Layout) and `'selection'` only the selected body content, laid out on its own pages.                                                                                                                                                                                                                                                                                                                                   |
| `markup`             | `'none'`                            | `'none'` writes the document with changes accepted (Word's "Document"); `'all'` shows tracked changes as All Markup does (insertions underlined, deletions struck through in each author's colour, change bars) and comments as PDF notes.                                                                                                                                                                                                                                                                                                              |
| `markupBalloons`     | `'commentsAndFormatting'`           | Word's Balloons setting for `markup: 'all'`. Pages are shrunk to leave a markup pane on the right, as Word prints them, with balloons (real text) joined to their text by dotted lines in the author's colour; comments stay PDF notes too. `'commentsAndFormatting'` puts comments (with replies, author and date) and formatting changes ("Formatted: Font: Bold") in balloons; `'all'` also moves deletions to balloons ("Deleted: ..."); `'comments'` shows only comments in balloons (Word's "Show all revisions inline"); `'none'` draws no pane. |
| `bookmarks`          | `'headings'`                        | Word's "Create bookmarks using": `'headings'` makes the headings the PDF's bookmarks, nested by level; `'wordBookmarks'` lists the document's bookmarks (not hidden ones such as `_Toc`); `'none'` writes none.                                                                                                                                                                                                                                                                                                                                         |
| `taggedPdf`          | `true`                              | Word's "Document structure tags for accessibility": headings, paragraphs, tables, figures (with their alternative text) and links in reading order; headers, footers and repeated table headings are artifacts.                                                                                                                                                                                                                                                                                                                                         |
| `pdfA`               | `false`                             | Writes PDF/A-2 metadata (XMP), an sRGB output intent and the file identifier, and writes text in families no supplied font is named for with a supplied font. With every font embedded the file declares and meets PDF/A-2u (validated with veraPDF); otherwise it declares PDF/A-2b and does not conform. `report.pdfA` says which (`conformant`, `level`, `issues`).                                                                                                                                                                                  |
| `fonts`              | the page's local `@font-face` files | TrueType files (`{ family, bold?, italic?, data }`) embedded as subsets; other text uses the standard PDF fonts.                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `documentProperties` | `true`                              | Writes the title, author, subject and keywords.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| `language`           |                                     | The document language (`/Lang`), e.g. `en-US`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `signal`             |                                     | An `AbortSignal` that stops the export.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |

The Options step offers the page range (All, Current page in Print Layout, Selection when text is selected, or Pages from and to), Publish what (Document or Document showing markup, with the Balloons choice for the markup), Create bookmarks using (Headings or Word bookmarks, or none when unticked), Document properties, Document structure tags for accessibility and PDF/A compliant; Export maps them onto the options above, and Cancel or Escape exports nothing. The choices are kept for the next export in the same session, and the range starts at All each time. Word's "Optimize for" has no counterpart: pictures always keep their own resolution. A host that cancels the `file-command` event for `exportPdf` gets it before the step opens.

The `report` lists the fonts embedded and those written with a standard font instead, characters no font could write, and what was left out by page. Charts and shapes are written as vector paths and real text in the browser and in Node alike (`exportDocxPdf` needs no DOM: their drawing is read with a DOM-free XML parser where there is no `DOMParser`), and a picture fill inside one follows the same rules as any picture (JPEG and PNG are written directly, other formats need the host's decoder). Shadows lose their blur, WordArt warps and text effects are written as plain text, and a text watermark is sized as Print Layout shows it rather than stretched to its shape; the report notes each. Hyperlinks, links to bookmarks, cross-references and table of contents entries made as hyperlinks (`\h`) stay clickable; a link whose bookmark is on no page written stays plain text (noted in the report). Lists are tagged as paragraphs, not `L`/`LI`. Remote fonts are never fetched: supply font files to embed the document's typefaces. Markup balloons describe run and paragraph formatting changes; table, row, cell and section property changes have no balloon or change bar (the document model keeps them only as source XML), and balloons that do not all fit beside a page overlap at the top of the pane (noted in the report).

## Events

All events bubble and are composed, so they cross the shadow boundary.

| Event                  | Detail                               | Meaning                                                                                                            |
| ---------------------- | ------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| `document-change`      | `DocumentModel`                      | The document changed through editing or review actions.                                                            |
| `document-error`       | `Error`                              | Loading, saving, printing or an edit failed.                                                                       |
| `document-warning`     | `string`                             | Non-fatal notice.                                                                                                  |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }` | The user toggled Editing/Viewing in the chrome.                                                                    |
| `file-command`         | `{ command, fileName? }`             | Cancelable. `command` is `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `exportPdf`, `print` or `share`. |
| `dirty-change`         | `boolean`                            | Unsaved-changes state flipped.                                                                                     |
| `page-change`          | `{ page, pageCount }`                | Print Layout page or page count changed (approximate pagination).                                                  |
| `ribbon-action`        | `RibbonAction`                       | A ribbon control was activated.                                                                                    |
| `ribbon-customize`     | `readonly string[]`                  | The user changed which ribbon commands are shown; persist the ids if wanted.                                       |
| `office-ribbon-add-in` | `{ tab, command }`                   | A command of a host tab (`ribbonAddIns`) was chosen.                                                               |
| `presence-send`        | `PresenceMessage`                    | Local presence changed; transport it to other clients.                                                             |
| `collaboration-send`   | `StepBatch`                          | Local steps are pending; transport them to the authority.                                                          |
| `collaboration-change` | `DocxCollaborationState`             | The collaboration state changed: `{ active, roomId, status, synced, people }`.                                     |

The package exports the typed event map (`DocxEditorEventMap`) and the runtime list `DOCX_EDITOR_EVENTS`.

The framework bindings forward the props `fileName`, `reviewAuthor` and `themeColors`, and the events `file-command`, `document-warning`, `readonly-change`, `ribbon-action` and `ribbon-customize`, under their framework names (for example `onFileCommand`, `fileCommand` and `onfilecommand`). The table is in [framework bindings](/bindings#binding-props-and-events).

## Shadow parts

The title bar (file name, save state, quick access and "Tell me") is the shadow part `title-bar`.
A host that draws its own document bar hides it with `docx-editor::part(title-bar) { display: none; }`;
the ribbon and the status bar stay.

## Cloud places in File > Open

`fileProviders` (a property of the element and a prop of every binding) takes
`CloudFileProvider`s from `ooxml-core/embedding`, for example `createOneDriveFileProvider`. File >
Open then lists "This device" and one place per provider. A file opened from a place saves through
its own storage and is read-only when the provider says it cannot be edited; see
[OneDrive integration](/onedrive#open-from-onedrive-in-file-open).

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

## Read Aloud {#read-aloud}

The building-blocks hook (`useViewerBuildingBlocks`) puts five Read Aloud fields on `toolbarProps`.
They are not element properties or events. They use the editor's own implementation (the browser's
speech synthesis; nothing is sent anywhere), and the [building-blocks guide](/building-blocks#read-aloud)
describes how they behave.

| Field             | Type                     | Behaviour                                                                                                                                                    |
| ----------------- | ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `canReadAloud`    | `boolean`                | True when a document is on screen and the browser can speak. Read-only documents count.                                                                      |
| `readingAloud`    | `boolean`                | True only in the editor that owns the voice. Another editor reading leaves it false.                                                                         |
| `readAloudIssue`  | `ReadAloudIssue \| null` | The last problem: `'unsupported'` while the browser has no speech synthesis, `'empty'` or `'error'`. `null` when there is none.                              |
| `onReadAloud`     | `() => ReadAloudStart`   | Reads the selection, or from the caret to the end when nothing is selected. Restarts if already reading. Returns `'started'`, or the reason nothing started. |
| `onStopReadAloud` | `() => boolean`          | Stops this editor's playback. Returns whether it was playing.                                                                                                |

The codes are plain string unions. The library supplies the codes only; the host words the message.

```ts
type ReadAloudIssue = 'unsupported' | 'empty' | 'error';
type ReadAloudStart = 'started' | ReadAloudIssue;
```

`'error'` means the voice failed. A cancelled or interrupted utterance is not reported as an error.

A React host button that toggles reading and words the empty case:

```tsx
import { DocumentCanvas, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel }) {
	const { canvasProps, toolbarProps } = useViewerBuildingBlocks({ documentModel });
	return (
		<>
			{toolbarProps && toolbarProps.readAloudIssue !== 'unsupported' && (
				<>
					<button
						disabled={!toolbarProps.canReadAloud}
						aria-pressed={toolbarProps.readingAloud}
						onClick={() =>
							toolbarProps.readingAloud
								? toolbarProps.onStopReadAloud()
								: toolbarProps.onReadAloud()
						}
					>
						{toolbarProps.readingAloud ? 'Stop reading' : 'Read aloud'}
					</button>
					{toolbarProps.readAloudIssue === 'empty' && <span role="status">Nothing to read.</span>}
				</>
			)}
			<DocumentCanvas {...canvasProps} />
		</>
	);
}
```

Playback stops when the document is replaced or reloaded and when the editor unmounts; an ordinary edit stops it only with `readAloud: { stopOnEdit: true }`. The
optional `Toolbar` draws its own "Read aloud" toggle from the same fields.
