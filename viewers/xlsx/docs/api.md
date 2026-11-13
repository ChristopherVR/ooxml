# Element API

`<xlsx-editor>` (class `XlsxEditorElement`) is the one editor every package mounts. Register it with `defineXlsxEditor()` (exported by every package; registration is idempotent and does nothing without `customElements`), or let `mountEditor` and the framework components do it for you.

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const editor = document.createElement('xlsx-editor');
editor.fileName = 'Budget.xlsx';
document.body.append(editor);
await editor.load(bytes, 'Budget.xlsx');
```

## Attributes and properties

| Attribute          | Property         | Type                                                  | Default      |
| ------------------ | ---------------- | ----------------------------------------------------- | ------------ |
| `locale`           | `locale`         | `en \| fr \| de \| es \| zh-CN` (any tag maps to one) | `en`         |
| `read-only`        | `readOnly`       | boolean                                               | `false`      |
| `file-name`        | `fileName`       | string                                                | `Book1.xlsx` |
| `author-name`      | `authorName`     | string                                                | `Author`     |
| `theme`            | `theme`          | `light \| dark \| auto`                               | `auto`       |
| `show-toolbar`     | `showToolbar`    | boolean (`"false"` hides)                             | `true`       |
| `show-formula-bar` | `showFormulaBar` | boolean (`"false"` hides)                             | `true`       |

Properties only:

| Property             | Type                                 | Notes                                                                                                                                                        |
| -------------------- | ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `themeColors`        | `Partial<Record<...>>`               | Token overrides applied as `--xve-*` custom properties ([theming](/theming)).                                                                                |
| `validateReplace`    | `boolean`                            | Replace and Replace All check data validation; refused cells stay unchanged (default `false`, like Excel).                                                   |
| `hiddenActions`      | `string[]`                           | Ribbon controls, title-bar items and context-menu entries to hide, by stable id.                                                                             |
| `ribbonAddIns`       | `RibbonAddInTab[]`                   | Tabs the host adds after the built-in ones ([bindings](/bindings#ribbon-add-in-tabs)).                                                                       |
| `workbook`           | `Workbook \| null`                   | Get the open workbook, or set one to replace it.                                                                                                             |
| `dirty`              | `boolean` (read only)                | Unsaved edits.                                                                                                                                               |
| `activeSheet`        | `number`                             | Zero-based index of the visible sheet.                                                                                                                       |
| `collaboration`      | `XlsxCollaborationOptions \| null`   | Shares the open workbook in a room while set ([collaboration](/collaboration)).                                                                              |
| `collaborationState` | `XlsxCollaborationState` (read only) | Whether sharing, the connection status and who is in the room.                                                                                               |
| `customization`      | `ExcelCustomization \| null`         | Capabilities, host File tab pages and host menu commands ([below](#customization)).                                                                          |
| `googleFonts`        | `false \| { baseUrl }`               | Webfont fallback for fonts the reader has neither installed nor embedded. Unset, the reader's Trust Center choice applies (off by default).                  |
| `surfaceOnly`        | boolean                              | Set before mounting to omit application UI; pair it with `surfaceActions`.                                                                                   |
| `surfaceActions`     | `SurfaceHostActions`                 | Host callbacks that stand in for dialogs when `surfaceOnly` is set. A workflow without a callback is reported as not enabled in `getToolbarProps().enabled`. |
| `loading`            | `boolean` (read only)                | Whether a workbook is being opened.                                                                                                                          |

## Customization

`customization` takes one `ExcelCustomization` object, the shared contract of `ooxml-core/embedding` that Word, PowerPoint and Visio take too. Every binding forwards it as the `customization` prop or input. It is applied when the object identity changes, so keep it stable; `null` restores the defaults.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send range', onSelect: ({ ref }) => send(ref) }],
	},
};
```

| Field                       | Effect                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `capabilities`              | `false` removes every path to the capability (`EXCEL_CAPABILITY_IDS`): `edit` (read-only whatever `readOnly` says), `open` (New, Open), `save`, `export` (Save As, Export), `print`, `share`, `settings` (Options), `copy` (Copy, Cut), `comment` (New and Delete Comment) and `find` (Find, Replace). The commands are hidden through the same registry as `hiddenActions` (ribbon, Tell me, context menus, shortcuts, title bar) and refuse to run, and their File tab pages go. `hiddenActions` is unchanged; the two are unioned. |
| `backstage.extraPages`      | Host File tab pages `{ id, label, group?, order?, render }` after the built-in pages of their group. `render` is a mount function `(container, context) => cleanup \| { update, destroy }`, a custom element tag or `{ element, properties }`; it mounts the first time the page is shown, receives an `ExcelRegionContext` (`editor`, `fileName`, `locale`, `canEdit`, `activeSheet`, `close`) and is torn down when the page is removed or the editor disconnects.                                                                  |
| `contextMenu.extraCommands` | Host commands `{ id, label, icon?, group?, when?, disabled?, onSelect }` in the cell, row header, column header and sheet-tab menus, above (`top`, default) or below (`bottom`) the built-in entries. `onSelect` runs after the menu closes with an `ExcelMenuContext` (`menu`: `cell`, `row`, `column` or `sheetTab`; `sheet`, `sheetName`, `ref`, `active`, `readOnly`); use `when` to offer a command in some menus only. Labels are shown as given.                                                                               |

`edit: false` is a read-only lock of its own beside the host's `readOnly` and a collaboration `viewer` role: none of them lifts another, and leaving a room keeps it. `readonly-change` reports it in `readOnly` with `locked: false` (only a `viewer` role sets `locked`) and leaves `hostReadOnly` alone, so a binding that mirrors `hostReadOnly` never writes the capability back. Capabilities are UI policy, not workbook protection. Region replacement (the Options dialog, panes) is not offered for Excel yet. All products side by side are mapped in the [embedding guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Methods

| Method                           | Returns                   | Notes                                                                                                                                                                                               |
| -------------------------------- | ------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`         | `Promise<void>`           | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv`; the format is sniffed from the bytes. Emits `workbook-error` and rejects on failure.                                                                     |
| `newWorkbook()`                  | `void`                    | A blank workbook named `Book1.xlsx`.                                                                                                                                                                |
| `save()`                         | `Promise<Blob>`           | `.xlsx`. Does not clear `dirty`.                                                                                                                                                                    |
| `saveBytes(format?)`             | `Promise<Uint8Array>`     | `'xlsx'` (default) or `'csv'` (the active sheet, formatted values).                                                                                                                                 |
| `download(name?)`                | `Promise<void>`           | Saves, downloads, then `markClean()`.                                                                                                                                                               |
| `markClean()`                    | `void`                    |                                                                                                                                                                                                     |
| `select(ref)`                    | `void`                    | An A1 cell or range on the active sheet (`'C3'`, `'A1:D20'`).                                                                                                                                       |
| `getSelection()`                 | `string`                  | The current selection as an A1 reference.                                                                                                                                                           |
| `setActiveSheet(index)`          | `void`                    |                                                                                                                                                                                                     |
| `undo()`, `redo()`               | `void`                    | While shared, only this window's edits are undone.                                                                                                                                                  |
| `startCollaboration(options)`    | `Promise<void>`           | Joins a room; rejects when already sharing or no workbook is open.                                                                                                                                  |
| `stopCollaboration()`            | `void`                    | Leaves the room and clears `collaboration`; the workbook stays as it is.                                                                                                                            |
| `reconnectCollaboration()`       | `void`                    | Restarts the connection, keeping changes made offline.                                                                                                                                              |
| `share()`                        | `void`                    | Opens the Share dialog.                                                                                                                                                                             |
| `focusGrid()`                    | `void`                    | Moves keyboard focus to the grid so typing edits the active cell.                                                                                                                                   |
| `openCloudFile(file)`            | `Promise<void>`           | Opens a file a host got from a `CloudFileProvider`, as File > Open > OneDrive does: read-only when `file.canEdit` is false, Save and Save As routed to `file.storage`. Rejects when the load fails. |
| `commitEdit()`                   | `boolean`                 | Commits an in-progress cell edit before a host navigates or persists. `false` means validation blocked it.                                                                                          |
| `getToolbarProps()`              | `SpreadsheetToolbarProps` | The headless toolbar state and `enabled` flags for custom controls ([building blocks](/building-blocks)).                                                                                           |
| `subscribeEditorState(listener)` | `() => void`              | Calls `listener` on each editor state change; returns the unsubscribe function.                                                                                                                     |

## Events

All events are `CustomEvent`s that bubble and cross the shadow boundary (`composed`).

| Event                  | `detail`                                                                                              | Notes                                                                                                                     |
| ---------------------- | ----------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                                        | After an edit, a load or a new workbook.                                                                                  |
| `workbook-error`       | `{ error, message }`                                                                                  | A load or save failed.                                                                                                    |
| `workbook-warning`     | `{ warnings }`                                                                                        | Content the loader read but could not model.                                                                              |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }`                                                                  | The effective read-only state. `locked` is set by a collaboration `viewer` role; `hostReadOnly` is the host's own switch. |
| `ribbon-action`        | `{ id }`                                                                                              | A ribbon command ran.                                                                                                     |
| `ribbon-customize`     | `{ hiddenActions }`                                                                                   | The user changed which ribbon commands are shown (File > Options > Customize Ribbon).                                     |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' \| 'share' }` | **Cancelable**: `preventDefault()` to handle it yourself.                                                                 |
| `selection-change`     | `{ sheet, ref, active }`                                                                              | `ref` is the selected range, `active` the active cell.                                                                    |
| `sheet-change`         | `{ index, name }`                                                                                     |                                                                                                                           |
| `dirty-change`         | `{ dirty }`                                                                                           |                                                                                                                           |
| `office-ribbon-add-in` | `{ tab, command }`                                                                                    | A command of a host tab (`ribbonAddIns`) was chosen.                                                                      |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                                          | Sharing started or stopped, the connection changed, or someone joined or left.                                            |
| `validation-error`     | `{ message, title?, ref, style? }`                                                                    | A cell edit in progress failed validation and stays open until corrected or canceled.                                     |

The framework bindings forward `file-command`, `workbook-warning`, `ribbon-action` and `sheet-change` as well (for example `onFileCommand`, `fileCommand` and `onfilecommand`); see [framework bindings](/bindings#callbacks-and-events).

```ts
editor.addEventListener('file-command', (event) => {
	if (event.detail.command === 'save') {
		event.preventDefault();
		void editor
			.save()
			.then(uploadToMyStorage)
			.then(() => editor.markClean());
	}
});
```

## Host-owned storage

`storage` (a property of the element and a prop of every binding) takes a `DocumentStorage` from
`ooxml-core/embedding`. When it is set:

- File > Save, Save As and Ctrl+S hand the saved workbook to `storage.save(bytes, meta)` instead of
  downloading it. `meta` is `{ fileName, mimeType, reason }`, with `reason` `'save'` or `'saveAs'`;
  `.xls` and `.csv` sources are saved as `.xlsx`, as a download is. Once it resolves the workbook
  is marked clean; if it resolves with `{ fileName }` the title bar takes that name. A rejection is
  reported as `workbook-error` and the workbook stays dirty.
- File > Open calls `storage.open()` when the storage has one and opens the `{ bytes, fileName }`
  it resolves with (`undefined` means the user canceled). Without `open` the file picker is used.
- One save runs at a time: Save while a save is pending is ignored, not queued. An edit made while
  the host save runs keeps the workbook dirty, since the saved bytes do not contain it.
- `file-command` still fires first; a host that calls `preventDefault()` handles the command
  itself and `storage` is not called.

Export, Save as CSV and Print still download or print in the browser. Excel has no AutoRecover,
so the recovery methods of `DocumentStorage` are not used.

```ts
editor.storage = {
	save: (bytes, { fileName }) => myDrive.put(fileName, bytes),
	open: () => myDrive.pick(), // resolves { bytes, fileName } or undefined
};
```

## Cloud places in File > Open

`fileProviders` (a property of the element and a prop of every binding) takes an array of
`CloudFileProvider` (`ooxml-core/embedding`), for example `createOneDriveFileProvider` from
`ooxml-core/embedding/onedrive`. Empty or omitted, File > Open is unchanged. With providers, File >
Open lists "This device" (Browse) and one place per provider, which shows the shared
`<office-ui-cloud-files>` browser limited to `.xlsx`. Its strings follow the editor `locale`.

- Opening a file asks to discard unsaved changes, as Open does, then loads the bytes through the
  normal load path, takes the file's name and closes the File view.
- The workbook then saves through the file's own `storage`, not the host `storage`: Save and Ctrl+S
  write to the same item, Save As writes a copy, and the workbook is marked clean only after the
  provider resolved. A rejected save is reported as `workbook-error` and the workbook stays dirty.
  Opening, creating or loading another workbook returns Save to the host `storage`.
- A file the provider reports as `canEdit: false` opens read-only (`readOnly` is true, a
  `readonly-change` event fires). The stricter of the host's `readOnly` and the file wins, and the
  host's own value is not changed.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken,
		signIn,
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

## Host menus and workflows {#host-menus-and-workflows}

A host-owned surface (the [building blocks](/building-blocks), or `<xlsx-editor>` with `surfaceOnly` and `surfaceActions`) has no context menus, Format Cells, Find and Replace or Insert Function dialogs of its own. The host draws them through the callbacks below, or through the operations on `toolbarProps`. The behaviour is described in [Context menus and workflow operations](/building-blocks#context-menus-and-workflow-operations); this section gives the types and the refusal conditions.

The callbacks are set on `surfaceActions`, or passed as options to `useViewerBuildingBlocks`, `createViewerBuildingBlocks` and `mountEditorSurface`.

### Context menu callbacks {#context-menu-callbacks}

```ts
onContextMenu?: (request: ContextMenuRequest) => boolean | void;
onContextMenuItems?: (request: ContextMenuRequest) => ContextMenuCustomization | void;
```

`onContextMenu` runs for a right click or Shift+F10 on a cell, a row header, a column header or a sheet tab. Return `true` to suppress the native menu and draw your own; any other return value shows the native menu. `onContextMenuItems` runs only when `onContextMenu` did not return `true`. It filters and extends the native menu, which still draws.

```ts
type ContextMenuKind = 'cell' | 'row' | 'column' | 'tab';

interface SurfaceActionContext {
	sheet: number;
	sheetName: string;
	selection: string; // A1 text
	activeCell: string; // A1 text
}

interface ContextMenuItem {
	id: string;
	label: string; // localized
	command?: string;
	enabled: boolean;
	checked?: boolean;
	shortcut?: string;
	separatorBefore?: boolean;
}

interface ContextMenuModel extends SurfaceActionContext {
	menu: ContextMenuKind;
	menuSheet: number; // the clicked tab for 'tab', the active sheet otherwise
	readOnly: boolean;
	items: ContextMenuItem[];
	execute(id: string): Promise<boolean>;
}

interface ContextMenuRequest extends ContextMenuModel {
	x: number; // client coordinates of the right click
	y: number;
}

interface HostContextMenuItem {
	id: string; // unique among the host items of one menu; an empty or repeated id is ignored
	label: string;
	enabled?: boolean | (() => boolean); // a function is evaluated again when the item runs
	requiresEdit?: boolean; // disabled, and refused when run, in a read-only workbook
	checked?: boolean;
	separatorBefore?: boolean;
	position?: 'top' | 'bottom'; // default 'top'
	onSelect(request: ContextMenuRequest): void | Promise<void>;
}

interface ContextMenuCustomization {
	hide?: readonly string[]; // native item ids from request.items
	add?: readonly HostContextMenuItem[];
}
```

`request.execute(id)` resolves `true` when the item ran. It resolves `false` when the cell edit in progress fails validation (the edit stays open), when `id` is not in `items`, or when the item cannot run at that moment (read-only, sheet protection, or a command that is not available in surface mode). A tab menu activates its sheet first. Host items from `onContextMenuItems` run `onSelect(request)` after the menu closes, unless they are disabled by then.

### toolbarProps.contextMenu {#toolbarprops-contextmenu}

```ts
contextMenu(menu: ContextMenuKind, sheet?: number): ContextMenuModel;
```

Returns the same model without a right click, for a host that draws its own sheet tabs or toolbar. `sheet` defaults to the active sheet and is the tab a `'tab'` menu is for. Items whose command is not available in surface mode are listed with `enabled: false`. Format Cells is enabled when `onOpenFormatCells` is installed and the workbook is editable. Entries that run an action of the bundled tab strip, such as Rename, are listed disabled.

### Workflow callbacks {#workflow-callbacks}

```ts
onOpenFormatCells?: (request: FormatCellsRequest) => void;
onOpenFindReplace?: (request: FindReplaceRequest) => void;
onOpenInsertFunction?: (request: InsertFunctionRequest) => void;

interface FormatCellsRequest extends SurfaceActionContext, FormatCellsOperations {}
interface FindReplaceRequest extends SurfaceActionContext, FindOperations {}
interface InsertFunctionRequest extends SurfaceActionContext, InsertFunctionOperations {}
```

Each request carries the context above plus the operations of that workflow, described below. The callback runs only when the workflow is available: Format Cells and Insert Function need edit rights, Find and Replace does not. The pending cell edit is committed first. The `onOpen*` functions on `toolbarProps` return a `Promise<boolean>`, which is `false` when the callback is missing, the edit rights are missing, or the pending edit fails validation.

### toolbarProps.workflows {#toolbarprops-workflows}

`toolbarProps.workflows` holds the same operations without a callback, so a panel that is always open needs no `onOpen*` callback. Each operation re-reads read-only state, sheet protection and the pending cell edit when it runs.

#### Format Cells {#format-cells-operations}

```ts
interface FormatCellsOperations {
	getState(): FormatCellsState;
	canApply(): boolean;
	apply(patch: StylePatch): boolean;
}

interface FormatAspect<T> {
	value: T; // the properties every examined cell agrees on
	mixed: string[]; // the names of the properties that differ
}

interface FormatCellsState {
	font: FormatAspect<Font>;
	alignment: FormatAspect<Alignment>;
	border: FormatAspect<Border>;
	protection: FormatAspect<Protection>;
	numFmt: string | null; // null when the selection mixes formats
	fill: Fill | null; // null when the selection mixes fills
	complete: boolean; // false when the selection has more than 4096 cells
}
```

- `getState()` returns the formatting of the selection. When `complete` is `false`, the mixed lists are partial.
- `canApply()` is true when the workbook is editable and sheet protection allows formatting.
- `apply(patch)` applies a `StylePatch` to the selection as one undo step. It returns `false`, with nothing changed, when `canApply()` is false or the pending cell edit fails validation.

#### Find and Replace {#find-and-replace-operations}

```ts
interface FindQuery {
	text: string;
	sheet?: number;
	range?: CellRange; // only with sheet
	matchCase?: boolean;
	wholeCell?: boolean;
	lookIn?: 'formulas' | 'values' | 'comments';
	wildcards?: boolean; // default true
	order?: 'rows' | 'columns';
}

interface FindMatch {
	sheet: number;
	row: number;
	col: number;
	text: string;
}

interface FindOperations {
	canReplace(): boolean;
	findAll(query: FindQuery): FindMatch[];
	findNext(query: FindQuery, options?: { backwards?: boolean }): FindMatch | undefined;
	select(match: FindMatch): boolean;
	replace(query: FindQuery, replacement: string): boolean;
	replaceAll(query: FindQuery, replacement: string): number;
}
```

- `canReplace()` is true when the workbook is editable. Finding never needs edit rights.
- `findAll(query)` returns every match and works in a read-only workbook. It returns `[]` for an empty `text`, or when the editor is not connected or is loading.
- `findNext(query, { backwards })` selects the next (or previous) match after the active cell, wrapping. It returns `undefined` when there is no match or the pending cell edit fails validation.
- `select(match)` selects a match from `findAll`. It returns `false` when the editor is not ready or the sheet does not exist.
- `replace(query, replacement)` replaces the match at the active cell, if there is one, and selects the next match. It returns `true` only when a cell was replaced. It returns `false` when the workbook is read-only, the pending edit fails validation, or the active cell is a locked cell on a protected sheet. Like Excel's own Replace, it does not apply data validation.
- `replaceAll(query, replacement)` returns how many cells changed. It returns `0` when the workbook is read-only or the pending edit fails validation. Cells locked by sheet protection are skipped. Data validation is not applied.

#### Insert Function {#insert-function-operations}

```ts
interface FunctionDoc {
	name: string;
	category: string;
	syntax: string; // for example SUM(number1, [number2], ...)
	description: string; // localized
}

interface InsertFunctionOperations {
	categories(): string[];
	search(options?: { text?: string; category?: 'all' | 'recent' | string }): FunctionDoc[];
	canInsert(): boolean;
	insert(name: string, args?: string): { text: string; caret: number } | undefined;
}
```

- `categories()` returns the distinct category names, sorted.
- `search(options)` matches `text` against the name first, then the description words. `category` is `'all'` (the default), `'recent'` for recently used functions, or a category name.
- `canInsert()` is true when the workbook is editable.
- `insert(name, args?)` inserts `NAME(`, or `NAME(args)` when `args` is given, at the caret of the open formula draft. When no draft is open it starts `=NAME(`. The insertion never goes before the leading `=`. It returns the draft `text` and `caret` afterwards, so input can continue from there. It returns `undefined`, and changes nothing, when the workbook is read-only, `name` is not a documented function, or the open draft is plain text (one that does not start with `=`).

### toolbarProps.enabled flags {#toolbarprops-enabled}

| Flag                 | True when                                                                                                    |
| -------------------- | ------------------------------------------------------------------------------------------------------------ |
| `contextMenu`        | `onContextMenu` is installed and the editor is connected.                                                    |
| `find`               | The editor is connected, loaded and has a session. No edit rights needed.                                    |
| `replace`            | The workbook is editable (`workflows.find.canReplace()`).                                                    |
| `applyFormatCells`   | Formatting may be applied now: editable and sheet protection allows it (`workflows.formatCells.canApply()`). |
| `insertFunctionCall` | The workbook is editable (`workflows.insertFunction.canInsert()`).                                           |
| `formatCells`        | `onOpenFormatCells` is installed and the workbook is editable.                                               |
| `findReplace`        | `onOpenFindReplace` is installed.                                                                            |
| `insertFunction`     | `onOpenInsertFunction` is installed and the workbook is editable.                                            |

The flags are computed for each `toolbarProps` snapshot. Each operation checks its state again when it runs.

### Example {#host-menus-and-workflows-example}

```tsx
import { useViewerBuildingBlocks } from '@christophervr/xlsx-react-viewer';

const { toolbarProps } = useViewerBuildingBlocks({
	workbook,
	canEdit: true,
	onContextMenu(request) {
		if (request.menu !== 'cell') return; // headers and tabs keep the native menu
		showMyMenu(request.x, request.y, request.items, (id) => request.execute(id));
		return true; // suppress the native menu
	},
	onOpenFindReplace: ({ findAll, select }) => showFindPanel({ findAll, select }),
});

// Later, from a panel of your own (toolbarProps is null until the canvas mounts):
toolbarProps?.workflows.formatCells.apply({ font: { bold: true } });
```

## Shadow parts

The element exposes its regions as CSS shadow parts, for styling with `xlsx-editor::part(grid)` and for tests: `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs`, `status-bar`, `title-bar`, `backstage` and `page-layout` (Page Layout view).
