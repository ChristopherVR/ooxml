# Building blocks

Every framework package (`@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer`, `xlsx-angular-viewer`,
`xlsx-svelte-viewer`, `xlsx-solid-viewer` and `xlsx-vanilla-viewer`) exports a building-block
hook or factory (`useViewerBuildingBlocks`, `createViewerBuildingBlocks`), a `SpreadsheetCanvas`
and optional `Toolbar`, `FormulaBar` and `SheetTabs` controls. All use the existing workbook
session, grid, formulas, validation, protection and history; there is no second state store.
The shape follows the PowerPoint viewer's `useViewerBuildingBlocks`.

```tsx
import {
	SpreadsheetCanvas,
	FormulaBar,
	SheetTabs,
	useViewerBuildingBlocks,
} from '@christophervr/xlsx-react-viewer';

function EmbeddedSheet({ workbook, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		workbook,
		canEdit,
		handle,
	});
	return (
		<>
			{toolbarProps && <FormulaBar {...toolbarProps} />}
			<SpreadsheetCanvas {...canvasProps} />
			{toolbarProps && <SheetTabs {...toolbarProps} />}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

The hook also returns `mode` (`'view'` or `'edit'`) and reactive `dirty`, `selection` (A1 text),
`collaboration` (who is in the room and the connection state) and `editError` (see
[Validation failures](#validation-failures)). Render the canvas while loading; `toolbarProps`
is null until mount. Give its container a
height. The canvas mounts only the grid. File, settings, sharing, title bar, ribbon, status
bar, formula bar and sheet tabs are absent unless the host supplies separate controls.

`toolbarProps` contains selection/active cell, sheet state, formula text, formatting and
`enabled` states. Connect `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange`,
`onFontSizeChange`, `onNumberFormatChange`, `onAlignmentChange`, `onUndo` and `onRedo`
to custom controls. Formatting fields are null for mixed values. Formatting inspection is
bounded to 4096 cells: larger selections return null and `formattingComplete=false`.

`enabled` has one flag per callback: `bold`, `italic`, `underline`, `fontFamily`, `fontSize`,
`alignment`, `numberFormat`, `undo`, `redo` and the host workflows `share`, `settings`,
`formatCells`, `findReplace` and `insertFunction`. `canUndo` and `canRedo` repeat the undo and
redo flags. `options` lists the choices a custom font, size, alignment or number format control
can offer (`fontFamilies`, `fontSizes`, `alignments`, `numberFormats`), and `labels` holds the
localized control names.

`onSelect`, `onSheetChange`, `onFormulaChange` and `commitEdit` return false when cell validation
blocks the action. Formatting callbacks return a `Promise<boolean>`. They commit an active
cell edit before acting, and retain its selection when custom controls take focus. Formula
text is committed on Enter by the optional FormulaBar; a rejected entry sets aria-invalid.
The existing handle retains save/serialization and dirty tracking. A save rejects if input
cannot be committed. Permissions and worksheet protection remain authoritative.

For every framework, `mountEditorSurface(host, options)` exposes the same surface, existing
handle, `getToolbarProps()`, `subscribe(listener)`, `update` and `destroy`. Hosts own file
and optional dialog workflows; no hidden library toolbar is required.

`formula`, `formulaCaret`, `editing` and `editMode` read the grid's live edit bridge.
`onFormulaInput(text, caret?)` updates that same draft; `onFormulaChange(text)` commits it.
A rejected draft can be corrected through `onFormulaInput` without first committing the invalid
value. `onCancelEdit()` cancels it. There is no separate formula state store.

The composition API requires the `ooxml-ui` runtime of the same workspace release; the framework packages depend on it and are released with it.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.

## Host callbacks for dialogs

The application opens dialogs for sharing, settings, Format Cells, Find and Replace and Insert
Function. A surface never mounts them. Pass a callback instead, as PowerPoint's building
blocks do (`onOpenSettings`, `onOpenShareDialog`), and open your own UI:

```ts
useViewerBuildingBlocks({
	workbook,
	onOpenShareDialog: ({ activeCell }) => openMyShareDialog(activeCell),
	onOpenFormatCells: ({ selection }) => openMyFormatPanel(selection),
});
```

The available callbacks are `onOpenShareDialog`, `onOpenSettings`, `onOpenFormatCells`,
`onOpenFindReplace` and `onOpenInsertFunction`. Each receives `{ sheet, sheetName, selection,
activeCell }`. `toolbarProps` exposes the matching `onOpenShareDialog`, `onOpenSettings`,
`onOpenFormatCells`, `onOpenFindReplace` and `onOpenInsertFunction` (they commit an in-progress
edit first and return a `Promise<boolean>`), and `enabled.openShareDialog` and the like are true only when
the callback was given (the formatting and insert ones also need edit permission). `handle.share()`
calls `onOpenShareDialog`; without one it does nothing, and `enabled.openShareDialog` is false. A
workflow without a callback is unreachable, never silently mounted. Hosts that want the
bundled dialogs render the full `SpreadsheetEditor` instead.

## Validation failures

A cell edit in progress must be committed before formatting, navigation or saving. When data
validation refuses it, the callbacks return `false`, the edit stays open and the failure is
surfaced three ways:

- `onValidationError({ message, title?, ref, style? })` in the options (the element dispatches
  a `validation-error` event with the same detail);
- the reactive `editError` string (also `toolbarProps.editError`), null again once the draft is
  changed, corrected or canceled;
- the optional `FormulaBar` sets `aria-invalid` and renders the message in a `role="alert"` span.

### The data-validation alert

The building blocks never draw the grid's Retry / Cancel alert. Excel's three alert styles are
the host's to show through `onOpenValidationAlert(request)` (also accepted in `surfaceActions`
and by every framework adapter). The request carries `{ ref, sheet, sheetName, value, style,
title, message }`, where `value` is the text the user entered, and the methods that match
Excel's buttons:

| `style`       | Excel buttons     | Methods                                      |
| ------------- | ----------------- | -------------------------------------------- |
| `stop`        | Retry / Cancel    | `retry()`, `cancel()`                        |
| `warning`     | Yes / No / Cancel | `accept()` (Yes), `retry()` (No), `cancel()` |
| `information` | OK / Cancel       | `accept()` (OK), `cancel()`                  |

`retry()` keeps the editor open with the draft, `cancel()` restores the old value and `accept()`
commits the value anyway, exactly like the bundled buttons. Resolve a request once; later calls
are ignored. Until it is resolved the cell editor ignores further commits, so always resolve it
(a callback that throws counts as `retry()`). `onValidationError` and `editError` still report
the failure.

Without the callback nothing is mounted: every style acts like Retry, the edit stays open and
the failure is reported through `onValidationError` / `editError`.
`toolbarProps.enabled.validationAlert` is true only when the host supplied the callback. The
full `SpreadsheetEditor` and `<xlsx-editor>` keep their built-in alert.

Custom controls should also show `editError` and let the user correct the draft with
`onFormulaInput` or discard it with `onCancelEdit`.

## Optional Toolbar

The library `Toolbar` renders every control the props support, so it can replace a hand-built
one: undo, redo, font family, font size, bold, italic, underline, left / center / right
alignment and number format. Buttons expose `aria-pressed` (`mixed` for a mixed selection),
selects show a blank value for mixed formatting, and all controls honour `enabled`. Buttons do
not take focus, so the grid keeps its selection while you format. The Toolbar does not render
the dialog workflows above; connect `onOpenShareDialog` and the other `onOpen*` callbacks to your own
controls. `toolbarModel(toolbarProps)` (exported by `xlsx-vanilla-viewer` and the Svelte runtime) returns
the same control list for custom renderers.

## Collaboration

Real-time co-editing works on a surface: pass `collaboration` (a room id, server URL, provider or
session) in the options, and the hook returns the reactive `collaboration` state (`active`,
`roomId`, `status`, `synced`, `people`). `onUndo` and `onRedo` undo only this window's edits while
shared. Leaving or unmounting leaves the room. The title bar's presence list belongs to the
application, so a host renders `collaboration.people` itself.

## Vanilla

`xlsx-vanilla-viewer` exports `createViewerBuildingBlocks(options)`, built on the same session as
the other frameworks:

```ts
import { createViewerBuildingBlocks } from 'xlsx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ workbook, canEdit: true });
const stops = [
	blocks.mountToolbar(toolbarHost),
	blocks.mountFormulaBar(formulaHost),
	blocks.mountSheetTabs(tabsHost),
	blocks.mountCanvas(canvasHost),
];
blocks.subscribe((state) => render(state.dirty, state.editError, state.toolbarProps));
blocks.update({ workbook, canEdit: false });
blocks.getState().toolbarProps?.onBold();
blocks.destroy();
```

The surface is **read-only until the host passes `canEdit: true`**, matching the PowerPoint and
Visio building blocks (the full `SpreadsheetEditor` stays editable by default). The permission is
applied before the canvas mounts, so `mode` starts as `'view'` and there is no editable flash.
`readOnly` still wins over `canEdit: true`.

Every `mount*` returns its disposer. `getState()` returns the same snapshot the hooks do
(`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration`, `editError`) and
`handle` is the usual imperative handle. The controls patch their elements in place, so
focus and drafts survive updates.

## Native framework composition

Framework code is shipped from each framework's package. All adapters own one existing grid
engine and expose the same `EditorHandle`; their native reactive snapshots expose
`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration` and `editError`. Optional native controls receive a
`toolbarProps` prop/input. Custom controls call that same snapshot's actions.

### Vue

Call the composable during `setup`. Pass a getter to track permission, document, bytes and
callback changes. Computed refs are unwrapped in templates; unmount disposes the engine.

```vue
<script setup lang="ts">
import {
	useViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-vue-viewer';
const props = defineProps<{ workbook: import('xlsx-vue-viewer').Workbook; canEdit: boolean }>();
const { canvasProps, toolbarProps, loading, error, handle } = useViewerBuildingBlocks(() => ({
	workbook: props.workbook,
	canEdit: props.canEdit,
}));
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<FormulaBar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<SpreadsheetCanvas v-bind="canvasProps" style="height: 600px" />
	<SheetTabs v-if="toolbarProps" :toolbar-props="toolbarProps" />
</template>
```

### Angular

Provide `ViewerBuildingBlocks` per host, then inject that instance into the host or its
custom controls. `options` is a writable signal; snapshots are computed signals. The
injector's `DestroyRef` owns cleanup. The service also exposes `workbookChange`,
`workbookError`, `selectionChange` and `dirtyChange` event emitters for host subscriptions,
alongside the callbacks in options. Use `takeUntilDestroyed()` when subscribing in hosts.

```ts
@Component({
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [SpreadsheetCanvasComponent, ToolbarComponent, FormulaBarComponent, SheetTabsComponent],
	template: `
		@if (blocks.toolbarProps(); as p) {
			<spreadsheet-toolbar [toolbarProps]="p" />
			<spreadsheet-formula-bar [toolbarProps]="p" />
			<button (click)="p.onBold()" [disabled]="!p.canEdit">Custom bold</button>
			<spreadsheet-sheet-tabs [toolbarProps]="p" />
		}
		<spreadsheet-canvas [mount]="blocks.canvasProps.mount" style="display:block;height:600px" />
	`,
})
export class WorkbookHost {
	readonly workbook = input.required<Workbook>();
	readonly canEdit = input(true);
	readonly blocks = inject(ViewerBuildingBlocks);
	constructor() {
		effect(() => this.blocks.options.set({ workbook: this.workbook(), canEdit: this.canEdit() }));
	}
}
```

Import the components, service and `Workbook` from `xlsx-angular-viewer`, and Angular's
`Component`, `input`, `effect` and `inject` from `@angular/core`. `SpreadsheetBuildingBlocks`
is the product-specific name for the same service. Its `handle` retains load/save/selection.

### Solid

Create inside a Solid owner. Input getters and output accessors participate in Solid's
reactive graph; cleanup follows that owner. Components accept reactive getters as props.

```tsx
import { Show } from 'solid-js';
import {
	createViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-solid-viewer';
function WorkbookHost(props: { workbook: Workbook; canEdit: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		workbook: props.workbook,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<Show when={blocks.toolbarProps()}>
				{(p) => (
					<>
						<Toolbar toolbarProps={p()} />
						<FormulaBar toolbarProps={p()} />
						<button onClick={() => p().onBold()}>Custom bold</button>
						<SheetTabs toolbarProps={p()} />
					</>
				)}
			</Show>
			<SpreadsheetCanvas {...blocks.canvasProps} />
		</>
	);
}
```

### Svelte 5

Create the store during component initialization. `$blocks` subscribes natively; use an
`$effect` to forward reactive options. Engine cleanup follows component destruction.
The component sources and their declarations are published as explicit subpaths.

```svelte
<script lang="ts">
  import { createViewerBuildingBlocks, type Workbook } from 'xlsx-svelte-viewer/runtime';
  import SpreadsheetCanvas from 'xlsx-svelte-viewer/SpreadsheetCanvas.svelte';
  import Toolbar from 'xlsx-svelte-viewer/Toolbar.svelte';
  import FormulaBar from 'xlsx-svelte-viewer/FormulaBar.svelte';
  import SheetTabs from 'xlsx-svelte-viewer/SheetTabs.svelte';
  let { workbook, canEdit = false }: {workbook: Workbook; canEdit?: boolean} = $props();
  const blocks = createViewerBuildingBlocks();
  $effect(() => blocks.update({workbook, canEdit}));
</script>
{#if $blocks.toolbarProps}
  <Toolbar toolbarProps={$blocks.toolbarProps} />
  <FormulaBar toolbarProps={$blocks.toolbarProps} />
  <button onclick={() => $blocks.toolbarProps?.onBold()}>Custom bold</button>
  <SheetTabs toolbarProps={$blocks.toolbarProps} />
{/if}
<SpreadsheetCanvas {...blocks.canvasProps} />
```

`FormulaBar` shares the grid's live draft. Enter commits and sets `aria-invalid` on a
rejected entry. Retry allows correction; Escape cancels. Conflicting formatting, selection,
sheet navigation and saving remain blocked while validation rejects a draft. Readonly
blocks mutation while preserving selection, sheet navigation and serialization.

## Shared appearance

Place the canvas and optional native controls under one `data-office-theme="light|dark|auto"`
ancestor. All five native framework integrations install the same shared control styles and
read inherited `--office-*` tokens. Customise `--office-surface`, `--office-background`,
`--office-foreground`, `--office-border`, `--office-accent`, `--office-ring`, `--office-font`
and `--office-radius` at that scope. Runtime changes preserve drafts and focus.

The canvas's existing `theme` and `themeColors` options continue to apply locally; they do
not rewrite the host's ancestors or sibling controls. Legacy `--xve-*` overrides on a common
ancestor still reach both the canvas and optional controls. See [Theming](./theming) for
scope and compatibility examples.
