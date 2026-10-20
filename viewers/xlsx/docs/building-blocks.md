# Compose a spreadsheet editor

`@christophervr/xlsx-react-viewer` exports `useViewerBuildingBlocks`, `SpreadsheetCanvas`
and optional `Toolbar`, `FormulaBar` and `SheetTabs` components. All use the existing workbook
session, grid, formulas, validation, protection and history.

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

Render the canvas while loading; `toolbarProps` is null until mount. Give its container a
height. The canvas mounts only the grid. File, settings, sharing, title bar, ribbon, status
bar, formula bar and sheet tabs are absent unless the host supplies separate controls.

`toolbarProps` contains selection/active cell, sheet state, formula text, formatting and
`enabled` states. Connect `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange`,
`onFontSizeChange`, `onNumberFormatChange`, `onAlignmentChange`, `onUndo` and `onRedo`
to custom controls. Formatting fields are null for mixed values. Formatting inspection is
bounded to 4096 cells: larger selections return null and `formattingComplete=false`.

`onSelect`, `onSheetChange`, `onFormulaChange` and `commitEdit` return false when cell validation
blocks the action. Formatting callbacks return a Promise<boolean>. They commit an active
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

The composition API requires the matching `ooxml-ui` runtime, released with these additions.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.

## Native framework composition

Framework code is shipped from each framework's package. All adapters own one existing grid
engine and expose the same `EditorHandle`; their native reactive snapshots expose
`toolbarProps`, `loading`, `error` and `mode`. Optional native controls receive a
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
  let { workbook, canEdit = true }: {workbook: Workbook; canEdit?: boolean} = $props();
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
