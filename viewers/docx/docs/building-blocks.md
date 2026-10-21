# Building blocks

`docx-react-viewer` exports `useViewerBuildingBlocks`, `DocumentCanvas` and optional
`Toolbar`, `PageNavigation` and `StatusBar` components. They use WordEditor's existing editor,
layout, history and serialization engine, and follow the shape of the PowerPoint
`useViewerBuildingBlocks`: the hook returns props, the host renders what it wants.

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, navigationProps, statusProps, loading, error } =
		useViewerBuildingBlocks({
			documentModel,
			canEdit,
			handle,
			onDirtyChange: (dirty) => console.log(dirty),
			// Host-owned dialogs: omit one and its action reports as not enabled.
			onOpenHyperlinkDialog: (request) => openMyLinkDialog(request),
			onOpenFindReplace: (request) => openMyFindBar(request),
			onOpenImagePicker: (request) => openMyFilePicker(request),
		});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{statusProps && <span>{statusProps.wordText}</span>}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Render the canvas even while loading. The props are null until it mounts. Give the canvas
container a height.

## What the hook returns

| Field                      | Contents                                                           |
| -------------------------- | ------------------------------------------------------------------ |
| `canvasProps`              | Spread onto `DocumentCanvas`. Mounts and releases the editor.      |
| `toolbarProps`             | Selection and formatting state plus editing callbacks (below).     |
| `navigationProps`          | Page, page count, view mode and page navigation callbacks.         |
| `statusProps`              | Word count and zoom with callbacks.                                |
| `loading`, `error`, `mode` | Load state, the last load or document error, `'view'` or `'edit'`. |

### toolbarProps

State: `selection` (`{ from, to }`), `bold`, `italic`, `underline` (`null` when the selection
mixes values), `fontFamily`, `fontSize` (`null` when mixed), `alignment`
(`'left' | 'center' | 'right' | 'justify'` or `null`), `list`
(`'bullet' | 'number' | 'other' | null`), `canEdit`, `canUndo`, `canRedo` and translated `labels`.

Callbacks, each returning `false` when the action is not allowed (read-only or still loading):
`onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange(family)`, `onFontSizeChange(points)`,
`onAlignmentChange(alignment)`, `onListChange('bullet' | 'number' | 'remove')`, `onUndo`, `onRedo`.
The `list` state and `onListChange` use the same `'bullet' | 'number'` values.

### navigationProps

`page`, `pageCount`, `pageText` ("Page 2 of 5"), `pagesAvailable`, `canGoPrevious`, `canGoNext`,
`viewMode` (`'draft' | 'print'`) and the callbacks `onGoToPage(n)`, `onPreviousPage`,
`onNextPage` and `onViewModeChange(mode)`. Pages exist in Print Layout only: in Draft,
`pagesAvailable` is false and the page callbacks return `false`. Switch with
`onViewModeChange('print')`.

### statusProps

`words`, `wordText` ("42 words"), `zoom` (percent), `minZoom`, `maxZoom` and the callbacks
`onZoomChange(percent)`, `onZoomIn` and `onZoomOut`. Scrolling updates `navigationProps` as well.

## Host dialogs

A surface mounts no dialog of its own. These workflows are the host's, named like the
PowerPoint `onOpen...` callbacks. Pass them as options; `toolbarProps` then carries a matching
`open...` opener (`openHyperlinkDialog`, `openFindReplace`, `openImagePicker`), a `can...` flag and `hostWorkflows` (which callbacks were supplied).

| Option                  | Request object                                                                   | Enabled while                 |
| ----------------------- | -------------------------------------------------------------------------------- | ----------------------------- |
| `onOpenHyperlinkDialog` | `{ href, text, apply(href, text?), remove() }` for the link at the selection     | editable                      |
| `onOpenFindReplace`     | `{ search, findNext, findPrevious, replaceOne, replaceAll, status, canReplace }` | loaded (read-only keeps find) |
| `onOpenImagePicker`     | `{ insert(file) }`: the host picks a file and hands it back                      | editable                      |

With a callback omitted, `canInsertLink`, `canFindReplace` or `canInsertPicture` is false and the
opener returns `false`. Requests re-check permission when used, so a dialog left open across a
switch to read-only cannot edit.

## Library components

`Toolbar` draws undo and redo, font family and size, bold, italic, underline, four alignments,
bulleted and numbered lists, and a button for each host workflow you supplied. `PageNavigation`
draws previous page, "Page n of m", next page and the Draft or Print Layout switch.
`StatusBar` draws the word count and a zoom select. All are optional and props-driven; each
control has a stable node, so keyboard focus survives formatting updates. They use the shared
Office styles. Custom controls never need them.

```tsx
{
	toolbarProps && <Toolbar {...toolbarProps} />;
}
{
	navigationProps && <PageNavigation {...navigationProps} />;
}
{
	statusProps && <StatusBar {...statusProps} />;
}
```

## Read-only surfaces

`canEdit={false}` (or `readOnly`) is authoritative: every editing callback returns `false` and
the document cannot change. The reader can still focus the surface, select text with the mouse or
keyboard (arrows, Shift+arrows, Home, End, Ctrl+A), copy, and read the selection from
`toolbarProps.selection`. Keyboard shortcuts that belong to the application (find, save, print)
are not installed on a surface; wire them to your own controls.

## Lifecycle

Use `bytes` to load a file, or the handle's `load`, `save`, `download`, `markClean` and
`dirty`. Saving does not mark a document clean until the host confirms persistence.
The canvas mounts no title bar, ribbon, backstage, status bar or application dialogs.

Every framework package also exports `mountEditorSurface(host, options)` for native lifecycle
adapters. Its binding exposes `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` and
`subscribe(listener)`, plus the editor handle, `update` and `destroy`. Surface options exclude the
ribbon-only options (`showToolbar`, `showThumbnails`, `hiddenActions`, `ribbonAddIns`,
`onRibbonAddIn`), so a TypeScript host cannot pass one that would be ignored.

The composition API requires the matching `ooxml-ui` runtime, released with these additions.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.

## Native framework composition

Framework adapters live in their own framework package. They share the engine lifetime
helper and the control descriptors (`bindings/src/controls.ts`) the optional components draw.
Every adapter exposes `toolbarProps`, `navigationProps`, `statusProps`, `loading`, `error`,
`mode` and `canvasProps`, plus the editor `handle` (React takes a `handle` ref as an option instead of returning one), and the components `Toolbar`, `PageNavigation` and
`StatusBar` (Angular: `ToolbarComponent`, `PageNavigationComponent`, `StatusBarComponent`; Svelte:
`docx-svelte-viewer/Toolbar`, `/PageNavigation`, `/StatusBar`). Their state uses the framework's normal reactive primitives and their canvas components
mount the existing Word document surface.

### Vue

```vue
<script setup lang="ts">
import { computed } from 'vue';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-vue-viewer';
const props = defineProps<{ documentModel: import('docx-core').DocumentModel; canEdit: boolean }>();
const { canvasProps, toolbarProps, error, handle } = useViewerBuildingBlocks(
	computed(() => ({ documentModel: props.documentModel, canEdit: props.canEdit })),
);
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<DocumentCanvas v-bind="canvasProps" class="document-surface" />
	<p v-if="error" role="alert">{{ error.message }}</p>
</template>
```

The composable watches a ref, computed ref or getter and disposes the engine with its scope.
`toolbarProps`, `loading`, `error` and `mode` are readonly computed refs. `handle` exposes the
existing imperative editor API.

### Angular

Provide one `ViewerBuildingBlocks` service per host component. Signals read in the template
participate in Angular change detection, including zoneless applications. `DestroyRef` releases
the engine when its provider is destroyed.

```ts
import { Component, signal } from '@angular/core';
import { createDocument } from 'docx-core';
import {
	createViewerBuildingBlocks,
	ViewerBuildingBlocks,
	DocumentCanvasComponent,
	ToolbarComponent,
} from 'docx-angular-viewer';

@Component({
	selector: 'embedded-word',
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent],
	providers: [ViewerBuildingBlocks],
	template: `
		@if (blocks.toolbarProps(); as toolbar) {
			<word-toolbar [toolbarProps]="toolbar" />
		}
		<word-document-canvas [canvasProps]="blocks.canvasProps" />
	`,
})
export class EmbeddedWord {
	readonly canEdit = signal(true);
	readonly model = createDocument();
	readonly blocks = createViewerBuildingBlocks(() => ({
		documentModel: this.model,
		canEdit: this.canEdit(),
	}));
}
```

The service exposes readonly `toolbarProps`, `loading`, `error` and `mode` signals, a stable
`handle`, and typed `documentChange`, `documentError`, `dirtyChange` and `pageChange`
event emitters. Subscribe to those emitters or supply the existing option callbacks. For
imperative hosts, inject the service and call `update(options)` instead of the helper.

### Svelte 5

```svelte
<script lang="ts">
  import { writable } from 'svelte/store';
  import { createDocument } from 'docx-core';
  import { createViewerBuildingBlocks } from 'docx-svelte-viewer/runtime';
  import DocumentCanvas from 'docx-svelte-viewer/DocumentCanvas';
  import Toolbar from 'docx-svelte-viewer/Toolbar';
  const options = writable({ documentModel: createDocument(), canEdit: true });
  const blocks = createViewerBuildingBlocks(options);
  const { toolbarProps, error } = blocks;
</script>
{#if $toolbarProps}<Toolbar toolbarProps={$toolbarProps} />{/if}
<DocumentCanvas {...blocks.canvasProps} />
{#if $error}<p role="alert">{$error.message}</p>{/if}
```

Call the factory during component initialization. It accepts a readable options store or an
initial options object; `blocks.update(next)` also supports rune-driven `$effect` updates.
Snapshots are readable stores. `documentCanvas` is exported as an action for custom host
markup. Component destruction unsubscribes from options and destroys the engine.

### Solid

```tsx
import { createSignal, Show } from 'solid-js';
import { createDocument } from 'docx-core';
import { createViewerBuildingBlocks, DocumentCanvas, Toolbar } from 'docx-solid-viewer';
function EmbeddedWord() {
	const [canEdit, setCanEdit] = createSignal(true);
	const model = createDocument();
	const blocks = createViewerBuildingBlocks(() => ({ documentModel: model, canEdit: canEdit() }));
	return (
		<>
			<Show when={blocks.toolbarProps()}>{(toolbar) => <Toolbar toolbarProps={toolbar()} />}</Show>
			<DocumentCanvas {...blocks.canvasProps} />
		</>
	);
}
```

The factory tracks its options accessor in an effect and returns signal accessors for
`toolbarProps`, `loading`, `error` and `mode`. Its owner controls cleanup through `onCleanup`.
Optional toolbar controls keep their DOM nodes while formatting state changes, so keyboard
focus is preserved. All adapters accept `bytes` reactively, clear failed loads when a new
model or load arrives, and keep their handle stable across canvas replacement.

### Vanilla

`docx-vanilla-viewer` exports the shared session as `createViewerBuildingBlocks(options)`
(`canvasProps`, `handle`, `state`, `subscribe`, `update`, `destroy`) and mount helpers that
return a disposer: `mountDocumentCanvas(host, blocks)`, `mountToolbar(host, blocks)`,
`mountPageNavigation(host, blocks)` and `mountStatusBar(host, blocks)`.

```ts
import {
	createViewerBuildingBlocks,
	mountDocumentCanvas,
	mountToolbar,
	mountStatusBar,
} from 'docx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ documentModel, canEdit: true });
const stops = [
	mountDocumentCanvas(document.querySelector('#page')!, blocks),
	mountToolbar(document.querySelector('#tools')!, blocks),
	mountStatusBar(document.querySelector('#status')!, blocks),
];
blocks.subscribe((state) => console.log(state.toolbarProps?.bold)); // your own controls
// later
stops.forEach((stop) => stop());
blocks.destroy();
```

## Shared styling

Optional native toolbar components use the shared `ooxml-ui` styles, including spacing,
backgrounds, borders, pressed/disabled states and keyboard focus rings. Place the toolbar and
canvas beneath one `data-office-theme` ancestor and customize `--office-*` variables there.
Changes apply immediately to every framework. Existing ancestor `--dve-*` overrides remain
supported. Host-created controls keep their own styling unless the host opts into the shared
`prepareOfficeUiBlock(element, 'toolbar', 'docx')` helper from `ooxml-ui/theme`.
