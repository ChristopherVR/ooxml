# Compose a Word editor

`docx-react-viewer` exports `useViewerBuildingBlocks`, `DocumentCanvas` and an optional
`Toolbar`. They use WordEditor's existing editor, layout, history and serialization engine.

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		documentModel,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
	});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Render the canvas even while loading. `toolbarProps` is null until it mounts. It exposes
selection, bold/italic/underline (null for mixed values), font family/size, alignment,
list state, undo/redo availability and direct `on*` callbacks. `Toolbar` consumes the same
props when library controls are wanted. Give the canvas container a height.

Use `bytes` to load a file, or the existing handle's `load`, `save`, `download`, `markClean`
and `dirty`. Saving does not mark a document clean until the host confirms persistence.
`canEdit={false}` or `readOnly` blocks editing while selection and viewing remain available.
The canvas mounts no title bar, ribbon, backstage, status bar or application dialogs.
Hosts own optional dialogs and file/settings/sharing flows.

Every framework package also exports `mountEditorSurface(host, options)`. Its binding exposes
`getToolbarProps()` and `subscribe(listener)`, plus the existing editor handle and
`update`/`destroy`. This is the framework-neutral equivalent for native lifecycle adapters:
subscribe to selection/format changes, wire controls directly, and destroy on unmount.

The composition API requires the matching `ooxml-ui` runtime, released with these additions.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.

## Native framework composition

Framework adapters live in their own framework package. They share only the engine lifetime
helper. Their state uses the framework's normal reactive primitives and their canvas components
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

## Shared styling

Optional native toolbar components use the shared `ooxml-ui` styles, including spacing,
backgrounds, borders, pressed/disabled states and keyboard focus rings. Place the toolbar and
canvas beneath one `data-office-theme` ancestor and customize `--office-*` variables there.
Changes apply immediately to every framework. Existing ancestor `--dve-*` overrides remain
supported. Host-created controls keep their own styling unless the host opts into the shared
`prepareOfficeUiBlock(element, 'toolbar', 'docx')` helper from `ooxml-ui/theme`.
