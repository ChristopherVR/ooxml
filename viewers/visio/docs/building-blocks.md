# Compose a diagram viewer

`visio-react-viewer` exports `useViewerBuildingBlocks`, `DiagramCanvas`, and optional
`Toolbar`, `PageNavigation` and `Inspector`. These reuse ViewerController and the existing
SVG renderer and selection inspector.

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
	});
	return (
		<>
			<PageNavigation {...toolbarProps} />
			<DiagramCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

`content` accepts the same bytes/Blob source as the viewer. `handle` receives the existing
ViewerController; use `load`, `applyEdits`, `exportVsdx` and its reactive `state.edit.dirty`
for host persistence. `onDocumentChange`, `onDirtyChange` and `onDocumentError` are optional.
Give the canvas container a height and keep it mounted while loading.

The flat `toolbarProps` exposes page, zoom, selection, effective edit availability, undo/redo,
fit, search and the supported duplicate/group/ungroup/plain-text editing callbacks.
`Inspector` consumes `inspectorProps`. The canvas itself mounts no title bar, ribbon, page
controls, stencil/shape panes, inspector, file UI, recovery, settings or sharing dialogs.
Hosts own those optional workflows. Existing VSDX editing/protection limits still apply.

`canEdit={false}` is also available on the full viewer in every framework. It rejects local
mutation methods and editing gestures, including pending edits when permission is revoked.
Page navigation, selection, search, scrolling/pan, zoom, fit and serialization remain available.
Source-format editing support is separate from host permission.

Every framework package exports `mountDiagramCanvas(host, controller)`,
`diagramToolbarProps(controller, fit, locale?)` and `mountDiagramInspector(host, controller)`.
Subscribe through the existing controller, and call returned disposers on unmount. No second
state store is introduced.

The composition API requires the matching `ooxml-ui` runtime, released with these additions.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.

## Native framework composition

Framework-specific composition lives in each framework package. All adapters own one
`ViewerController`, expose its state through framework reactivity, cancel replaced loads,
and dispose subscriptions and canvas handlers with their owner. An independently removed
canvas can remount against that same controller without reloading the source.

### Vue

```vue
<script setup lang="ts">
import {
	DiagramCanvas,
	Toolbar,
	PageNavigation,
	Inspector,
	useViewerBuildingBlocks,
} from 'visio-vue-viewer';
const props = defineProps<{ content?: Uint8Array; canEdit?: boolean }>();
const blocks = useViewerBuildingBlocks(() => ({ content: props.content, canEdit: props.canEdit }));
const { toolbarProps, loading, error, mode } = blocks;
defineExpose({ controller: blocks.controller });
</script>
<DiagramCanvas v-bind="blocks.canvasProps" />
<Toolbar :toolbar-props="toolbarProps" />
<PageNavigation :toolbar-props="toolbarProps" />
<Inspector v-bind="blocks.inspectorProps" />
```

`toolbarProps`, `state`, `loading`, `error` and `mode` are refs; Vue templates unwrap them.
Pass an options getter or reactive ref to track source, permission and locale changes.

### Solid

```tsx
import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-solid-viewer';
function EmbeddedDiagram(props: { content?: Uint8Array; canEdit?: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		content: props.content,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<DiagramCanvas {...blocks.canvasProps} />
			<Toolbar {...blocks.toolbarProps()} />
		</>
	);
}
```

`toolbarProps`, `state`, `loading`, `error` and `mode` are accessors. The controller is
available as `blocks.controller`; disposal follows the Solid owner. `PageNavigation`
accepts the same flat toolbar descriptor and `Inspector` accepts `inspectorProps`.

### Svelte 5

```svelte
<script lang="ts">
 import { writable } from 'svelte/store';
 import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-svelte-viewer';
 const options = writable({ canEdit: true });
 const blocks = createViewerBuildingBlocks(options);
 const toolbarProps = blocks.toolbarProps;
 // options.set({ canEdit: false }) updates permissions through the store subscription.
</script>
<DiagramCanvas {...blocks.canvasProps} />
<Toolbar toolbarProps={$toolbarProps} />
```

The options argument accepts a readable store or a plain initial object; plain-object hosts
can call `blocks.update(options)` from a rune effect. `toolbarProps`, `state`, `loading`,
`error` and `mode` are readable stores. `PageNavigation` takes `toolbarProps` and `Inspector`
takes `inspectorProps`. The controller is available as `blocks.controller`.

### Angular

```ts
import { Component, Input, inject, type OnChanges } from '@angular/core';
import {
	ViewerBuildingBlocks,
	DiagramCanvasComponent,
	ToolbarComponent,
} from 'visio-angular-viewer';
@Component({
	selector: 'embedded-diagram',
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [DiagramCanvasComponent, ToolbarComponent],
	template: '<visio-diagram-canvas /><visio-diagram-toolbar />',
})
export class EmbeddedDiagram implements OnChanges {
	@Input() content?: Uint8Array;
	@Input() canEdit = true;
	readonly blocks = inject(ViewerBuildingBlocks);
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit });
	}
}
```

The scoped service exposes readonly `state`, `toolbarProps`, `loading`, `error` and `mode`
signals, the existing controller, and `documentChange`, `dirtyChange` and `documentError`
event emitters. `DiagramCanvasComponent`, `ToolbarComponent`, `PageNavigationComponent`
and `InspectorComponent` inject that owner. Place the provider on each editor host, rather
than application-wide, to isolate multiple documents. Alternatively wrap controls in
`<visio-building-blocks [content]="bytes" [canEdit]="allowed" (dirtyChange)="saveNeeded = $event">`
using `ViewerBuildingBlocksComponent`, which supplies inputs, outputs and a scoped provider.

All optional controls render native framework templates. Hosts may omit them and call the
flat toolbar actions from their own UI. Toolbar callbacks retain the controller's permission,
loading and source-format checks. Browser regression tests cover all four native compositions,
including editing, undo/redo, permission revocation, serialization/reload, error recovery and teardown.

For a host already using Angular signals, `createViewerBuildingBlocks(() => ({
content: bytes(), canEdit: allowed() }))` injects the scoped `ViewerBuildingBlocks` service
and connects the options getter with an Angular effect. Call it in the component's injection
context. The service's event emitters are active even before the first explicit `update`.

Changing a standalone canvas's `mount` prop releases its previous owner before attaching
the next one to the same host. Changing an inspector's controller similarly replaces its
subscription. This supports switching between independent editor sessions without orphaned
input handlers or state subscriptions.
