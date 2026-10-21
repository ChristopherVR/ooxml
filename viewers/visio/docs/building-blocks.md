# Compose a diagram viewer

`visio-react-viewer` exports `useViewerBuildingBlocks`, `DiagramCanvas`, and optional
`Toolbar`, `PageNavigation`, `Inspector` and `ShapesPane`. These reuse ViewerController, the
existing SVG renderer, the selection inspector and the stencil window. Every other framework
has the same pieces (see below), including vanilla JavaScript.

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error, markClean } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
		onOpenSettings: () => openMySettingsDialog(),
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

The flat `toolbarProps` is the whole contract a custom UI needs; every field comes from the
existing controller, so nothing is stored twice:

| Field                                                                                                                               | Meaning                                                                                                                                                                                                                                                            |
| ----------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `mode`, `canEdit`                                                                                                                   | `mode` is `'edit'` only when the host granted `canEdit` **and** the open drawing is a source-backed `.vsdx` (a legacy `.vsd` or a model-only drawing reports `'view'`). `canEdit` is the same permission, further false while loading or while an edit is running. |
| `pageIndex`, `pages`, `zoom`, `selection`                                                                                           | Current page, page list, zoom factor and selected shapes.                                                                                                                                                                                                          |
| `canUndo`, `canRedo`, `canDuplicate`, `canGroup`, `canUngroup`, `canDelete`, `canCut`, `canCopy`, `canPaste`                        | Enabled flags. Edit actions are false while read-only; `canCopy` only needs a selection.                                                                                                                                                                           |
| `canZoomIn`, `canZoomOut`                                                                                                           | Zoom limits (10% to 800%).                                                                                                                                                                                                                                         |
| `search`                                                                                                                            | `{ query, resultCount, activeIndex, truncated }` from the controller's search. `activeIndex` is -1 until the host navigates.                                                                                                                                       |
| `tool`, `tools`, `canDraw`                                                                                                          | Active canvas tool (`'pointer'` or one of `rectangle`, `ellipse`, `line`, `freeform`, `arc`, `pencil`), the list with localized labels and active flags, and whether drawing is allowed.                                                                           |
| `labels`                                                                                                                            | Typed, localized labels (`DiagramToolbarLabels`), one entry for every control, in `en`, `de`, `es`, `fr` and `zh-CN`.                                                                                                                                              |
| `onPageChange`, `onZoomChange`, `onZoomIn`, `onZoomOut`, `onFit`                                                                    | Viewing actions, always available.                                                                                                                                                                                                                                 |
| `onUndo`, `onRedo`, `onDuplicate`, `onGroup`, `onUngroup`, `onDelete`, `onCut`, `onCopy`, `onPaste`, `onToolChange`, `onTextChange` | Editing actions. They keep the controller's permission, loading and source-format checks. Clipboard actions use the mounted canvas and the browser clipboard, so they need a user gesture.                                                                         |
| `onSearchChange`, `onNextSearchResult`, `onPreviousSearchResult`                                                                    | Search, always available.                                                                                                                                                                                                                                          |

`Inspector` consumes `inspectorProps`, and `ShapesPane` consumes `shapesProps`. The canvas
itself mounts no title bar, ribbon, page controls, shape/stencil pane, inspector, file UI,
recovery, settings or sharing dialogs. Existing VSDX editing/protection limits still apply.

## Permission and mode

`canEdit` follows the PowerPoint building blocks: the drawing is **read-only until the host
passes `canEdit: true`**, and the permission is applied when the controller is created, before the
canvas mounts, so nothing is briefly editable (the full viewer keeps its editable default).
Revoking it cancels pending edits, drops the draw tool and removes the resize, rotate and
end-point handles. Page navigation, selection, search, scrolling/pan, zoom, fit, copy and
serialization remain available. Source-format editing support is separate from host permission,
which is why `mode` also checks it.

## What the canvas can do

The standalone canvas uses the full viewer's own gestures (one shared wiring, not a copy): move,
inline text (double-click, F2 or typing on a selected shape), resize and rotate handles, line end
points, and the draw tools picked with `onToolChange`. Delete, Ctrl+Z, Ctrl+Y and Escape (leave
the draw tool) work on the focused canvas, and the browser's copy/cut/paste events are honoured.
Every one of them is gated by `canEdit`, exactly like the full viewer. Edit refusals are announced
in a polite live region inside the canvas.

## Optional library UI

`Toolbar` renders every control `toolbarProps` supports: history, zoom and fit, duplicate, group,
ungroup and delete, cut/copy/paste, the drawing tools, search and the host dialog buttons below,
with enabled state. It is optional: pass `include` (`'file' | 'history' | 'zoom' | 'selection' |
'clipboard' | 'tools' | 'search'`) to show only some groups, or skip it and wire `toolbarProps` to your
own controls. `PageNavigation` is the page selector.

`ShapesPane` is the Shapes (stencil) window: Search shapes, the drawing's Document Stencil and the
built-in stencils. Clicking a master, or dropping it on the mounted canvas, goes through the
same creation flow as the full viewer, so it obeys the same permission. Mount the canvas on the
same controller; without a canvas, masters are added at the page centre.

## Host-owned dialogs

As in PowerPoint's `onOpen...` callbacks, dialogs belong to the host and the library never mounts
one in surface mode. Pass any of `onOpenFile`, `onOpenPageSetup`, `onOpenSettings` and
`onOpenShareDialog` to the hook or factory. They come back on `toolbarProps` with
`canOpenFile`, `canOpenPageSetup`, `canOpenSettings` and `canOpenShareDialog`; a callback you
omit is absent and its flag is `false` (not enabled), and the library `Toolbar` shows a button
only for the ones you supplied.

## Persistence and dirty tracking

`controller.exportVsdx()` returns the bytes and never changes `dirty`. After the host has persisted
them it calls `markClean()` (on the hook result, the controller or the handle). That clears
`state.edit.dirty`, fires `onDirtyChange(false)` and a `document-change` event of kind `save`.
Later edits and undo steps are measured from that saved point.

## Packages

Every framework package re-exports `ooxml-ui/visio`, so `mountDiagramCanvas(host, controller)`,
`mountDiagramInspector(host, controller)`, `mountDiagramShapes(host, controller, { locale? })`,
`diagramToolbarProps(controller, fit, locale?, hostCallbacks?)` and `fitDiagramCanvas` are available
from any of them; they need an `ooxml-ui` runtime that contains them. Subscribe through the existing
controller and call the returned disposers on unmount. No second state store is introduced.

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
accepts the same flat toolbar descriptor, `Inspector` accepts `inspectorProps` and
`ShapesPane` accepts `shapesProps`.

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
`error` and `mode` are readable stores. `PageNavigation` takes `toolbarProps`, `Inspector`
takes `inspectorProps` and `ShapesPane` takes `shapesProps`. The controller is available as `blocks.controller`.

### Vanilla JavaScript

```ts
import { createViewerBuildingBlocks } from 'visio-vanilla-viewer';
const blocks = createViewerBuildingBlocks({ canEdit: false, onOpenSettings: openSettings });
blocks.mountCanvas(document.querySelector('#canvas')!);
blocks.mountToolbar(document.querySelector('#toolbar')!, ['history', 'zoom']);
blocks.mountPageNavigation(document.querySelector('select')!);
blocks.mountInspector(document.querySelector('#inspector')!);
blocks.mountShapes(document.querySelector('#shapes')!);
blocks.update({ content: bytes, canEdit: true });
blocks.subscribe(() => render(blocks.toolbarProps, blocks.mode));
// later: blocks.markClean(); blocks.destroy();
```

The vanilla factory runs the same session as the other adapters. `toolbarProps`, `state`,
`loading`, `error` and `mode` are getters, `subscribe(listener)` reports every state or option
change, and each `mount...` helper returns its disposer; `destroy()` disposes whatever is
still mounted. The mounted toolbar updates its elements in place, so focus survives refreshes.

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
	@Input() canEdit = false;
	readonly blocks = inject(ViewerBuildingBlocks);
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit });
	}
}
```

The scoped service exposes readonly `state`, `toolbarProps`, `loading`, `error` and `mode`
signals, the existing controller, and `documentChange`, `dirtyChange` and `documentError`
event emitters. `DiagramCanvasComponent`, `ToolbarComponent`, `PageNavigationComponent`,
`InspectorComponent` and `ShapesPaneComponent` inject that owner. Place the provider on each editor host, rather
than application-wide, to isolate multiple documents. Alternatively wrap controls in
`<visio-building-blocks [content]="bytes" [canEdit]="allowed" (dirtyChange)="saveNeeded = $event">`
using `ViewerBuildingBlocksComponent`, which supplies inputs, outputs and a scoped provider.

All optional controls render native framework templates. Hosts may omit them and call the
flat toolbar actions from their own UI. Toolbar callbacks retain the controller's permission,
loading and source-format checks. Browser regression tests cover all six native compositions,
including editing, undo/redo, permission revocation, serialization/reload, error recovery and teardown.

For a host already using Angular signals, `createViewerBuildingBlocks(() => ({
content: bytes(), canEdit: allowed() }))` injects the scoped `ViewerBuildingBlocks` service
and connects the options getter with an Angular effect. Call it in the component's injection
context. The service's event emitters are active even before the first explicit `update`.

Changing a standalone canvas's `mount` prop releases its previous owner before attaching
the next one to the same host. Changing an inspector's controller similarly replaces its
subscription. This supports switching between independent editor sessions without orphaned
input handlers or state subscriptions.

## Shared Office theme

Every native optional toolbar, page selector and inspector follows the same `--office-*`
CSS variables as the drawing canvas. The blocks install their shared stylesheet automatically;
there is no framework-specific theme object or second document state store. Scope themes on
your own host element:

```html
<section
	data-office-theme="dark"
	style="--office-accent: #a78bfa; --office-font: Inter, sans-serif"
>
	<!-- DiagramCanvas, Toolbar, PageNavigation and Inspector -->
</section>
```

Use `data-office-theme="light"`, `"dark"` or `"auto"` on the document root or a containing
section. `auto` follows the system preference, including runtime changes. Scoped light/dark
hosts can coexist on one page. With no scoped attribute, blocks inherit the root theme and
host variables across the canvas's shadow boundary. Changing the theme does not reload the
canvas, reset selection or modify document colors or serialized bytes.

For a host-specific palette, override `--office-background`, `--office-surface`,
`--office-foreground`, `--office-muted-foreground`, `--office-border`, `--office-accent`,
`--office-ring` and `--office-font` on the containing section. Legacy Visio overrides
`--vv-surface`, `--vv-ink`, `--vv-border`, `--vv-accent`, `--vv-focus` and `--vv-font`
remain supported, with precedence over their canonical counterparts. The canvas additionally
keeps `--vv-background`, `--vv-secondary` and the existing status-bar overrides. The diagram
page retains its own document appearance while the surrounding controls follow the UI theme.
