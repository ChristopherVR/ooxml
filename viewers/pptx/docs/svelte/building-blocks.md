---
title: Building Blocks
description: Compose your own PowerPoint shell in Svelte 5 from the canvas, a compact toolbar and the slide thumbnail rail, and finish picker and dialog workflows with your own UI through the shared building-block actions.
---

# Building Blocks

`createViewerBuildingBlocks` is the Svelte counterpart of React's
[`useViewerBuildingBlocks`](/react/hooks#public-hooks). It builds the full viewer state behind a
slide canvas with its chrome turned off, and gives you three components to place yourself:

| Block           | Component                  | Notes                                                                    |
| --------------- | -------------------------- | ------------------------------------------------------------------------ |
| Slide canvas    | `BuildingBlocksCanvas`     | The deck loads here. Selection, inline text, gestures and undo all work. |
| Compact toolbar | `BuildingBlocksToolbar`    | History, insert, editing, view and navigation controls.                  |
| Slides pane     | `BuildingBlocksSlidesPane` | The thumbnail rail; selecting a thumbnail shows that slide.              |

Call `createViewerBuildingBlocks` once, synchronously in the script of the component that renders
the blocks (it creates the viewer state, which registers effects and Svelte context, so the blocks
must be rendered by that component or its descendants), and pass a function returning the current
options. Everything it returns is a runes-backed getter, so a template that reads it tracks the
editor.

```svelte
<script lang="ts">
	import {
		BuildingBlocksCanvas,
		BuildingBlocksSlidesPane,
		BuildingBlocksToolbar,
		createViewerBuildingBlocks,
	} from 'pptx-svelte-viewer';

	let { content }: { content: Uint8Array } = $props();
	let dirty = $state(false);

	const blocks = createViewerBuildingBlocks(() => ({
		content,
		canEdit: true,
		onRequest: (request) => openMyPicker(request),
		onDirtyChange: (value) => (dirty = value),
	}));
</script>

<BuildingBlocksToolbar {blocks} />
<div style="display: flex; height: 600px">
	<div style="width: 200px"><BuildingBlocksSlidesPane {blocks} /></div>
	<div style="flex: 1"><BuildingBlocksCanvas {blocks} /></div>
</div>
<button disabled={!dirty}>Save</button>
```

The blocks are view-only until `canEdit: true`, as in every binding, and `blocks.state.mode` reads
`preview` until then. The options function is read live, so `content`, `canEdit`, `locale`,
`hiddenActions`, `fileLimits`, `onRequest` and every callback apply at once;
`limits`, `initialSlide`, `googleFonts` and `allowExternalContent` apply from the next load.
`blocks.state` is `{ mode, loading, error, slideIndex, slideCount, canUndo, canRedo, dirty }` for
your status line, and `blocks.viewerState` is the full viewer state behind the canvas (its `deck`,
`editingApi` and `exportingApi`) for saving, export and everything else. AutoRecover snapshots are
off unless you pass `autosave: true`.

`BuildingBlocksToolbar` takes an optional `include` (`['history', 'insert']`) to render only some
groups. `BuildingBlocksCanvas` takes an optional `theme` (a `ViewerTheme`); without it the canvas
inherits your tokens.

## Host-owned workflows

The blocks mount no file inputs and no dialogs. A toolbar control or a canvas gesture that needs one
raises a request instead: `onRequest(request)` is called and `blocks.actions.pending` holds it until
you finish it with the matching operation or call `actions.cancel()`.

| Request                              | Raised by                                | Finish with                              |
| ------------------------------------ | ---------------------------------------- | ---------------------------------------- |
| `{ kind: 'image' }`                  | Pictures                                 | `actions.insertImage(file)`              |
| `{ kind: 'media' }`                  | Media                                    | `actions.insertMedia(file)`              |
| `{ kind: 'chart' }`                  | Chart                                    | `actions.chart.insert(kind, options)`    |
| `{ kind: 'smartArt' }`               | SmartArt                                 | `actions.smartArt.insert(layout, items)` |
| `{ kind: 'equation', existingOmml }` | Equation, or double-clicking an equation | `actions.equation.insert / update`       |
| `{ kind: 'hyperlink' }`              | Link, Ctrl+K                             | `actions.hyperlink.apply / remove`       |
| `{ kind: 'shortcuts', open }`        | `?`, Ctrl+/, `actions.shortcuts.setOpen` | render `actions.shortcuts.catalogue`     |

Ctrl+F and Ctrl+H on the canvas open your find bar (`actions.findReplace.open`); render it from
`actions.findReplace` (query, options, matches, `next`, `previous`, `replaceCurrent`, `replaceAll`).
The canvas outlines the elements holding matches (`data-pptx-find-match`, `current` for the current
one).

The file and slide show dialogs are groups on `actions` too, with the same shape as React's:
`passwordProtection`, `documentProperties`, `fontEmbedding`, `digitalSignatures`, `setUpSlideShow`,
`customShows`, `broadcast` (reports that it is not supported) and `shortcuts`. Render your own UI
from their state and apply through them; a host raises them from its own controls.
Layout apply and Reset are `actions.layout.apply(path)` and `actions.layout.reset()`.

Every operation returns `{ ok: true }` or `{ ok: false, code }` and runs the viewer's own insert,
find and layout code, so history and `onDirtyChange` behave as in the full viewer.
`actions.availability` says which operations can run now and why not (`read-only`, `protected`,
`no-selection`, ...). The `actions` object has the same shape in every binding; see
[React's building blocks](/react/hooks#picker-dialog-and-layout-workflows-from-your-own-controls) for each operation.

```svelte
<script lang="ts">
	$effect(() => {
		const { pending } = blocks.actions;
		if (pending?.kind === 'image') {
			pickFile('image/*').then((file) =>
				file ? blocks.actions.insertImage(file) : blocks.actions.cancel(),
			);
		}
	});
</script>
```

## Toolbar controls

The toolbar renders `<button data-control="...">` elements inside `<span data-group="...">`, so you
can style them, or skip `BuildingBlocksToolbar` and draw your own buttons from `blocks.toolbar` (the
shared descriptor: groups, controls, labels as dictionary keys, `disabled` and `pressed`) and run
them with `blocks.runToolbarControl(id)`. `hiddenActions` removes controls as in the full viewer:
`undo`, `redo`, `insert` (every insert control), `zoom` and `navigation`.

## Not included

The blocks render no dialogs, panes, ribbon, status bar, notes or presentation chrome, and the
compact toolbar has no controls for the File and Slide Show dialogs (raise those from your own
controls and the `actions` groups). Broadcast cannot be run from the building blocks. Hosts that
need the bundled chrome should use `PowerPointViewer`.
