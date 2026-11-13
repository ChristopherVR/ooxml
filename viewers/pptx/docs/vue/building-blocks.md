---
title: Building Blocks
description: Compose your own PowerPoint shell in Vue from the canvas, a compact toolbar and the slide thumbnail rail, and finish picker and dialog workflows with your own UI through the shared building-block actions.
---

# Building Blocks

`useViewerBuildingBlocks` is the Vue counterpart of React's
[`useViewerBuildingBlocks`](/react/hooks#public-hooks). It runs the full viewer engine behind a
slide canvas with its chrome turned off, and lets you place the pieces yourself:

| Block           | Component                  | Notes                                                                    |
| --------------- | -------------------------- | ------------------------------------------------------------------------ |
| Slide canvas    | `BuildingBlocksCanvas`     | The deck loads here. Selection, inline text, gestures and undo all work. |
| Compact toolbar | `BuildingBlocksToolbar`    | History, insert, editing, view and navigation controls (`groups` prop).  |
| Slides pane     | `BuildingBlocksSlidesPane` | The thumbnail rail; selecting a thumbnail shows that slide.              |

The components are named apart from the exported low-level `SlideCanvas` and `SlidesPane`, which
stay as they were. Each takes the composable's result as `blocks`.

```vue
<script setup lang="ts">
import {
	BuildingBlocksCanvas,
	BuildingBlocksSlidesPane,
	BuildingBlocksToolbar,
	useViewerBuildingBlocks,
} from 'pptx-vue-viewer';
import 'pptx-vue-viewer/styles';

const props = defineProps<{ bytes: Uint8Array }>();
const blocks = useViewerBuildingBlocks(() => ({
	content: props.bytes,
	canEdit: true,
	onRequest: (request) => openMyPicker(request),
	onDirtyChange: (dirty) => (saveDisabled.value = !dirty),
}));
</script>

<template>
	<BuildingBlocksToolbar :blocks="blocks" />
	<div style="display: flex; height: 600px">
		<div style="width: 200px; overflow: auto"><BuildingBlocksSlidesPane :blocks="blocks" /></div>
		<div style="flex: 1; position: relative"><BuildingBlocksCanvas :blocks="blocks" /></div>
	</div>
</template>
```

The options may be an object, a ref or a getter. The blocks are view-only until `canEdit: true`,
and `mode` reads `preview` until then. Every option applies at once, the `limits` and
`initialSlide` ones from the next load. The result holds `actions`, `state`
(`{ mode, loading, error, slideIndex, slideCount, canUndo, canRedo, dirty }`), `mode`, `loading`
and `error` as computed refs, and `viewer`, the full viewer's imperative handle (saving, export;
null until the canvas mounts). The blocks need the `vue-i18n` plugin installed, as the full viewer
does.

## Host-owned workflows

The blocks mount no file inputs and no dialogs. A toolbar control or a canvas gesture that needs
one raises a request instead: `onRequest(request)` is called and `actions.pending` holds it until
you finish it with the matching operation or call `actions.cancel()`.

| Request                              | Raised by                                   | Finish with                              |
| ------------------------------------ | ------------------------------------------- | ---------------------------------------- |
| `{ kind: 'image' }`                  | Pictures                                    | `actions.insertImage(file)`              |
| `{ kind: 'media' }`                  | Media                                       | `actions.insertMedia(file)`              |
| `{ kind: 'chart' }`                  | Chart                                       | `actions.chart.insert(kind, options)`    |
| `{ kind: 'smartArt' }`               | SmartArt                                    | `actions.smartArt.insert(layout, items)` |
| `{ kind: 'equation', existingOmml }` | Equation, or double-clicking an equation    | `actions.equation.insert / update`       |
| `{ kind: 'hyperlink' }`              | Link, Ctrl+K, the element menu's Link entry | `actions.hyperlink.apply / remove`       |
| `{ kind: 'shortcuts', open }`        | `?`, Ctrl+/, `actions.shortcuts.setOpen`    | render `actions.shortcuts.catalogue`     |

Ctrl+F and Ctrl+H on the canvas toggle your find bar (`actions.findReplace.open`); render it from
`actions.findReplace` (query, options, matches, `next`, `previous`, `replaceCurrent`,
`replaceAll`). The canvas marks the matched text itself, as the full viewer and React do: each match is a `<mark data-pptx-find-match>` (`current` for the current match, `match` for the others) inside the run, decided by the shared `findTextHighlightsByElement` / `highlightFindMatches`.

The file and slide show dialogs are groups on `actions` too, with the same shape as React's:
`passwordProtection`, `documentProperties`, `fontEmbedding`, `digitalSignatures`, `setUpSlideShow`,
`customShows`, `broadcast` (reports that it is not supported) and `shortcuts`. Layout apply and
Reset are `actions.layout.apply(path)` and `actions.layout.reset()`.

Every operation returns `{ ok: true }` or `{ ok: false, code }` and runs the viewer's own insert,
find and layout code, so history and `onDirtyChange` behave as in the full viewer.
`actions.availability` says which operations can run now and why not. The `actions` object has the
same shape in every binding; see
[React's building blocks](/react/hooks#picker-dialog-and-layout-workflows-from-your-own-controls).

```ts
watch(
	() => blocks.actions.value.pending,
	async (pending) => {
		if (pending?.kind === 'image') {
			const file = await pickFile('image/*');
			file ? await blocks.actions.value.insertImage(file) : blocks.actions.value.cancel();
		}
	},
);
```

## Toolbar groups and hidden controls

`<BuildingBlocksToolbar :groups="['history', 'insert']" />` renders only those groups.
`hiddenActions` removes controls as in the full viewer: `undo`, `redo`, `insert`, `zoom` and
`navigation`. Controls are plain `<button data-control="...">` elements inside
`<span data-group="...">`, so you can style them or replace the toolbar with your own buttons that
call `actions` and `blocks.viewer`.

## Not included

The blocks render no dialogs, panes, ribbon, status bar, notes or presentation chrome. Broadcast
cannot be run from the building blocks (`actions.unmountedDialogActions`). Hosts that need the
bundled chrome should use `PowerPointViewer`.
