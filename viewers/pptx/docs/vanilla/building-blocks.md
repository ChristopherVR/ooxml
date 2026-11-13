---
title: Building Blocks
description: Compose your own PowerPoint shell in plain JavaScript from the canvas, a compact toolbar and the slide thumbnail rail, and finish picker and dialog workflows with your own UI through the shared building-block actions.
---

# Building Blocks

`createViewerBuildingBlocks` is the vanilla counterpart of React's
[`useViewerBuildingBlocks`](/react/hooks#public-hooks). It runs the full viewer engine behind a
slide canvas with its chrome turned off, and lets you place the pieces yourself:

| Block           | Mount with                    | Notes                                                                    |
| --------------- | ----------------------------- | ------------------------------------------------------------------------ |
| Slide canvas    | `mountCanvas(host)`           | The deck loads here. Selection, inline text, gestures and undo all work. |
| Compact toolbar | `mountToolbar(host, groups?)` | History, insert, editing, view and navigation controls.                  |
| Slides pane     | `mountSlidesPane(host)`       | The thumbnail rail; selecting a thumbnail shows that slide.              |

Each `mount...` returns its disposer; `destroy()` disposes whatever is still mounted.

```ts
import { createViewerBuildingBlocks } from 'pptx-vanilla-viewer';

import 'pptx-vanilla-viewer/styles.css';

const blocks = createViewerBuildingBlocks({
	content: bytes,
	canEdit: true,
	onRequest: (request) => openMyPicker(request),
	onDirtyChange: (dirty) => (saveButton.disabled = !dirty),
});
blocks.mountToolbar(document.querySelector('#toolbar')!);
blocks.mountSlidesPane(document.querySelector('#slides')!);
blocks.mountCanvas(document.querySelector('#canvas')!);
```

The blocks are view-only until `canEdit: true`, as in every binding, and `mode` reads `preview`
until then. `update(options)` applies `content`, `canEdit`, `locale`, `theme`, `customization` and
every callback at once, and options the viewer reads when it uses them (`listConversionScope`,
`limits` on the next load) follow too. `googleFonts`, `allowExternalContent`, `collaboration`,
`fonts`, `registry`, `availableThemes` and `availableLocales` apply the next time the canvas mounts.
`subscribe(listener)` calls `listener` after every change, and `getState()` returns
`{ mode, loading, error, slideIndex, slideCount, canUndo, canRedo, dirty }` for your status line.
`blocks.viewer` is the full [viewer instance](/vanilla/api) behind the canvas (null until it
mounts), for saving, export and everything else.

## Host-owned workflows

The blocks mount no file inputs and no dialogs. A toolbar control or a canvas gesture that needs one
raises a request instead: `onRequest(request)` is called and `blocks.actions.pending` holds it until
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

```ts
blocks.subscribe(() => {
	const { pending } = blocks.actions;
	if (pending?.kind === 'image') {
		pickFile('image/*').then((file) =>
			file ? blocks.actions.insertImage(file) : blocks.actions.cancel(),
		);
	}
});
```

## Toolbar groups and hidden controls

`mountToolbar(host, ['history', 'insert'])` renders only those groups. `hiddenActions` removes
controls as in the full viewer: `undo`, `redo`, `insert` (every insert control), `zoom` and
`navigation`. Controls are plain `<button data-control="...">` elements inside
`<span data-group="...">`, so you can style them or replace the toolbar with your own buttons that
call `actions` and `blocks.viewer`.

## Not included

The blocks render no dialogs, panes, ribbon, status bar, notes or presentation chrome, and have no
`shellState` or collaboration status object yet (use `blocks.viewer`). Broadcast cannot be run from
the building blocks (`actions.unmountedDialogActions`). Hosts that need the bundled chrome should
use `createPptxViewer`.
