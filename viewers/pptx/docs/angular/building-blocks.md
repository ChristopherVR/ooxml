---
title: Building Blocks
description: Compose your own PowerPoint shell in Angular from the canvas, a compact toolbar and the slide thumbnail rail, and finish picker and dialog workflows with your own UI through the shared building-block actions.
---

# Building Blocks

`ViewerBuildingBlocks` is the Angular counterpart of React's
[`useViewerBuildingBlocks`](/react/hooks#public-hooks) and vanilla's
[`createViewerBuildingBlocks`](/vanilla/building-blocks). It runs the viewer's own editor services
behind three standalone components and lets you place them yourself:

| Block           | Component                      | Notes                                                                   |
| --------------- | ------------------------------ | ----------------------------------------------------------------------- |
| Slide canvas    | `pptx-building-blocks-canvas`  | The deck loads with the blocks. Selection, inline text, gestures, undo. |
| Compact toolbar | `pptx-building-blocks-toolbar` | History, insert, editing, view and navigation controls (`[include]`).   |
| Slides pane     | `pptx-slides-pane`             | The thumbnail rail; selecting a thumbnail shows that slide.             |

Each takes the blocks through `[blocks]`. The canvas is not named `pptx-slide-canvas` because the
public `SlideCanvasComponent` already owns that tag.

```ts
import {
	BuildingBlocksCanvasComponent,
	BuildingBlocksSlidesPaneComponent,
	BuildingBlocksToolbarComponent,
	createViewerBuildingBlocks,
	VIEWER_BUILDING_BLOCKS_PROVIDERS,
} from 'pptx-angular-viewer';

@Component({
	selector: 'my-shell',
	standalone: true,
	// One composition per host component: the editor services are scoped to it.
	providers: [...VIEWER_BUILDING_BLOCKS_PROVIDERS],
	imports: [
		BuildingBlocksToolbarComponent,
		BuildingBlocksSlidesPaneComponent,
		BuildingBlocksCanvasComponent,
	],
	template: `
		<pptx-building-blocks-toolbar [blocks]="blocks" />
		<pptx-slides-pane [blocks]="blocks" />
		<pptx-building-blocks-canvas [blocks]="blocks" />
	`,
})
export class MyShell {
	readonly content = signal<Uint8Array | null>(null);
	readonly blocks = createViewerBuildingBlocks(() => ({
		content: this.content(),
		canEdit: true,
		onRequest: (request) => this.openMyPicker(request),
		onDirtyChange: (dirty) => (this.saveDisabled = !dirty),
	}));
}
```

`createViewerBuildingBlocks(() => options)` must run in an injection context (a field initialiser
or constructor); it keeps the options in step with the signals the function reads. `blocks.update(options)`
does the same by hand. The blocks are view-only until `canEdit: true`, as in every binding.

`content`, `canEdit`, every callback, `onRequest`, `fileLimits`, `hiddenActions`,
`listConversionScope`, `googleFonts` and `allowExternalContent` apply at once; `limits` and
`initialSlide` apply to the next load. Changing `content` to a different file reloads the deck.

## State and actions

Both are signals:

- `blocks.state()` is `{ mode, loading, error, slideIndex, slideCount, canUndo, canRedo, dirty }`
  (`mode` reads `preview` until editing is granted); `blocks.mode()`, `blocks.loading()` and
  `blocks.error()` are computed from it.
- `blocks.actions()` is the shared `ViewerBuildingBlockActions` (a new object whenever anything it
  reports changes), identical in shape to every other binding's.

## Host-owned workflows

The blocks mount no file inputs and no dialogs. A toolbar control or a canvas gesture that needs one
raises a request instead: `onRequest(request)` is called and `blocks.actions().pending` holds it until
you finish it with the matching operation or call `actions().cancel()`.

| Request                              | Raised by                                  | Finish with                                |
| ------------------------------------ | ------------------------------------------ | ------------------------------------------ |
| `{ kind: 'image' }`                  | Pictures                                   | `actions().insertImage(file)`              |
| `{ kind: 'media' }`                  | Media                                      | `actions().insertMedia(file)`              |
| `{ kind: 'chart' }`                  | Chart                                      | `actions().chart.insert(kind, options)`    |
| `{ kind: 'smartArt' }`               | SmartArt                                   | `actions().smartArt.insert(layout, items)` |
| `{ kind: 'equation', existingOmml }` | Equation, or double-clicking an equation   | `actions().equation.insert / update`       |
| `{ kind: 'hyperlink' }`              | Link, Ctrl+K                               | `actions().hyperlink.apply / remove`       |
| `{ kind: 'shortcuts', open }`        | `?`, Ctrl+/, `actions().shortcuts.setOpen` | render `actions().shortcuts.catalogue`     |

Ctrl+F and Ctrl+H on the canvas open your find bar (`actions().findReplace.open`); render it from
`actions().findReplace`. The canvas marks the matched text itself, as the full viewer and React do: each match is a `<mark data-pptx-find-match>` (`current` for the current match, `match` for the others) inside the run, decided by the shared `findTextHighlightsByElement` / `highlightFindMatches`. Keyboard shortcuts act while focus is in the canvas or on the page
body.

The file and slide show dialogs are groups on `actions()` too: `passwordProtection`,
`documentProperties`, `fontEmbedding`, `digitalSignatures`, `setUpSlideShow`, `customShows`,
`broadcast` (reports that it is not supported) and `shortcuts`. Layout apply and Reset are
`actions().layout.apply(path)` and `actions().layout.reset()`. Every operation returns
`{ ok: true }` or `{ ok: false, code }`; `actions().availability` says which can run now and why not.
See [React's building blocks](/react/hooks#picker-dialog-and-layout-workflows-from-your-own-controls)
for each operation.

## Toolbar groups and hidden controls

`<pptx-building-blocks-toolbar [include]="['history', 'insert']" />` renders only those groups.
`hiddenActions` removes controls as in the full viewer: `undo`, `redo`, `insert`, `zoom` and
`navigation`. Controls are plain `<button data-control="...">` elements inside
`<span data-group="...">`.

## Not included

The blocks render no dialogs, panes, ribbon, status bar, notes or presentation chrome. Starting a
custom show (`actions().customShows.start`) selects the show and shows its first slide; the blocks
mount no slide show overlay. Broadcast cannot be run from the building blocks
(`actions().unmountedDialogActions`). Hosts that need the bundled chrome should use
`<pptx-viewer>`.
