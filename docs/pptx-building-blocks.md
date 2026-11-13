# PowerPoint building blocks in every binding

Status: the shared session, React's refactor onto it and the vanilla reference port are in place.
Vue, Svelte and Angular are ported from this note. It is the contract those ports follow; change it
in the same commit as a change to the contract.

Building blocks let a host compose its own PowerPoint shell from the slide canvas, a toolbar and
the slide thumbnail rail, and finish picker and dialog workflows (pictures, media, hyperlinks,
SmartArt, charts, equations, Find / Replace, layouts) with its own UI. Word, Excel and Visio ship
them in all six of their bindings on `ooxml-ui/bindings`; PowerPoint shipped them only in React.
The repository policy is that every binding has them.

## Why PowerPoint differs from Word, Excel and Visio

Word, Excel and Visio have one editor engine in `ooxml-ui` (a custom element or a
`ViewerController`) that every binding mounts, so their building blocks are one session plus thin
per-framework wrappers. PowerPoint has five ports of the editor: each binding keeps its own editor
state and operations (React's hook tree, vanilla's `PptxViewer` and its editor controller, Vue's
composables, Svelte's runes state, Angular's services). The building blocks therefore share
**everything that is not the editor itself**: the gating, the request lifecycle, the operations'
validation and stale-target checks, the find bar session, the toolbar descriptor and the parity
contract. Each binding supplies one port over its own editor.

## The shared session (`src/ui/src/pptx/render/building-blocks/`)

Strict TypeScript, DOM-free apart from `File`, exported from `ooxml-ui/pptx`.

| Module                   | What it holds                                                                                                                                                                                                                                                                                                                                              |
| ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `session.ts`             | `createPptxBuildingBlockSession(port)`: the operations (`insertImage`, `insertMedia`, hyperlink, SmartArt, chart, equation, find, layout), their gating by availability, stale-target checks for the asynchronous inserts, `actions()` (the current `actions` object), `request`, `applyLayout` / `resetSlide` for the toolbar, and `register` for groups. |
| `requests.ts`            | `BuildingBlockRequest` (built from `BuildingBlockRequestMap`), `BUILDING_BLOCK_REQUEST_ACTIONS` (the availability id that gates each kind) and `createBuildingBlockRequests`, a headless controller whose snapshot is the pending request.                                                                                                                 |
| `actions-types.ts`       | The public `ViewerBuildingBlockActions` type every binding returns, and `UseViewerBuildingBlocksActionsOptions` (`onRequest`, `fileLimits`).                                                                                                                                                                                                               |
| `find-layout.ts`         | `BuildingBlockFindReplacePort`, `buildFindReplaceActions` and `buildLayoutActions`.                                                                                                                                                                                                                                                                        |
| `find-bar-controller.ts` | `createFindBarController(host)`: the host find bar's open flag plus the shared `find-replace-session`, as a controller, for ports that keep no find state of their own.                                                                                                                                                                                    |
| `hyperlink-edit-data.ts` | `HyperlinkEditData` (the hyperlink dialog's values), its conversions to and from `actionClick`, and `hyperlinkEditPatch` (refuses unsafe URL schemes).                                                                                                                                                                                                     |
| `toolbar.ts`             | `describeBuildingBlockToolbar(input)`: the compact toolbar's groups, controls (existing dictionary keys, so no new strings), disabled and pressed state, and `runBuildingBlockToolbarControl(id, session, view)`.                                                                                                                                          |
| `action-groups.ts`       | Registration of further action groups (see below).                                                                                                                                                                                                                                                                                                         |

`building-block-actions.ts` (beside the folder) keeps the decisions it already had:
`describeBuildingBlockActions`, the failure codes, file validation and
`BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS`.

### The port

`BuildingBlockSessionPort` is what a binding supplies. Every member is read **at call time**, so an
operation handed out long ago sees the latest deck, slide and selection:

- `availability()`: the binding gathers `BuildingBlockAvailabilityInput` from its own state
  (`ready`, `hasActiveSlide`, `canEdit`, `editing`, `canInsert`, `hasSelection`, ...) and calls
  `describeBuildingBlockActions`. See `vanillaAvailability` in
  `viewers/pptx/packages/vanilla/src/building-blocks/session-port.ts` and the React hook for the
  two existing mappings.
- `target()`: `{ deck, slideId, slideIndex }`, the identity an asynchronous insert must still
  target when it lands (`deck` is the binding's handler).
- `canvasSize`, `selectedElement`, `slideCount`, `activeSlide`, `scopedLayouts`, `fileLimits`,
  `onRequest`.
- The editor operations: `addElement`, `insertMediaFile(file, sameTarget)`, `applyHyperlink`,
  `removeHyperlink`, `insertSmartArt`, `insertChart`, `insertEquation`, `updateEquation`,
  `applyLayout(path): Promise<void>`, `findReplace()`.

## What stays per binding

- **Template and block components.** The canvas, toolbar and slides pane are the binding's own
  components over its existing canvas, rail and (for React) ribbon.
- **Lifecycle.** Creating the session once per building-blocks instance, disposing it, and reading
  the request controller and the find bar controller through the binding's one store adapter
  (`useController`, `controllerState`, `controllerSignal`, `subscribe`).
- **Focus and portals.** The blocks mount no dialogs; anything that would open one becomes a
  request. Keyboard shortcuts stay with the canvas.
- **Turning the binding's own dialog openers into requests.** React clears its dialog flags in
  effects and calls `session.request`. Vanilla offers each dialog to
  `PptxViewer.requestHostDialog` (overridden by `BuildingBlocksViewer`). A port does the same at
  its own dialog openers: hyperlink (Ctrl+K, the element menu), equation (double-click), find
  (Ctrl+F / Ctrl+H toggles the find bar controller), and the Insert controls.

## Public contract per binding

The names follow docx, xlsx and visio.

| Binding | Entry                                                                                                                           | Blocks                                                                                                     | Package path (new files)                                               |
| ------- | ------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| React   | `useViewerBuildingBlocks(input)` returning `{ toolbarProps, canvasProps, slidesPaneProps, actions, mode, loading, error, ... }` | `Toolbar` (full ribbon), `SlideCanvas`, `SlidesPane`, `SlideThumbnail`                                     | existing `packages/react/src/viewer/hooks/useViewerBuildingBlocks*.ts` |
| Vanilla | `createViewerBuildingBlocks(options)` returning `ViewerBuildingBlocks`                                                          | `mountCanvas(host)`, `mountToolbar(host, groups?)`, `mountSlidesPane(host)`                                | `packages/vanilla/src/building-blocks/`                                |
| Vue     | `useViewerBuildingBlocks(options: MaybeRefOrGetter<ViewerBuildingBlocksOptions>)` composable                                    | `<SlideCanvas>`, `<BuildingBlocksToolbar>`, `<SlidesPane>` components taking the composable's result       | `packages/vue/src/building-blocks/`                                    |
| Svelte  | `createViewerBuildingBlocks(() => options)` returning an object with runes-backed getters (`.actions`, `.state`)                | `SlideCanvas.svelte`, `BuildingBlocksToolbar.svelte`, `SlidesPane.svelte` taking `blocks`                  | `packages/svelte/src/building-blocks/`                                 |
| Angular | `ViewerBuildingBlocks` injectable plus `createViewerBuildingBlocks(() => options)`; `actions()` and `state()` are signals       | `pptx-slide-canvas`, `pptx-building-blocks-toolbar`, `pptx-slides-pane` standalone components (`[blocks]`) | `packages/angular/src/building-blocks/`, exported from `public-api.ts` |

Every entry:

- takes `content`, `canEdit` (view-only until `true`), `onRequest`, `fileLimits`, `hiddenActions`,
  `onDirtyChange`, `onOpenLayoutGallery` and the binding's viewer options;
- returns `actions: ViewerBuildingBlockActions` (the shared type, re-exported) and a state with
  `loading`, `error` and the slide count; the view-only ports also report `mode`;
- renders the compact toolbar from `describeBuildingBlockToolbar` and runs its controls with
  `runBuildingBlockToolbarControl` (React keeps its ribbon `Toolbar`);
- marks its blocks with `data-pptx-block="canvas" | "toolbar" | "slides-pane"` and each toolbar
  control with `data-control="<id>"` inside `data-group="<group>"`;
- mounts no dialog, picker, ribbon, status bar or notes pane.

## How the vanilla reference port is built

`packages/vanilla/src/building-blocks/`:

- `blocks-viewer.ts`: `BuildingBlocksViewer`, the full `PptxViewer` with `requestHostDialog`
  overridden. The canvas block creates it with `showToolbar`, `showThumbnails` and
  `showCompatibilityToasts` off and the status bar, inspector, notes, title bar and quick access
  toolbar hidden through `customization.hiddenPanels`.
- `session-port.ts`: the port over the viewer's store and editor controller.
- `create-viewer-building-blocks.ts`: the session, the find bar controller (committing replaces
  through `editor.commitElementUpdates`), `subscribe`, `update` and the three `mount...` helpers.
- `toolbar-view.ts`: keyed DOM for the toolbar descriptor, updated in place so focus survives.
- `slides-pane-block.ts`: `createSlidesPane` fed from the store.

A Vue, Svelte or Angular port has the same four pieces: the port over its editor, the dialog
openers that become requests, the find bar controller wired to its commit and navigation, and its
native block components.

## Adding an action group

The agents adding host operations for the remaining dialogs (password protection, document
properties, font embedding, digital signatures, set up slide show, custom shows, broadcast,
keyboard shortcuts) plug in without editing the session:

1. Put the group's DOM-free operations and types in a new file in `src/ui/src/pptx/render/`.
2. Add its key and type to `BuildingBlockActionGroupMap` (`action-groups.ts`) and, if it raises a
   request, its kind to `BuildingBlockRequestMap` and `BUILDING_BLOCK_REQUEST_ACTIONS`
   (`requests.ts`). If it needs its own availability id, add it to `BuildingBlockActionId` and
   `describeBuildingBlockActions` in `building-block-actions.ts`.
3. Define it with `defineBuildingBlockActionGroup({ id, create(context, port) })`. `context` has
   `refuse(id)`, `request(next)` and `complete(kind, result)`; `port` is the group's own port.
4. Each binding registers it once with `session.register(group, port)`; it then appears as
   `actions[id]` in every binding. Remove its handler from `BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS`
   when it lands, since a host can then finish it.

React registers in a new hook file next to `useViewerBuildingBlocksActions.ts`: the hook's result
carries `session`, so the new hook registers once (`useState(() => session.register(group, port))`
with a port that reads a ref, and the unregister in an unmount effect), and
`useViewerBuildingBlocks` already reads `session.actions()` last, after every hook ran. The only
edit to `useViewerBuildingBlocks.ts` is the one call to the new hook. Vanilla registers in
`create-viewer-building-blocks.ts`.

## Parity is enforced by a shared contract test

`viewers/pptx/packages/shared/test-utils/building-blocks-contract.mjs` (typed by its `.d.mts`) runs
the same assertions against any binding through a harness: view-only refusal, edit permission,
selection-bound availability, a toolbar request and its cancel, validation, an insert that marks
the deck dirty, Find / Replace with the host bar, layout refusal, taking permission away again, and
for compact-toolbar ports the block markers and the control order (`COMPACT_TOOLBAR_CONTROLS`).

- React: `packages/react/src/viewer/hooks/useViewerBuildingBlocks.contract.test.tsx`.
- Vanilla: `packages/vanilla/src/building-blocks/create-viewer-building-blocks.test.ts`.
- Each port adds one test that builds a harness (`mount`, `update`, `actions`, `state`,
  `clickPictures`, `settle`, `destroy`, `toolbar: true`) and calls `runBuildingBlocksContract`. It
  must pass unchanged; a port that needs the contract changed is a design change for this note.

Ports also add a demo page under `demos/pptx/demo-<binding>` once a framework-neutral browser spec
for building blocks exists (today `e2e/pptx/host-owned-building-blocks.spec.ts` drives React only).

## Known gaps

- `mode` differs: React reports the viewer mode (`edit` for a view-only deck, since permission is
  `canEdit`), vanilla reports `preview` until editing is granted. The contract does not assert it.
- The vanilla canvas does not highlight the host find bar's matches (its highlight reads the
  ribbon's own find session); React's canvas does.
- Vanilla's state has no `shellState` or `collaboration` yet; use `blocks.viewer` for collaboration.
- Vanilla's `actions.layout.apply` resolves once the layout landed or could not be resolved; a
  layout the handler rejects still reports `{ ok: true }` (React reports `unknown-layout`).
