# PowerPoint building blocks in every binding

Status: all five bindings (React, Vue, Svelte, Angular, vanilla) have building blocks on the
shared session, pass the parity contract and run `e2e/pptx/building-blocks.spec.ts`. This note is
the contract they follow; change it in the same commit as a change to the contract.

Building blocks let a host compose its own PowerPoint shell from the slide canvas, a toolbar and
the slide thumbnail rail, and finish picker and dialog workflows (pictures, media, hyperlinks,
SmartArt, charts, equations, Find / Replace, layouts, and the file and slide show dialogs) with its
own UI. Word, Excel and Visio ship them in all six of their bindings on `ooxml-ui/bindings`;
PowerPoint shipped them only in React. The repository policy is that every binding has them.

## Why PowerPoint differs from Word, Excel and Visio

Word, Excel and Visio have one editor engine in `ooxml-ui` (a custom element or a
`ViewerController`) that every binding mounts, so their building blocks are one session plus thin
per-framework wrappers. PowerPoint has five ports of the editor: each binding keeps its own editor
state and operations (React's hook tree, vanilla's `PptxViewer` and its editor controller, Vue's
composables, Svelte's runes state, Angular's services). The building blocks therefore share
**everything that is not the editor itself**: the gating, the request lifecycle, the operations'
validation and stale-target checks, the dialog groups, the find bar session, the toolbar
descriptor and the parity contract. Each binding supplies ports over its own editor.

## The shared session (`src/ui/src/pptx/render/building-blocks/`)

Strict TypeScript, DOM-free apart from `File`, exported from `ooxml-ui/pptx`.

| Module                   | What it holds                                                                                                                                                                                                                                                                             |
| ------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `session.ts`             | `createPptxBuildingBlockSession(port)`: the core operations, their gating, the stale-target checks of the asynchronous inserts, `actions()` (the current `actions` object), `request`, `announce`, `applyLayout` / `resetSlide` for the toolbar, and `register` for groups.               |
| `requests.ts`            | `BuildingBlockRequest` (from `BuildingBlockRequestMap`), `BUILDING_BLOCK_REQUEST_ACTIONS` (the availability id gating each kind), `BUILDING_BLOCK_VIEWABLE_REQUESTS`, `BuildingBlockAvailability` and `createBuildingBlockRequests` (a controller whose snapshot is the pending request). |
| `groups/`                | The built-in action groups (below), `BuildingBlockGroupPorts`, `BUILDING_BLOCK_GROUP_IDS` and `registerBuildingBlockGroups(session, ports)`.                                                                                                                                              |
| `actions-types.ts`       | The public `ViewerBuildingBlockActions` type every binding returns, and `UseViewerBuildingBlocksActionsOptions` (`onRequest`, `fileLimits`).                                                                                                                                              |
| `find-layout.ts`         | `BuildingBlockFindReplacePort`, `buildFindReplaceActions`, and `buildLayoutActions`.                                                                                                                                                                                                      |
| `find-bar-controller.ts` | `createFindBarController(host)`: the host find bar's open flag plus the shared `find-replace-session`, as a controller, for ports that keep no find state of their own.                                                                                                                   |
| `hyperlink-edit-data.ts` | `HyperlinkEditData` (the hyperlink dialog's values), its conversions to and from `actionClick`, and `hyperlinkEditPatch` (refuses unsafe URL schemes).                                                                                                                                    |
| `toolbar.ts`             | `describeBuildingBlockToolbar(input)`: the compact toolbar's groups and controls (existing dictionary keys, so no new strings), and `runBuildingBlockToolbarControl(id, session, view)`.                                                                                                  |
| `mode.ts`                | `resolveBuildingBlockMode(mode, canEdit)`: the `mode` every binding reports (below).                                                                                                                                                                                                      |
| `shell.ts`               | `createBuildingBlockShell(read)` and `describeBuildingBlockShell`: the `shellState` (`resolveCollaborationShellState`) and `collaborationState` (`buildCollaborationState`) every binding reports, `BUILDING_BLOCK_SHELL_IDLE` before a canvas mounts, and the identity rule (below).     |
| `live-options.ts`        | `BUILDING_BLOCK_LIVE_OPTIONS`, `collaborationOptionChange`, `googleFontsOptionChanged` and `applyBuildingBlockLiveOptions(previous, next, port)`: how an imperative engine (vanilla) applies the options it reads once.                                                                   |
| `action-groups.ts`       | `BuildingBlockActionGroupMap`, `defineBuildingBlockActionGroup` and the group context.                                                                                                                                                                                                    |

`building-block-actions.ts` (beside the folder) keeps `describeBuildingBlockActions`, the failure
codes, file validation and `BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS` (today only
`onOpenBroadcastDialog`). The dialog decisions are the `building-block-password`,
`-document-properties`, `-font-embedding`, `-signatures`, `-show-settings`, `-custom-shows` and
`-show-dialogs` modules beside it; the groups bind them to the session.

### The port

`BuildingBlockSessionPort` is what a binding supplies. Every member is read **at call time**, so an
operation handed out long ago sees the latest deck, slide and selection:

- `availabilityInput()`: `BuildingBlockAvailabilityInput` from the binding's state (`ready`,
  `hasActiveSlide`, `canEdit`, `editing`, `canInsert`, `hasSelection`, ...). The session derives
  every availability entry from it (`describeBuildingBlockAvailability`: the core workflows and the
  slide show dialogs). See `vanillaAvailabilityInput` in
  `viewers/pptx/packages/vanilla/src/building-blocks/session-port.ts` and the React hook.
- `target()`: `{ deck, slideId, slideIndex }`, the identity an asynchronous insert must still
  target when it lands (`deck` is the binding's handler).
- `canvasSize`, `selectedElement`, `slideCount`, `activeSlide`, `scopedLayouts`, `fileLimits`,
  `onRequest`, and `onCancel` (close the shortcut help the binding mirrors).
- The editor operations: `addElement`, `insertMediaFile(file, sameTarget)`, `applyHyperlink`,
  `removeHyperlink`, `insertSmartArt`, `insertChart`, `insertEquation`, `updateEquation`,
  `applyLayout(path): Promise<void>` (rejects when the layout was not applied; the session reports
  `unknown-layout`), `findReplace()`.

### The groups

Every binding registers all of them (`registerBuildingBlockGroups`), so each is a required key of
`ViewerBuildingBlockActions`; the contract checks they are there.

| `actions` key        | Operations                                                                                           | Request kind                                     | Port (`BuildingBlockGroupPorts`)                                                                                   |
| -------------------- | ---------------------------------------------------------------------------------------------------- | ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| `passwordProtection` | `state`, `setOpenPassword`, `clearOpenPassword`, `setMarkedFinal`, `setModifyPassword` (unsupported) | `passwordProtection`, viewable                   | `input()`, `customProperties()`, `applyOpenPassword(password or null)`, `applyCustomProperties`, `markDirty`       |
| `documentProperties` | `current`, `apply(patch)`                                                                            | `documentProperties`, viewable                   | `current()`, `applyUpdate(update)`                                                                                 |
| `fontEmbedding`      | `state`, `setEmbedFonts(enabled)`                                                                    | `fontEmbedding`, viewable                        | `input()` (what `describeFontEmbeddingState` reads), `apply(enabled)`, `markDirty`                                 |
| `digitalSignatures`  | `state`, `list()`, `remove()`, `sign()` (unsupported)                                                | `digitalSignatures`, viewable                    | `signed()`, `count()`, `content()` (the loaded bytes), `markDirty`                                                 |
| `setUpSlideShow`     | `current`, `apply(patch)`                                                                            | `setUpSlideShow`                                 | `presentationProperties()`, `slideCount()`, `customShows()`, `save(settings)`                                      |
| `customShows`        | `shows`, `slides`, `activeId`, `create`, `rename`, `setSlides`, `copy`, `remove`, `start`            | `customShows`                                    | `customShows()` (including edits stored this tick), `slides()`, `activeId()`, `commit(shows, slides)`, `start(id)` |
| `broadcast`          | `{ supported: false, reason: 'unsupported', requires: 'collaboration-service' }`                     | none; the binding drops its Broadcast flag       | none                                                                                                               |
| `shortcuts`          | `open`, `setOpen`, `toggle`, `catalogue`                                                             | `{ kind: 'shortcuts', open }` through `announce` | `open()`, `setOpen(open)`, `toggle()`, `translate(key)`                                                            |

A viewable request is raised whenever the deck is loaded, also read-only or protected; its apply
operations are refused there. The other requests need their action to be available. The shortcut
help is a toggle the host mirrors: every change of the binding's flag (`?`, Ctrl+/, `setOpen`) is
passed to `session.announce(request, open)`, which makes an open request pending, settles it on
close and calls `onRequest` for both.

### Shell state and collaboration

Every binding's state carries `shellState` (`{ canEdit, status, remoteUsers, connectedCount }`) and
`collaborationState` (the shared `CollaborationState`, `COLLABORATION_OFF` while not sharing).
The binding reads the raw inputs from its own collaboration session (the room's status, read-only
role or readiness, remote presences, the host's `canEdit`, whether a source is loading or failed)
and the shared shell derives both: vanilla and Svelte call `createBuildingBlockShell` over their
session, Vue over the engine's `shell()` (its viewer's `useCollaboration` result), React
`describeBuildingBlockShell` in a memo, and Angular maps its custom-shell service's `state()` and
`collaboration.state()` signals, which run the same two functions. The shell keeps its identity
while nothing a shell shows changed (cursor moves alone keep it), so a computed, signal or getter
only re-renders on real changes.

Collaboration is an option: setting `collaboration` joins the room, a new config object rejoins
(compared by `resolveCollaborationConfig` identity), clearing it leaves, and none of these remounts
the canvas or reloads the deck. React, Vue, Svelte and Angular run their viewer's own collaboration
wiring (the `createCollaborationSessionController` lifecycle, the document sync and the cursor and
remote-selection overlays); vanilla's `update()` calls `applyBuildingBlockLiveOptions`, which
starts or stops its viewer's session. A `viewer` role or an unsynced room takes editing away
(`mode` reads `preview`); the host's `canEdit` stays a ceiling. Nothing connects unless the host
passes a `collaboration` config.

### Mode

Every binding reports the effective `mode`: `preview` while the blocks may not edit (no `canEdit`,
a collaboration viewer role), otherwise the viewer's mode (`edit`, `present`, `master`). React maps
its internal mode with `resolveBuildingBlockMode(mode, canEdit)`; `toolbarProps.mode` and
`canvasProps.mode` keep the viewer's own. Vanilla's `viewerModeOf` already has these semantics. The
contract asserts `preview`, then `edit`, then `preview` again.

## What stays per binding

- **Template and block components.** The canvas, toolbar and slides pane are the binding's own
  components over its existing canvas, rail and (for React) ribbon.
- **Lifecycle.** Creating the session once per building-blocks instance, registering the groups,
  disposing it, and reading the request controller, the find bar controller and the shortcut flag
  through the binding's one store adapter (`useController`, `controllerState`,
  `controllerSignal`, `subscribe`).
- **Focus and portals.** The blocks mount no dialogs; anything that would open one becomes a
  request. Keyboard shortcuts stay with the canvas.
- **Turning the binding's own dialog openers into requests.** React clears its dialog flags in
  effects and calls `session.request` (`useViewerBuildingBlocksActions.ts`,
  `useViewerBuildingBlocks-file-dialog-actions.ts`, `useViewerBuildingBlocks-show-dialog-actions.ts`).
  Vanilla offers each dialog to `PptxViewer.requestHostDialog` (overridden by
  `BuildingBlocksViewer`). A port does the same at its own openers: hyperlink (Ctrl+K, the element
  menu), equation (double-click), find (Ctrl+F / Ctrl+H toggle the find bar controller), shortcuts
  (`?`, Ctrl+/ toggle the shortcut flag), and the Insert, File and Slide Show controls.
- **Find highlights.** Every binding marks the matched text, in the blocks and in its full viewer.
  The decision is shared (`render/find-text-highlight.ts`): `findTextHighlightsByElement(matches,
  current, slide)` groups the matches of the active slide by element and segment, and
  `highlightFindMatches(paragraphs, highlights)` splits the runs of `buildParagraphs` at the match
  boundaries, setting `findMatch` (`match` or `current`) on the highlighted pieces (React splits
  its segment text with `splitTextByFindHighlights`, the same rule). A binding's text renderer
  only wraps such a piece in `<mark data-pptx-find-match="match|current">` styled with
  `findMatchMarkStyle`. Only the editing canvas passes highlights (React and vanilla through
  props, Vue through an injection, Svelte through a context, Angular through a provider on
  `SlideCanvasComponent`), so thumbnails and the show stage draw none. As in React, only a
  shape's own text is marked; matches in tables, SmartArt, group members, charts and notes still
  select their element but draw no mark. No element is outlined.
- **Media from a host file.** Ports whose editor has no file-based media insert use the shared
  `mediaElementFromFile(file, canvasSize, doc?)` (`media-element-from-file.ts`).
- **Dialog flag routing.** A flag that becomes a request is handled by an effect or watcher that
  tracks the flag alone, clears it and then raises: React `useDialogFlagRequest`, Vue
  `routeDialogFlag` (a sync watcher, `engine.ts`), Svelte `routeDialogFlag`
  (`route-dialog-flag.svelte.ts`), Angular `routeDialogFlag` (`route-dialog-flag.ts`). Raising
  notifies the session's store; an effect that also tracked the render's objects re-ran with the
  flag still set and looped under React 18. Each binding has a regression test that drives a flag
  through a real request store and its store adapter.

## Public contract per binding

The names follow docx, xlsx and visio.

| Binding | Entry                                                                                                                                                            | Blocks (exported names)                                                                                                                                                                                       | Files                                                                  |
| ------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| React   | `useViewerBuildingBlocks(input)` returning `{ toolbarProps, canvasProps, slidesPaneProps, actions, mode, loading, error, ... }`                                  | `Toolbar` (full ribbon), `SlideCanvas`, `SlidesPane`, `SlideThumbnail`, spread with the hook's props                                                                                                          | `packages/react/src/viewer/hooks/useViewerBuildingBlocks*.ts`          |
| Vanilla | `createViewerBuildingBlocks(options)` returning `ViewerBuildingBlocks` (`actions`, `getState()`, `subscribe`, `update`)                                          | `mountCanvas(host)`, `mountToolbar(host, groups?)`, `mountSlidesPane(host)`                                                                                                                                   | `packages/vanilla/src/building-blocks/`                                |
| Vue     | `useViewerBuildingBlocks(options: MaybeRefOrGetter<ViewerBuildingBlocksOptions>)`; `actions` and `state` are refs                                                | `BuildingBlocksCanvas`, `BuildingBlocksToolbar`, `BuildingBlocksSlidesPane` (`:blocks`)                                                                                                                       | `packages/vue/src/building-blocks/`                                    |
| Svelte  | `createViewerBuildingBlocks(() => options)` with runes-backed getters (`.actions`, `.state`)                                                                     | `BuildingBlocksCanvas`, `BuildingBlocksToolbar`, `BuildingBlocksSlidesPane` (`blocks` prop)                                                                                                                   | `packages/svelte/src/building-blocks/`                                 |
| Angular | `ViewerBuildingBlocks` (provided with `VIEWER_BUILDING_BLOCKS_PROVIDERS`) and `createViewerBuildingBlocks(() => options)`; `actions()` and `state()` are signals | `BuildingBlocksCanvasComponent` (`pptx-building-blocks-canvas`), `BuildingBlocksToolbarComponent` (`pptx-building-blocks-toolbar`), `BuildingBlocksSlidesPaneComponent` (`pptx-slides-pane`), each `[blocks]` | `packages/angular/src/building-blocks/`, exported from `public-api.ts` |

Every entry:

- takes `content`, `canEdit` (view-only until `true`), `onRequest`, `fileLimits`, `hiddenActions`,
  `onDirtyChange`, `onOpenLayoutGallery` and the binding's viewer options;
- returns `actions: ViewerBuildingBlockActions` (the shared type, re-exported) with every group,
  and a state with the effective `mode`, `loading`, `error`, the slide count, `shellState` and
  `collaborationState`;
- takes `collaboration` (joined, rejoined and left live) and `googleFonts` / `allowExternalContent`
  (applied live, locking their Trust Center settings);
- renders the compact toolbar from `describeBuildingBlockToolbar` and runs its controls with
  `runBuildingBlockToolbarControl` (React keeps its ribbon `Toolbar`);
- marks its blocks with `data-pptx-block="canvas" | "toolbar" | "slides-pane"` and each toolbar
  control with `data-control="<id>"` inside `data-group="<group>"`;
- mounts no dialog, picker, ribbon, status bar or notes pane.

### Which options apply without a remount

`BUILDING_BLOCK_LIVE_OPTIONS` lists the options every binding applies at once; "at once" means the
mounted canvas, the open deck, its edits and undo history stay.

| Option                                                                               | React                      | Vanilla                                               | Vue                        | Svelte                     | Angular                    |
| ------------------------------------------------------------------------------------ | -------------------------- | ----------------------------------------------------- | -------------------------- | -------------------------- | -------------------------- |
| `content`, `canEdit`, every callback, `onRequest`, `fileLimits`, `hiddenActions`     | at once                    | at once                                               | at once                    | at once                    | at once                    |
| `listConversionScope`, `userName` / `authorName`                                     | at once                    | at once (read live)                                   | at once                    | at once                    | at once                    |
| `limits`, `initialSlide`                                                             | next load                  | next load (read live)                                 | next load                  | next load                  | next load                  |
| `locale`, `theme`, `customization`                                                   | at once                    | at once (`setLocale`, `setTheme`, `setCustomization`) | at once                    | `locale` at once           | not offered                |
| `googleFonts`, `allowExternalContent`                                                | at once (locks the option) | at once (locks the option)                            | at once (locks the option) | at once (locks the option) | at once (locks the option) |
| `collaboration`                                                                      | at once                    | at once (`applyBuildingBlockLiveOptions`)             | at once                    | at once                    | at once                    |
| `fonts`, `registry`, `availableThemes`, `availableLocales`, `inspectorInitiallyOpen` | not offered                | next canvas mount                                     | as the viewer              | not offered                | not offered                |

"Locks the option" means the host value also locks the Trust Center setting
(`trust.allowFontDownloads`, `trust.allowExternalContent`), as the full viewer does. Each binding
has a test for it (React `useViewerBuildingBlocks.google-fonts.test.tsx`, vanilla
`create-viewer-building-blocks.live.test.ts`, Svelte `create-viewer-building-blocks.svelte.test.ts`,
Angular `viewer-building-blocks.contract.test.ts`; Vue forwards both to its viewer, whose
`useViewerCustomization` watchers apply them). The contract checks every binding keeps the deck
across these updates.

## How the vanilla reference port is built

`packages/vanilla/src/building-blocks/`:

- `blocks-viewer.ts`: `BuildingBlocksViewer`, the full `PptxViewer` with `requestHostDialog`
  overridden. The canvas block creates it with `showToolbar`, `showThumbnails` and
  `showCompatibilityToasts` off and the status bar, inspector, notes, title bar and quick access
  toolbar hidden through `customization.hiddenPanels`.
- `session-port.ts`: the session port over the viewer's store and editor controller.
- `group-ports.ts`: the ports of the dialog groups over the same.
- `viewer-options.ts`: the viewer's options as a live view (a `Proxy`) of the latest building-block
  options (its collaboration callbacks also notify the blocks), and the options `update()` pushes
  explicitly, including `applyBuildingBlockLiveOptions` over `BuildingBlocksViewer.liveOptionPort`.
- `shell-state.ts`: `createBuildingBlockShell` over the viewer's store and collaboration session.
- `create-viewer-building-blocks.ts`: the session, group registration, the find bar controller
  (committing replaces through `editor.commitElementUpdates`, its matches copied to the viewer
  store's `findMatches` / `findMatchIndex`, which the canvas marks), the shortcut flag, `subscribe`,
  `update` and the three `mount...` helpers.
- `toolbar-view.ts`: keyed DOM for the toolbar descriptor, updated in place so focus survives.
- `slides-pane-block.ts`: `createSlidesPane` fed from the store.

## Porting a binding (how Vue, Svelte and Angular were ported)

1. Create `packages/<binding>/src/building-blocks/` with the session port over the binding's editor
   (`session-port.ts`, mirror vanilla's), the group ports (`group-ports.ts`), the entry
   (`useViewerBuildingBlocks`, `createViewerBuildingBlocks` or `ViewerBuildingBlocks`) and the three
   block components. Per instance: `createPptxBuildingBlockSession(port)`, then
   `registerBuildingBlockGroups(session, ports)`, then `createFindBarController(host)`. Read
   `session.requests`, the find bar and the shortcut flag through the binding's store adapter.
2. Route the binding's dialog openers to `session.request`, Ctrl+F to the find bar, `?` / Ctrl+/ to
   the shortcut flag plus `session.announce`, and drop the Broadcast flag.
3. Report `mode` with `resolveBuildingBlockMode`, render the compact toolbar from
   `describeBuildingBlockToolbar`, pass the find bar's `findTextHighlightsByElement` to the canvas, and set the
   `data-pptx-block` and `data-control` markers. Route dialog flags as described above.
4. Export the entry, its option and state types and the shared `ViewerBuildingBlockActions`,
   `BuildingBlockRequest` and `BuildingBlockResult` types from the package root. The full viewer's
   props do not change, so `viewer-props-manifest.ts` stays as it is.
5. Tests: one contract test (below), plus unit tests for anything binding-specific (vanilla's
   `create-viewer-building-blocks.live.test.ts` is the model: highlights, a rejected layout,
   options applied without a remount, the shortcut chord).
6. Demo and browser spec: add the `?blocks=1` page to `demos/pptx/demo-<binding>` (copy
   `demos/pptx/demo-vanilla/src/building-blocks-demo.ts`: the binding's blocks plus
   `mountBuildingBlocksPanel` from `demos/pptx/shared/building-blocks-panel.ts`, refreshed after
   every update), and add `metadata: { buildingBlocks: true }` to the binding's project in
   `viewers/pptx/playwright.config.ts`. `e2e/pptx/building-blocks.spec.ts` then runs against it
   unchanged.
7. Docs: `viewers/pptx/docs/<binding>/building-blocks.md` (copy the vanilla page) and its sidebar
   entry.

## Adding an action group

The eight dialog groups above were added this way; a new one plugs in without editing the session:

1. Put the group's DOM-free decisions and types in a new file in `src/ui/src/pptx/render/`.
2. Add its key and type to `BuildingBlockActionGroupMap` (`action-groups.ts`) and, if it raises a
   request, its kind to `BuildingBlockRequestMap` and `BUILDING_BLOCK_REQUEST_ACTIONS`
   (`requests.ts`, plus `BUILDING_BLOCK_VIEWABLE_REQUESTS` when it only shows state). If it needs
   its own availability id, add it to `BuildingBlockActionId` and `describeBuildingBlockActions`.
3. Define it in `groups/` with `defineBuildingBlockActionGroup({ id, create(context, port) })`.
   `context` has `refuse(id)`, `request(next)`, `complete(kind, result)` and `ready()`; `port` is
   the group's own port, read at call time. `create` runs on every `actions()` call, so keep it
   cheap and keep state that spans calls (like the custom shows stored this tick) in the port.
4. Add it to `BuildingBlockGroupPorts`, `BUILDING_BLOCK_GROUP_IDS` and
   `registerBuildingBlockGroups`, and its id to the contract's `BUILDING_BLOCK_GROUP_IDS`. Every
   binding then supplies its port. Remove its handler from `BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS`.

React's ports come from small hooks (`useFileDialogPorts`, `useShowDialogPorts`) that create them
once and read the latest render through a ref; `useViewerBuildingBlocksActions` registers them once
for the session's lifetime, and `useViewerBuildingBlocks` reads `session.actions()` last.

## Parity is enforced by a shared contract test

`viewers/pptx/packages/shared/test-utils/building-blocks-contract.mjs` (typed by its `.d.mts`) runs
the same assertions against any binding through a harness: view-only refusal, edit permission,
selection-bound availability, the effective `mode`, every group present, one operation per dialog
group (password, document properties, font embedding, signatures, set up slide show, custom shows,
broadcast, shortcuts with its announced request and cancel), a toolbar request and its cancel,
validation, an insert that marks the deck dirty, Find / Replace with the host bar, layout refusal,
taking permission away again, and for compact-toolbar ports the block markers and the control order
(`COMPACT_TOOLBAR_CONTROLS`). Its shell part (`building-blocks-contract-shell.mjs`) checks the shell
state while not sharing, live `googleFonts` / `allowExternalContent` updates that keep the deck
(its edited title, custom show and dirty flag), joining a host-owned room through the options
(`contractRoom(Y)`: an `externalSession` on the binding's own `Y.Doc` with a dependency-free
awareness, connected and synced), a `viewer` role taking editing away, and leaving again with the
deck kept.

- React: `packages/react/src/viewer/hooks/useViewerBuildingBlocks.contract.test.tsx`.
- Vanilla: `packages/vanilla/src/building-blocks/create-viewer-building-blocks.test.ts`.
- Vue: `packages/vue/src/building-blocks/useViewerBuildingBlocks.test.ts`.
- Svelte: `packages/svelte/src/building-blocks/create-viewer-building-blocks.svelte.test.ts`.
- Angular: `packages/angular/src/building-blocks/viewer-building-blocks.contract.test.ts`.
- Each port adds one test that builds a harness (`mount`, `update`, `actions`, `state` with
  `shellState` and `collaborationState`, `clickPictures`, `settle`, `destroy`, `toolbar: true`) and
  calls `runBuildingBlocksContract` with its own `yjs` (`Y`). It
  must pass unchanged; a port that needs the contract changed is a design change for this note.

The browser side is `e2e/pptx/building-blocks.spec.ts`: it drives each demo's `?blocks=1` page
(the binding's own blocks plus the shared host panel) for every project with
`metadata.buildingBlocks`, all five bindings. Run it with
`PPTX_E2E_PORT_OFFSET=500 bun run e2e -- --project=react --project=vue --project=angular --project=vanilla --project=svelte building-blocks.spec.ts`
in `viewers/pptx`. `e2e/pptx/building-blocks-collaboration.spec.ts` runs collaboration through the blocks on the same
projects: two `?blocks=1` pages join one host-owned room on the local relay (the panel's
"Host join room", `demos/pptx/shared/building-blocks-room.ts`: a Y.Doc, awareness and
`y-websocket` provider the page owns, handed over as `externalSession`), see each other in
`shellState.connectedCount`, share an insert, and one leaves with its deck kept. Nothing connects
until the button is pressed. `e2e/pptx/host-owned-*.spec.ts` keep covering the custom-shell demos.

## Known gaps

- The compact toolbar has no controls for the File and Slide Show dialogs; a host raises
  those workflows from its own controls and the `actions` groups.
