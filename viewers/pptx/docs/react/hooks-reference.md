---
title: Complete Hooks Reference
description: The full list of every internal hook that composes PowerPointViewer, and how to import them via pptx-react-viewer/internals.
---

# Complete Hooks Reference

This is the complete list referenced from the [Hooks](/react/hooks) page: every hook that composes
`PowerPointViewer` internally, grouped by concern. For the small curated subset that has a
semver-stable public API, see [Hooks -> Public hooks](/react/hooks#public-hooks) instead.

::: warning Internal building blocks
All of the hooks below are importable from `pptx-react-viewer/internals`, but the name is
literal: they are not covered by semver, so signatures, behavior, and existence can change without
a major bump. Reach for this entry only when the curated `pptx-react-viewer/viewer` export and the
public `PowerPointViewer` props/handle genuinely can't do what you need. Pin an exact version if
you depend on it.

```tsx
import { useViewerState, useEditorHistory } from 'pptx-react-viewer/internals';
```

:::

## Core state & lifecycle

| Hook                     | Concern                                                         |
| ------------------------ | --------------------------------------------------------------- |
| `useViewerState`         | Composite state hook (composes core + UI state).                |
| `useViewerCoreState`     | Document state: slides, selection, canvas size, mode.           |
| `useViewerUIState`       | UI state: panel visibility, dialog flags, toolbar flags.        |
| `useDerivedSlideState`   | Computed visible indexes, sections, master pseudo-slide.        |
| `useDerivedElementState` | Computed per-element derived state (bounds, handles, z-order).  |
| `useContentLifecycle`    | Content sync, dirty tracking, recovery detection wiring.        |
| `useLoadContent`         | Parses the PPTX buffer on mount via `PptxHandler`.              |
| `useRecoveryDetection`   | Detects a prior unsaved session to offer recovery.              |
| `useAutosave`            | Periodic autosave scheduling and status.                        |
| `useSerialize`           | Serializes the current document back to bytes.                  |
| `useViewerIntegration`   | Top-level integration: I/O, export, print, pointers, lifecycle. |

## Editing & history

| Hook                         | Concern                                                                     |
| ---------------------------- | --------------------------------------------------------------------------- |
| `useEditorHistory`           | Undo/redo snapshot stack with deferred capture during pointer interactions. |
| `useEditorOperations`        | Composes all editor operations into one result.                             |
| `useElementOperations`       | Element create/update/delete primitives.                                    |
| `useElementManipulation`     | Move / resize / rotate / delete elements.                                   |
| `useSectionOperations`       | Section (slide grouping) create/rename/delete.                              |
| `useTableOperations`         | Row/column insert/delete, merge/split cells.                                |
| `useSlideManagement`         | Add / delete / duplicate / reorder / hide slides.                           |
| `useInsertElements`          | Shape / image / text box / table / chart insertion.                         |
| `useGroupAlignLayerHandlers` | Group/ungroup, align/distribute, layer (front/back) operations.             |
| `useMergeShapesHandler`      | Boolean shape merge (union/subtract/intersect/exclude).                     |
| `usePropertyHandlers`        | Property-panel change handlers for selected elements.                       |
| `useThemeHandlers`           | Applying/switching the document's PowerPoint theme.                         |
| `useThemeSwitching`          | Higher-level theme switch orchestration (also public, see below).           |
| `useLayoutSwitching`         | Switching a slide's layout while remapping placeholders.                    |
| `useClipboardHandlers`       | Copy / cut / paste of elements.                                             |
| `useFindReplace`             | Find & replace across slide text.                                           |
| `useComments`                | Comment thread state and CRUD.                                              |
| `useAnnotationHandlers`      | Freehand/shape annotation drawing handlers.                                 |

## Canvas interaction

| Hook                        | Concern                                            |
| --------------------------- | -------------------------------------------------- |
| `usePointerHandlers`        | Mouse/touch event processing for the canvas.       |
| `useCanvasInteractions`     | Selection box, drag-select, canvas-level gestures. |
| `useZoomViewport`           | Zoom level, fit-to-width, viewport DOM ref.        |
| `useKeyboardShortcuts`      | Hotkey definitions.                                |
| `useKeyboardShortcutWiring` | Wires hotkey definitions to DOM event listeners.   |
| `useResizablePanels`        | Resizable inspector/sidebar panel widths.          |
| `useSheetDismissDrag`       | Swipe-to-dismiss for bottom sheets (mobile).       |
| `useModalDismissDrag`       | Swipe-to-dismiss for modals (mobile).              |

## Export, print & I/O

| Hook                                             | Concern                                            |
| ------------------------------------------------ | -------------------------------------------------- |
| `useExportHandlers`                              | PNG / SVG / PDF / GIF / video / PPTX export logic. |
| `useExportSaveAs`                                | "Save as" file-picker flow around export.          |
| `usePrintHandlers`                               | Print dialog and layout.                           |
| `useIOHandlers`                                  | Open/import file handling.                         |
| `useFontInjection`                               | Injects embedded document fonts into the page.     |
| `useVirtualizedSlides` (+ `computeVirtualRange`) | Virtualizes the slide panel/list for large decks.  |

## Dialogs

| Hook                   | Concern                                          |
| ---------------------- | ------------------------------------------------ |
| `useViewerDialogs`     | Open/close state for every modal dialog.         |
| `useDialogCustomShows` | Custom "show" gating logic for specific dialogs. |

## Presentation mode

| Hook                         | Concern                                                            |
| ---------------------------- | ------------------------------------------------------------------ |
| `usePresentationMode`        | Slideshow navigation, animation, transitions (top-level composer). |
| `usePresentationSetup`       | Presentation-mode entry setup (fullscreen, initial slide).         |
| `usePresentationAnnotations` | Pen/highlighter annotations while presenting.                      |
| `useAnimationPlayback`       | Plays back element animation timelines.                            |
| `useRehearsalTimings`        | Records per-slide timings during a rehearsal pass.                 |
| `usePresentationKeyboard`    | Keyboard navigation while presenting.                              |
| `useSlideNavigation`         | Next/previous/jump slide navigation.                               |
| `useZoomNavigation`          | Zoom/pan while presenting.                                         |
| `usePresenterWindow`         | Presenter-window <-> audience-window messaging.                    |
| `useAudienceMode`            | Audience-window-side rendering mode.                               |

## Slide-transition helpers

The framework-agnostic logic every binding uses lives in `ooxml-ui/pptx`; the private
`pptx-viewer-shared` package is only its compatibility facade. A host embedding its own presentation stage (a custom `SlideStage`
instead of the full `PowerPointViewer`) still needs the transition resolver/keyframes, so the
whole surface is re-exported here instead:

| Export                                                                                                                                                          | Concern                                                                                                                                             |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `resolveSlideTransition`                                                                                                                                        | Resolve a `PptxSlideTransition` to CSS `animation` shorthands for the outgoing/incoming layers.                                                     |
| `resolveTransitionDurationMs`                                                                                                                                   | Effective duration (ms) for a transition, honoring an authored duration, the legacy `spd` token, and PowerPoint's own defaults.                     |
| `getSlideTransitionAnimations`                                                                                                                                  | Lower-level resolver `resolveSlideTransition` calls internally (classic 2-D transition family).                                                     |
| `getCinematicTransitionAnimations` / `getP14TransitionAnimations`                                                                                               | The Office 2013+ (p15) cinematic family and the Office 2010 (p14) exotic/3-D family, respectively.                                                  |
| `SLIDE_TRANSITION_KEYFRAMES` (alias `SLIDE_TRANSITION_KEYFRAMES_CSS`)                                                                                           | The full `@keyframes` block every resolved animation name above references. Inject once via a `<style>` tag.                                        |
| `CINEMATIC_TRANSITION_KEYFRAMES` / `P14_TRANSITION_KEYFRAMES_ALL`                                                                                               | The keyframe sub-blocks already folded into `SLIDE_TRANSITION_KEYFRAMES`; exported individually too.                                                |
| `resolveDirection` / `resolveDirection8` / `resolveOrientation`                                                                                                 | Normalize an OOXML `dir`/`orient` token to a resolved direction/orientation.                                                                        |
| `resolveWheelSpokeCount`                                                                                                                                        | Snap an authored Wheel spoke count to the nearest PowerPoint-offered value.                                                                         |
| `RANDOM_ELIGIBLE_TYPES`, `INSTANT`, `DEFAULT_TRANSITION_DURATION_MS`, `DEFAULT_MORPH_DURATION_MS`, `TRANSITION_SPEED_DURATION_MS`, `EASE`, `WHEEL_SPOKE_COUNTS` | Supporting constants used by the resolvers above.                                                                                                   |
| `PresentationTransitionOverlay`, `MorphTransitionOverlay`, `SlideLayer`, `FragmentedTransitionLayer`                                                            | The presentation-mode transition overlay components themselves, for a host that wants the whole rendered overlay rather than just the CSS resolver. |

## Collaboration

See [Collaboration](/react/collaboration) for the curated public subset. The full internal set:

| Hook                      | Concern                                                  |
| ------------------------- | -------------------------------------------------------- |
| `useYjsProvider`          | Manages the Yjs WebSocket provider lifecycle.            |
| `useYjsDocumentSync`      | Syncs the Yjs document into viewer state.                |
| `usePresenceTracking`     | Tracks remote cursors, selection, and connection status. |
| `useCollaborativeState`   | CRDT-backed shared document state.                       |
| `useCollaborativeHistory` | Collaborative undo/redo.                                 |
| `useBroadcastFollower`    | Follows another user's viewport/selection broadcast.     |
| `useFollowMode`           | "Follow presenter" mode orchestration.                   |

## Mobile & responsive

| Hook                 | Concern                                                       |
| -------------------- | ------------------------------------------------------------- |
| `useIsMobile`        | Device/viewport classification (mobile/tablet/orientation).   |
| `useTouchGestures`   | Pinch-zoom, pan, and tap-gesture recognition.                 |
| `useKeyboardInsets`  | Adjusts layout for the on-screen keyboard (mobile).           |
| `useSwipeNavigation` | Swipe-to-change-slide navigation (mobile).                    |
| `useReducedMotion`   | Respects `prefers-reduced-motion` for animations/transitions. |

This list is generated from `packages/react/src/viewer/hooks/**/*.ts` and re-exported in full from
[`pptx-react-viewer/internals`](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/react/src/internals.ts).
If you add or rename a hook there, update this page in the same change.

## useViewerBuildingBlocks: host workflows {#building-block-host-workflows}

`useViewerBuildingBlocks` is exported from the package root (`pptx-react-viewer`), not only from
`pptx-react-viewer/internals`. Its `actions` object finishes the picker, dialog and layout workflows
that the building blocks do not mount. The prose and examples are on the [Hooks](/react/hooks) page;
this section lists the exact shapes.

### Options

| Option       | Type                                      | Description                                                                                                                                                                    |
| ------------ | ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `onRequest`  | `(request: BuildingBlockRequest) => void` | Called when a toolbar control or the canvas raises a request the building blocks cannot mount. The request also stays in `actions.pending` until it is completed or cancelled. |
| `fileLimits` | `HostFileLimits`                          | Ceilings for files from the host's own picker. A larger file returns `'file-too-large'`. The defaults are 25 MiB for images and 100 MiB for audio and video.                   |

```ts
interface HostFileLimits {
	/** Largest accepted image, in bytes. Defaults to 25 MiB (26214400). */
	imageBytes?: number;
	/** Largest accepted audio or video file, in bytes. Defaults to 100 MiB (104857600). */
	mediaBytes?: number;
}
```

### Requests

```ts
type BuildingBlockRequest =
	| { kind: 'image' }
	| { kind: 'media' }
	| { kind: 'hyperlink' }
	| { kind: 'smartArt' }
	| { kind: 'chart' }
	| { kind: 'equation'; existingOmml: Record<string, unknown> | null };
```

| `kind`      | Raised by                                                                                                                 | Finish with                                                                             | Available when                     |
| ----------- | ------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- | ---------------------------------- |
| `image`     | The Picture control                                                                                                       | `actions.insertImage(file)`                                                             | `insertImage`                      |
| `media`     | The Video control                                                                                                         | `actions.insertMedia(file)`                                                             | `insertMedia`                      |
| `hyperlink` | The Link control and its canvas gesture                                                                                   | `actions.hyperlink.apply(data)`                                                         | `hyperlink`                        |
| `smartArt`  | The SmartArt control                                                                                                      | `actions.smartArt.insert(layout, items)`                                                | `insertSmartArt`                   |
| `chart`     | The Chart control                                                                                                         | `actions.chart.insert(kind, options?)`                                                  | `insertChart`                      |
| `equation`  | The Equation control, or a double-click on an equation. `existingOmml` is the equation being edited, `null` for a new one | `actions.equation.insert(omml)`, or `actions.equation.update(omml)` for an existing one | `insertEquation` or `editEquation` |

A request is raised only when its action is available. A successful matching operation clears
`actions.pending`; a failed one leaves it set. `actions.cancel()` clears it without changing the deck.

### Operations

Operations return `BuildingBlockResult` (see [Results and failure codes](#building-block-results))
unless the signature says otherwise.

| Member                   | Signature                                                                      | Availability id              |
| ------------------------ | ------------------------------------------------------------------------------ | ---------------------------- |
| `pending`                | `BuildingBlockRequest \| null`                                                 | none                         |
| `cancel`                 | `() => void`                                                                   | none                         |
| `insertImage`            | `(file: File) => Promise<BuildingBlockResult>`                                 | `insertImage`                |
| `insertMedia`            | `(file: File) => Promise<BuildingBlockResult>`                                 | `insertMedia`                |
| `hyperlink.current`      | `HyperlinkEditData \| null`                                                    | none (reads the selection)   |
| `hyperlink.slideCount`   | `number`                                                                       | none                         |
| `hyperlink.apply`        | `(data: HyperlinkEditData) => BuildingBlockResult`                             | `hyperlink`                  |
| `hyperlink.remove`       | `() => BuildingBlockResult`                                                    | `removeHyperlink`            |
| `smartArt.insert`        | `(layout: SmartArtLayout, items: string[]) => BuildingBlockResult`             | `insertSmartArt`             |
| `chart.insert`           | `(kind: InsertChartKind, options?: InsertChartOptions) => BuildingBlockResult` | `insertChart`                |
| `equation.current`       | `Record<string, unknown> \| null`                                              | none (reads the selection)   |
| `equation.insert`        | `(omml: Record<string, unknown>) => BuildingBlockResult`                       | `insertEquation`             |
| `equation.update`        | `(omml: Record<string, unknown>) => BuildingBlockResult`                       | `editEquation`               |
| `findReplace`            | `BuildingBlockFindReplace` (below)                                             | `replace` (replace only)     |
| `layout`                 | `BuildingBlockLayoutActions` (below)                                           | `applyLayout`, `resetLayout` |
| `availability`           | `Record<BuildingBlockActionId, BuildingBlockActionState>`                      | none                         |
| `unmountedDialogActions` | `readonly string[]`                                                            | none                         |

`hyperlink.slideCount` is the number of slides a slide-target picker can link to. `hyperlink.current`
is the selected element's link as editor values. `equation.current` is the selected element's equation
(OMML), or `null`. `insertImage` centres the picture on the slide, fitted to it.

`unmountedDialogActions` lists the `toolbarProps` handlers whose dialog the building blocks do not
mount: `onOpenPasswordProtection`, `onOpenDocumentProperties`, `onOpenFontEmbedding`,
`onOpenDigitalSignatures`, `onOpenSetUpSlideShow`, `onOpenCustomShows`, `onOpenBroadcastDialog` and
`onToggleShortcuts`. Hide them with `hiddenActions` or render the dialog yourself.

#### Find and replace

`actions.findReplace` exposes the viewer's find state and operations. Searching works in a view-only
viewer; replacing is gated by the `replace` availability id.

| Member                          | Type                                                                   | Notes                                                                                         |
| ------------------------------- | ---------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| `open`, `setOpen`               | `boolean`, `(open: boolean) => void`                                   | Whether the find bar is open. Ctrl+F and Ctrl+H toggle it. The host renders the bar.          |
| `query`, `setQuery`             | `string`, `(query: string) => void`                                    |                                                                                               |
| `replacement`, `setReplacement` | `string`, `(replacement: string) => void`                              |                                                                                               |
| `options`, `setOptions`         | `FindReplaceOptions`, `(options: Partial<FindReplaceOptions>) => void` | `matchCase`, `includeNotes`, `includeChartTitles` and `includeChartData`, each a `boolean`.   |
| `matches`                       | `readonly FindResult[]`                                                |                                                                                               |
| `currentIndex`                  | `number`                                                               | Zero-based. `-1` when there is no current match.                                              |
| `currentMatch`                  | `FindResult \| null`                                                   |                                                                                               |
| `find`                          | `() => void`                                                           | Searches again with the current query and options.                                            |
| `next`, `previous`              | `() => void`                                                           | Shows the match's slide and selects its element.                                              |
| `replaceCurrent`, `replaceAll`  | `() => BuildingBlockResult`                                            | Edit permission, validation and history apply. `'no-match'` when there is nothing to replace. |

#### Layouts

`actions.layout` uses the active slide's master's layouts as the allowed set. The Layout dropdown and
Reset in `toolbarProps` run the same operations.

| Member        | Type                                             | Notes                                                                     |
| ------------- | ------------------------------------------------ | ------------------------------------------------------------------------- |
| `current`     | `PptxLayoutOption \| undefined`                  | The layout the active slide records, when the deck lists it.              |
| `currentPath` | `string \| undefined`                            |                                                                           |
| `available`   | `PptxLayoutOption[]`                             | The layouts the active slide may switch to.                               |
| `apply`       | `(path: string) => Promise<BuildingBlockResult>` | `'unknown-layout'` for a path not in `available`, or when applying fails. |
| `reset`       | `() => Promise<BuildingBlockResult>`             | Re-applies the slide's own layout. `'no-layout'` when it records none.    |

### Charts

```ts
type InsertChartKind =
	| 'column'
	| 'bar'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'histogram'
	| 'pareto'
	| 'funnel'
	| 'treemap'
	| 'sunburst'
	| 'boxWhisker'
	| 'regionMap';

type InsertChartGrouping = 'clustered' | 'stacked' | 'percentStacked';

interface InsertChartOptions {
	/** Default sample title when undefined; no title when empty. */
	title?: string;
	/** Only for the `column`, `bar`, `line` and `area` kinds. */
	grouping?: InsertChartGrouping;
	/** Defaults to true. */
	hasLegend?: boolean;
	/** Host data in place of the sample data. */
	data?: InsertChartData;
}

interface InsertChartData {
	categories: readonly string[];
	series: ReadonlyArray<{ name: string; values: readonly number[] }>;
}
```

`data` needs at least one category and one series, and every series needs one finite value per
category. Otherwise `chart.insert` returns `'invalid-data'`, after the availability check.

### Availability

```ts
type BuildingBlockActionId =
	| 'insertImage'
	| 'insertMedia'
	| 'hyperlink'
	| 'removeHyperlink'
	| 'insertSmartArt'
	| 'insertChart'
	| 'insertEquation'
	| 'editEquation'
	| 'replace'
	| 'applyLayout'
	| 'resetLayout';

interface BuildingBlockActionState {
	/** The operation would run now. When false, `reason` says why. */
	available: boolean;
	reason?: BuildingBlockFailureCode;
	/** The toolbar control only raises a request; the host must render the picker or dialog. */
	needsHostUi: boolean;
}
```

Every entry checks the shared conditions first: `'not-ready'` (loading or failed), `'read-only'` (not
in edit mode, or no edit permission) and `'protected'` (read-only recommended or Protected View). The
action's own conditions follow. The first failing condition is the `reason`.

| Id                | `needsHostUi` | Own conditions, in order                                             |
| ----------------- | ------------- | -------------------------------------------------------------------- |
| `insertImage`     | true          | `'no-slide'`                                                         |
| `insertMedia`     | true          | `'no-slide'`                                                         |
| `hyperlink`       | true          | `'no-selection'`                                                     |
| `removeHyperlink` | false         | `'no-selection'`, then `'no-hyperlink'`                              |
| `insertSmartArt`  | true          | `'no-slide'`                                                         |
| `insertChart`     | true          | `'no-slide'`                                                         |
| `insertEquation`  | true          | `'no-slide'`                                                         |
| `editEquation`    | true          | `'no-selection'`, then `'no-equation'`                               |
| `replace`         | false         | none                                                                 |
| `applyLayout`     | false         | `'no-slide'`, then `'no-layout'` (no handler, or no layouts)         |
| `resetLayout`     | false         | `'no-slide'`, then `'no-layout'` (no handler, or no recorded layout) |

### Results and failure codes {#building-block-results}

```ts
type BuildingBlockResult = { ok: true } | { ok: false; code: BuildingBlockFailureCode };
```

Codes are not English text, so render them in your own language. An operation reports its
availability reason first, then checks its arguments. A failed operation leaves the deck unchanged.

| `code`             | Meaning                                                                                                        | Returned by                                                                                                        |
| ------------------ | -------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `not-ready`        | The deck is still loading or failed to load.                                                                   | Every operation                                                                                                    |
| `read-only`        | Not editable: view-only, not in edit mode, or no edit permission.                                              | Every operation                                                                                                    |
| `protected`        | Read-only recommended or Protected View, so inserting is blocked.                                              | Every operation                                                                                                    |
| `no-slide`         | There is no active slide.                                                                                      | `insertImage`, `insertMedia`, `smartArt.insert`, `chart.insert`, `equation.insert`, `layout.apply`, `layout.reset` |
| `no-selection`     | Nothing is selected.                                                                                           | `hyperlink.apply`, `hyperlink.remove`, `equation.update`                                                           |
| `no-file`          | No file was given.                                                                                             | `insertImage`, `insertMedia`                                                                                       |
| `empty-file`       | The file has no content.                                                                                       | `insertImage`, `insertMedia`                                                                                       |
| `unsupported-type` | Not an image (for `insertImage`), or not an audio or video type (for `insertMedia`).                           | `insertImage`, `insertMedia`                                                                                       |
| `file-too-large`   | The file is larger than `fileLimits` allows.                                                                   | `insertImage`, `insertMedia`                                                                                       |
| `decode-failed`    | The image or media file could not be read or decoded.                                                          | `insertImage`, `insertMedia`                                                                                       |
| `stale-target`     | The deck, the active slide or its position changed while the file was decoded or probed. Nothing was inserted. | `insertImage`, `insertMedia`                                                                                       |
| `unsafe-url`       | The link uses a script-like scheme (`javascript:`, `data:`, `vbscript:` or `mhtml:`).                          | `hyperlink.apply`                                                                                                  |
| `invalid-target`   | The link has neither a URL nor an action.                                                                      | `hyperlink.apply`                                                                                                  |
| `no-hyperlink`     | The selected element has no link.                                                                              | `hyperlink.remove`                                                                                                 |
| `no-equation`      | The selected element has no equation.                                                                          | `equation.update`                                                                                                  |
| `no-match`         | There is no current match (`replaceCurrent`), or no match at all (`replaceAll`).                               | `findReplace.replaceCurrent`, `findReplace.replaceAll`                                                             |
| `unknown-layout`   | The path is not in `layout.available`, or applying the layout failed.                                          | `layout.apply`                                                                                                     |
| `invalid-data`     | The `data` given to `chart.insert` does not meet the rules in [Charts](#charts).                               | `chart.insert`                                                                                                     |
| `no-layout`        | The slide records no layout (Reset), or there is no handler or no layouts (apply).                             | `layout.apply`, `layout.reset`                                                                                     |
