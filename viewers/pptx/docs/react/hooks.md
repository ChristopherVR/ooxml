---
title: Hooks
description: The hooks-based architecture of PowerPointViewer, the curated set of public, tree-shakeable hooks exported from pptx-react-viewer/viewer, and the full internal set exposed via pptx-react-viewer/internals.
---

# Hooks

`PowerPointViewer` is a thin `forwardRef` orchestrator. Almost all of its logic lives in **80+ custom
hooks** composed inside `PowerPointViewer.tsx`, while the visual components are largely
presentational. State is held entirely in React hooks; there is no external state library.

::: info Public vs internal
Most of these hooks are **internal architecture**: they assume a specific composition order and
shared inputs. A small curated subset is exported from `pptx-react-viewer/viewer` with a normal
semver-stable API. The **complete** set is also importable from `pptx-react-viewer/internals`:
internal building blocks that are not covered by semver, so prefer the stable root exports. See
[Complete Hooks Reference](/react/hooks-reference).
:::

## Architecture (internal)

These hooks describe how the viewer is wired. They are importable (see below) but assume a
specific composition order and shared inputs; treat this table as conceptual reference, not an
API contract.

| Hook                     | Concern                                                                     |
| ------------------------ | --------------------------------------------------------------------------- |
| `useViewerState`         | Composite state hook (composes core + UI state).                            |
| `useViewerCoreState`     | Document state: slides, selection, canvas size, mode.                       |
| `useViewerUIState`       | UI state: panel visibility, dialog flags, toolbar flags.                    |
| `useDerivedSlideState`   | Computed visible indexes, sections, master pseudo-slide.                    |
| `useEditorHistory`       | Undo/redo snapshot stack with deferred capture during pointer interactions. |
| `useZoomViewport`        | Zoom level, fit-to-width, viewport DOM ref.                                 |
| `useEditorOperations`    | Composes all editor operations into one result.                             |
| `useLoadContent`         | Parses the PPTX buffer on mount via `PptxHandler`.                          |
| `useContentLifecycle`    | Content sync, dirty tracking, recovery detection.                           |
| `usePresentationMode`    | Slideshow navigation, animation, transitions.                               |
| `useExportHandlers`      | PNG / SVG / PDF / GIF / video / PPTX export logic.                          |
| `usePrintHandlers`       | Print dialog and layout.                                                    |
| `useInsertElements`      | Shape / image / text box / table / chart insertion.                         |
| `useElementManipulation` | Move / resize / rotate / delete elements.                                   |
| `useSlideManagement`     | Add / delete / duplicate / reorder / hide slides.                           |
| `useTableOperations`     | Row/column insert/delete, merge/split cells.                                |
| `usePointerHandlers`     | Mouse/touch event processing for the canvas.                                |
| `useKeyboardShortcuts`   | Hotkey definitions.                                                         |
| `useViewerIntegration`   | Top-level integration: I/O, export, print, pointers, lifecycle.             |

There are dozens more (clipboard, comments, sections, autosave, font injection, recovery, theme
handlers, presentation sub-hooks, etc.). See the **[Complete Hooks Reference](/react/hooks-reference)**
for the full list, grouped by concern.

Internal hooks are wiring-heavy by design. `useEditorHistory`, for example, takes the full editor
state plus a setter for every state slice (8 setters) and returns
`{ canUndo, canRedo, undoLabel, redoLabel, handleUndo, handleRedo, resetHistory, markDirty, buildHistorySnapshot }`.
That shape only makes sense inside the component's composition; for standalone undo/redo, mutate
`PptxData` yourself and snapshot it (see [Editing Programmatically](/core/editing)).

## Public hooks

The following hooks are exported from `pptx-react-viewer/viewer` and are opt-in and
tree-shakeable. `useViewerBuildingBlocks` is also exported from the package root,
alongside the standalone `Toolbar`, `SlideCanvas` and `SlidesPane` components.

```tsx
import { useThemeSwitching, useCollaborativeState } from 'pptx-react-viewer/viewer';
```

### Element render errors

The React canvas, thumbnails and static preview/transition renderers isolate client
render and lifecycle errors per element. A failed element renders nothing; neighboring
content and editor chrome remain mounted. This does not modify or remove the element
from the document. Missing 3D dependencies and pending scenes retain their poster
fallback; a 3D scene render error is reported and omitted like any other failed element.

Pass an optional `onElementRenderError` to `useViewerBuildingBlocks` to report failures
from both returned `canvasProps` and `slidesPaneProps`:

```tsx
import { SlideCanvas, SlidesPane, useViewerBuildingBlocks } from 'pptx-react-viewer';
import type { ElementRenderErrorHandler } from 'pptx-react-viewer';

const reportElementError: ElementRenderErrorHandler = (error, context) => {
	console.error('Element could not be rendered', error, context);
};

function Editor({ content }) {
	const blocks = useViewerBuildingBlocks({ content, onElementRenderError: reportElementError });
	return (
		<>
			<SlidesPane {...blocks.slidesPaneProps} />
			<SlideCanvas {...blocks.canvasProps} />
		</>
	);
}
```

The exported `ElementRenderErrorContext` contains `elementId`, `elementType` and an
optional `slideId` (when an owning slide is supplied). Without a callback, the viewer
calls `console.error(error, context)`. A throwing callback falls back to that logger
without taking down the viewer. React's development build may also log the caught error.
An element rendered in multiple surfaces can report once per failed mounted instance.

Standalone `SlideCanvas`, `SlidesPane`, `SlideThumbnail`, `ScaledSlidePreview`,
`SlideTemplatePreview`, `PresentationStage`, `ReadingViewStage` and
`PresentationTransitionOverlay` also accept `onElementRenderError`. This is not a new
`PowerPointViewer` prop. Replacing a failed element object, even with the same id, retries
its rendering; changing its id remounts its boundary. Rerendering the same failed object
does not retry, so in-place mutation alone does not clear the failure. These boundaries
do not catch event-handler errors, arbitrary asynchronous/timer errors, or server-render
errors. They are React-specific; other bindings do not gain this callback or guarantee.

### Custom editor collaboration

`useViewerBuildingBlocks` accepts the same optional `collaboration` configuration
as `PowerPointViewer`, including a [host-owned Yjs session](/guide/host-owned-collaboration).
It returns `collaboration` connection/presence state for a custom status bar,
the same `shellState` and `collaborationState` (the shared `CollaborationState`) every binding's
building blocks report, and includes cursor and remote-selection overlays in `canvasProps`.
Changing `collaboration` joins, rejoins or leaves without remounting the canvas. No extra collaboration provider is needed.

```tsx
import { SlideCanvas, Toolbar, useViewerBuildingBlocks } from 'pptx-react-viewer';

function Editor({ content, collaboration }) {
	const blocks = useViewerBuildingBlocks({ content, collaboration, canEdit: true });
	return (
		<>
			<Toolbar {...blocks.toolbarProps} />
			<SlideCanvas {...blocks.canvasProps} />
		</>
	);
}
```

The hook uses the bundled editor's document sync, readiness gates and write-back
path. `role: 'viewer'` disables editing. This does not change its existing undo
semantics or add separate host comments integration.

The additional source-loading and error gate applies only when collaboration is
configured. Without it, the hook preserves existing host authorization, including
`content: null` for a blank document. An active collaborative session also respects
the host's edit permission and readiness; a connected transport alone does not grant
permission to edit. `CollaborationContextValue` is available from the package root
when the host needs to name the returned state type.

### Slide thumbnail rail

`slidesPaneProps` drives `SlidesPane`, the rail of slide thumbnails that `PowerPointViewer`
shows beside the canvas. It lists every slide (grouped by section when the deck has sections)
and highlights the active one; clicking a thumbnail moves the canvas there without marking the
deck dirty. With `canEdit`, it also adds, duplicates, deletes, hides, reorders and groups slides
through the same operations, undo history and collaboration sync as the full editor.

```tsx
import { SlideCanvas, SlidesPane, Toolbar, useViewerBuildingBlocks } from 'pptx-react-viewer';

function Editor({ content }) {
	const blocks = useViewerBuildingBlocks({ content, canEdit: true });
	return (
		<>
			<Toolbar {...blocks.toolbarProps} />
			<div style={{ display: 'flex', height: 600 }}>
				<SlidesPane {...blocks.slidesPaneProps} />
				<div style={{ flex: 1, minWidth: 0, position: 'relative' }}>
					<SlideCanvas {...blocks.canvasProps} />
				</div>
			</div>
		</>
	);
}
```

The rail's open state is the toolbar's sidebar toggle: collapsing either hides the rail. The
"Layout" entry of a slide's context menu makes that slide active and calls the optional
`onOpenLayoutGallery({ slideIndex, x, y })`, since these building blocks do not render the Layout
gallery. `SlidesPane` needs the full `styles.css`, not `canvas.css`.

### Single slide thumbnail

`SlideThumbnail` (exported, with its `SlideThumbnailProps` type, from the package root) draws one
slide the way `SlidesPane` draws it, for a host that keeps its own thumbnail strip. It renders from
the slide model, so it follows unsaved edits, and it is static: no selection, editing, pointer
handlers or element ids. A slide with a transition other than Cut shows the transition's name in a
small badge in its top-right corner, as the thumbnail does in every binding. Its props are `slide`,
`templateElements` (the slide's layout and master elements, drawn beneath its own), `canvasSize`,
`previewWidth` (pixels; the height follows the deck's aspect ratio), `fieldContext` and
`tableStyleContext`. Feed them from `slidesPaneProps`:

```tsx
import { SlideThumbnail, useViewerBuildingBlocks } from 'pptx-react-viewer';

function MyStrip({ blocks }: { blocks: ReturnType<typeof useViewerBuildingBlocks> }) {
	const { slides, templateElementsBySlideId, canvasSize, fieldContext, tableStyleContext } =
		blocks.slidesPaneProps;
	return (
		<ul className='my-strip'>
			{slides.map((slide) => (
				<li key={slide.id} style={{ width: 120 }}>
					<SlideThumbnail
						slide={slide}
						templateElements={templateElementsBySlideId[slide.id] ?? []}
						canvasSize={canvasSize}
						previewWidth={120}
						fieldContext={fieldContext}
						tableStyleContext={tableStyleContext}
					/>
				</li>
			))}
		</ul>
	);
}
```

Without `previewWidth`, the thumbnail is the rail's width and gets the rail's border. It redraws
when it receives a different `slide` object, as every edit made through the viewer produces; a host
that edits slides itself should replace the slide object rather than mutate it.

### Opening on a slide

`useViewerBuildingBlocks` takes the same `initialSlide` option as `PowerPointViewer`: the
zero-based slide shown after each load, clamped into the deck. It is applied in the render that
first shows the loaded deck, so a custom shell never shows slide 1 first and does not need to
call `goTo()` after load or hide the canvas until it lands.

```tsx
const blocks = useViewerBuildingBlocks({ content, initialSlide: 4 });
```

Changing `initialSlide` later applies to the next load, not the open deck; use the handle's
`goTo(index)` to move the current slide.

### Picker, dialog and layout workflows from your own controls

The building blocks mount no file inputs, dialogs or panes. `blocks.actions` is how a custom shell
finishes those workflows from host-rendered controls. Every operation runs the viewer's own insert,
find / replace and layout code, so selection, history (undo) and `onDirtyChange` behave as in
`PowerPointViewer`. Every operation returns `{ ok: true }` or `{ ok: false, code }` and is refused
with `'read-only'` (view-only, not in edit mode, no edit permission), `'protected'` (read-only
recommended or Protected View) or `'not-ready'` (loading or failed). The codes are not English
text: render the feedback in your own language.

The `actions` object is the shared building-blocks session (`ooxml-ui/pptx`), so it has the same
shape in every binding; the vanilla [`createViewerBuildingBlocks`](/vanilla/building-blocks)
returns the same `actions`.

```tsx
const blocks = useViewerBuildingBlocks({
	content,
	canEdit: true,
	// Called when a toolbar control or a canvas gesture needs UI you own.
	onRequest: (request) => openMyDialog(request.kind),
	fileLimits: { imageBytes: 10_000_000 }, // defaults: 25 MiB image, 100 MiB audio/video
});
const { actions } = blocks;
```

- **Requests.** The toolbar's Picture, Video, Link, SmartArt, Chart and Equation controls, and
  double-clicking an equation on the canvas, raise a `BuildingBlockRequest`
  (`{ kind: 'image' | 'media' | 'hyperlink' | 'smartArt' | 'chart' | 'passwordProtection' |
  'documentProperties' | 'fontEmbedding' | 'digitalSignatures' }` or `{ kind: 'equation', existingOmml }`). It is passed to `onRequest` and kept in
  `actions.pending` until a matching operation succeeds or you call `actions.cancel()`. A request is
  not raised when the action is unavailable.
- **Images and media.** `actions.insertImage(file)` and `actions.insertMedia(file)` take the `File`
  your own picker returned. They validate type, emptiness and size (`'unsupported-type'`,
  `'empty-file'`, `'file-too-large'`, `'decode-failed'`), insert the element centred on the slide
  and select it. Both return `'stale-target'` and add nothing if the deck, the active slide or its
  position changed while the picture decoded or the video's size was probed, so an element is never
  dropped onto the wrong slide.
- **Hyperlinks.** `actions.hyperlink.current` is the selected element's link as editor values (or
  `null`), `slideCount` bounds a slide-target picker, `apply(data)` sets it (`'unsafe-url'` for
  script-like schemes, `'invalid-target'` for an empty target) and `remove()` clears it.
- **Charts.** The Insert Chart control raises `{ kind: 'chart' }`. Finish it with
  `actions.chart.insert(kind, options?)`, where `kind` is an Insert Chart entry (`'column'`, `'bar'`,
  `'line'`, `'pie'`, ...) and `options` takes `title`, `grouping`, `hasLegend` and `data`
  (`{ categories, series: [{ name, values }] }`, one value per category). Without `data` the usual
  sample data is used; data that does not fit returns `'invalid-data'`.
- **SmartArt and equations.** `actions.smartArt.insert(layout, items)`;
  `actions.equation.insert(omml)`, `update(omml)` for the selected element and `current` for its
  OMML.
- **Find / Replace.** `actions.findReplace` has `query`, `replacement`, `options` (`matchCase`,
  `includeNotes`, `includeChartTitles`, `includeChartData`), `matches`, `currentIndex`,
  `currentMatch`, `setQuery`, `setReplacement`, `setOptions`, `find`, `next`, `previous`,
  `replaceCurrent()` and `replaceAll()` (`'no-match'` when there is nothing to replace), plus `open`
  and `setOpen` for the bar (Ctrl+F and Ctrl+H toggle it). Searching works in a view-only viewer;
  replacing does not. `next` and `previous` show the match's slide and select its element.
- **Layouts.** `toolbarProps` now wires Home > Layout and Reset. `actions.layout` has `current`,
  `currentPath`, `available` (the active slide's master's layouts), `apply(path)`
  (`'unknown-layout'` for a path that is not listed) and `reset()` (`'no-layout'` when the slide
  records none).
- **File dialogs.** File > Info > Protect, Properties, Fonts and Signatures raise
  `{ kind: 'passwordProtection' | 'documentProperties' | 'fontEmbedding' | 'digitalSignatures' }`.
  The request is also raised while the viewer is read-only or in Protected View, so you can show the
  state; `availability[id]` says whether applying is allowed, and every operation checks that again
  when it runs. Changes mark the deck dirty and are saved with it.
  - `actions.passwordProtection`: `state` is `{ openPassword, modifyPassword, readOnlyRecommended,
    markedFinal, supported }` and never contains a password or hash. `setOpenPassword(password,
    confirmation)` (`'password-required'`, `'password-mismatch'`, `'password-too-short'`) encrypts
    the next save; `clearOpenPassword()`; `setMarkedFinal(boolean)` writes or removes `_MarkAsFinal`.
    `setModifyPassword()` always returns `'unsupported'`: the viewer preserves a modify password and
    the read-only recommendation the deck arrived with and cannot write them.
  - `actions.documentProperties`: `current` is `{ core, custom, app }` (copies). `apply({ core?, app?,
    custom?: { set?, remove? } })` patches the Summary fields (title, subject, creator, keywords,
    description, category), company and manager, and custom properties by name. A malformed patch
    (blank name, unknown type, a value that does not parse as its type) returns `'invalid-data'` and
    changes nothing.
  - `actions.fontEmbedding`: `state` is `{ embedFonts, embeddedFamilies, usedFamilies, canEmbed,
    unavailableReasonKey? }`; `setEmbedFonts(enabled)` keeps or strips the deck's embedded fonts on
    save. Turning it on for a deck that embeds nothing returns `'unsupported'`: a browser cannot read
    the bytes of an installed font, so the viewer cannot add one.
  - `actions.digitalSignatures`: `state` is `{ signed, count, removedOnSave, supported }`. `list()`
    reads the signatures from the loaded package (signer, issuer, validity dates, `expired`, and
    `integrity`: `'digestsMatch' | 'invalid' | 'unverified'` for the parts each signature covers).
    That is an integrity check only: the signature value, certificate chain and revocation are not
    evaluated, so nothing is reported as trusted. `remove()` marks the deck changed so the saved file
    carries no signatures (saving always drops them; `'no-signatures'` when there are none).
    `sign()` always returns `'unsupported'`: the viewer cannot create a signature.
- **Slide show dialogs.** Slide Show > Set Up Slide Show, Custom Slide Show, Broadcast and the
  shortcut help (`?` or Ctrl+/) are host-owned too. The first two raise `{ kind: 'setUpSlideShow' }`
  and `{ kind: 'customShows' }`; the shortcut help raises `{ kind: 'shortcuts', open }` every time it
  opens or closes (and keeps it in `actions.pending` while open).
  - `actions.setUpSlideShow.current` has the settings with defaults resolved (`showType`, `loop`,
    `playNarration`, `playAnimation`, `subtitles`, `advanceMode`, `slidesMode`, `rangeFrom`,
    `rangeTo`, `customShowId`, `penColor`, `kioskRestartTime`, `showScrollbar`, `slideCount`,
    `customShows`). `apply(patch)` takes the changed fields as the deck stores them
    (`showType`, `loopContinuously`, `showWithNarration`, `showWithAnimation`, `showSubtitles`,
    `advanceMode`, `showSlidesMode`, `showSlidesFrom`, `showSlidesTo`, `showSlidesCustomShowId`,
    `penColor`, `kioskRestartTime`, `showScrollbar`) and applies the bundled dialog's rules (a kiosk
    show loops, a range defaults to the whole deck). It returns `'invalid-range'`,
    `'unknown-show'` or `'invalid-data'` for a patch the deck cannot hold. The monitor and a laser
    colour are not modelled by the deck, so they are not offered.
  - `actions.customShows` has `shows` (`id`, `name`, `slideRIds`, `slideCount`), `slides` (the deck's
    `rId` and number), `activeId`, and `create(name, slideRIds?)`, `rename(id, name)`,
    `setSlides(id, slideRIds)`, `copy(id, name?)`, `remove(id)` (jumps into the show are cleared) and
    `start(id)`, which selects the show and starts the slide show from its first slide. Each edit is
    stored at once. Codes: `'unknown-show'`, `'unknown-slide'`, `'invalid-name'`, `'no-slide'` (a show
    with no slide to play).
  - `actions.shortcuts` has `open`, `setOpen`, `toggle` and `catalogue`: the keymap as data, one entry
    per shortcut with `id`, `descriptionKey`, the translated `description`, `keys`, `alternatives`,
    `group` and the translated `groupLabel`.
  - **Broadcast is not supported.** It streams the presenter to viewers through a collaboration
    service the building blocks do not run, so `availability.broadcast` is
    `{ available: false, reason: 'unsupported', needsHostUi: true }`, `actions.broadcast` is
    `{ supported: false, reason: 'unsupported', requires: 'collaboration-service' }`, and the
    toolbar's Broadcast control raises no request. A host with its own service passes `collaboration`
    (role `owner`) and presents with its own controls; `onOpenBroadcastDialog` stays in
    `unmountedDialogActions`.
  - The edits need the same edit permission as the other workflows (`'read-only'`, `'protected'`,
    `'not-ready'`, checked again when they run), and mark the deck dirty and update history like the
    bundled dialogs. Reading the settings, the shows and the catalogue always works.
- **Which actions can run.** `actions.availability[id]` is `{ available, reason?, needsHostUi }` for
  `insertImage`, `insertMedia`, `hyperlink`, `removeHyperlink`, `insertSmartArt`, `insertChart`,
  `insertEquation`, `editEquation`, `replace`, `applyLayout`, `resetLayout`, the file dialogs
  `passwordProtection`, `documentProperties`, `fontEmbedding` and `digitalSignatures`, and the slide
  show dialogs `setUpSlideShow`, `customShows`, `broadcast` and `shortcuts`. `needsHostUi` is true for
  the controls that only raise a request. `actions.unmountedDialogActions` lists the `toolbarProps`
  handler that still sets state for a dialog nothing renders here (`onOpenBroadcastDialog`);
  hide it with `hiddenActions` or show your own.

The exact options, request kinds, operation signatures, availability ids and failure codes are in
[`useViewerBuildingBlocks: host workflows`](/react/hooks-reference#building-block-host-workflows).

### `useThemeSwitching`

Switches the loaded **document's** PowerPoint theme (the OOXML color/font scheme, not the viewer
chrome theme; see [Theming](/guide/theming) for that distinction). Works against the same
`PptxHandler` + `PptxData` pair you get from a manual load.

```ts
interface UseThemeSwitchingInput {
	handlerRef: RefObject<PptxHandler | null>;
	data: PptxData | null;
	onDataChange: (newData: PptxData) => void;
	onThemeChanged?: (preset: PptxThemePreset) => void;
}

interface ThemeSwitchingResult {
	presets: readonly PptxThemePreset[]; // built-in presets (office, facet, ion, ...)
	switchToPreset: (preset: PptxThemePreset) => Promise<void>;
	switchToCustom: (
		colorScheme: PptxThemeColorScheme,
		fontScheme?: PptxThemeFontScheme,
		themeName?: string,
	) => Promise<void>;
	currentPreset: PptxThemePreset | undefined; // preset matching the current theme, if any
}
```

`switchToPreset` updates both the in-memory ZIP (so the change survives `save()`) and the parsed
data's resolved element colors.

```tsx
function ThemePicker({ handlerRef, data, setData }: Props) {
	const { presets, switchToPreset, currentPreset } = useThemeSwitching({
		handlerRef,
		data,
		onDataChange: setData,
	});

	return (
		<div>
			{presets.map((preset) => (
				<button
					key={preset.id}
					onClick={() => switchToPreset(preset)}
					aria-pressed={preset.id === currentPreset?.id}
				>
					{preset.name}
				</button>
			))}
		</div>
	);
}
```

### Collaboration hooks

For building custom collaboration UIs or driving sync yourself. See
[Collaboration](/react/collaboration). They require the `yjs` / `y-websocket` optional peers
(loaded dynamically, so they tree-shake away when unused). All of them take the same
`CollaborationConfig` the component's `collaboration` prop accepts: `roomId`, `serverUrl`,
`userName`, plus optional `transport` (`'websocket' | 'webrtc'`), `signaling`, `userColor`,
`userAvatar`, `authToken`, `role`, `sessionIntent`, and the elected-writer `onWriteBack` /
`writeBackDebounceMs` pair.

| Hook                      | Signature (input => result)                                                                                                                |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `useYjsProvider`          | `{ config?: CollaborationConfig }` => `{ status, awareness, doc, clientId, synced, retry }`                                                |
| `usePresenceTracking`     | `{ awareness, localClientId, userName, userColor, userAvatar?, role?, canvasWidth, canvasHeight }` => `{ remoteUsers, broadcastPresence }` |
| `useCollaborativeState`   | `{ config?, canvasWidth, canvasHeight }` => `CollaborationContextValue \| null`                                                            |
| `useCollaborativeHistory` | `{ localClientId, handleUndo, handleRedo, canUndo, canRedo }` => same four, wrapped for local-only undo scoping                            |

`useCollaborativeState` is the composition root the built-in `CollaborationProvider` uses: it
manages the transport (`useYjsProvider`) and presence (`usePresenceTracking`) and returns `null`
while `config` is `undefined` (the hooks stay dormant so your tree shape is stable).

```tsx
import { useCollaborativeState } from 'pptx-react-viewer/viewer';

function PresenceBar({ roomId, userName }: { roomId: string; userName: string }) {
	const collab = useCollaborativeState({
		config: { roomId, serverUrl: 'wss://collab.example.com', userName },
		canvasWidth: 960,
		canvasHeight: 540,
	});

	if (!collab) return null;
	return (
		<span>{collab.status === 'connected' ? `${collab.connectedCount} online` : collab.status}</span>
	);
}
```

`useYjsProvider` on its own is the thin transport layer: it lazily imports the Yjs packages,
creates the `Y.Doc` and provider, times out to `status: 'error'` when the connection fails
(recover with `retry()`), and exposes `synced` so late joiners can gate local writes until the
room's document has arrived.

The `CollaborationProvider` component and presence UI (`RemoteUserCursors`, `UserAvatarBar`,
`CollaborationStatusIndicator`) are exported alongside these.

### Audience-window helpers

Not hooks, but exported from `pptx-react-viewer/viewer` for the presenter/audience-window flow:
`isAudienceTab`, `loadAudienceContent`, `storeAudienceContent`, `clearAudienceContent`, and
`parseAudienceNonce`.

## Using an internal hook directly

If the curated public hooks above don't cover what you need, every internal hook is also
importable in full from `pptx-react-viewer/internals`:

```tsx
import { useEditorHistory, useViewerState } from 'pptx-react-viewer/internals';
```

::: warning Internal building blocks
`pptx-react-viewer/internals` re-exports the same hooks `PowerPointViewer` composes
internally, unmodified. They are **not covered by semver**: signatures and behavior can change,
and hooks can be renamed or removed, without a major bump. Prefer the props/handle API or the
curated `pptx-react-viewer/viewer` hooks first; reach for `internals` only for advanced
integrations, and pin an exact version if you depend on it.
:::

See the **[Complete Hooks Reference](/react/hooks-reference)** for the full list and
[Overview](/react/#hooks-based-architecture) for the broader architectural picture.
