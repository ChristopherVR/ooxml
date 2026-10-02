# Native Visio viewer bindings

Six thin native integrations use one browser element and one `mountViewer` implementation:

- Vanilla: `src/vanilla.ts`, `mountViewer`
- React: `src/react.ts`, `VisioViewer`
- Vue 3: `src/vue.ts`, `VisioViewer`
- Angular: `src/angular.ts`, `VisioViewerComponent` (`visio-viewer-host`)
- Svelte 5: `src/VisioViewer.svelte`, default `VisioViewer`
- Solid: `src/solid.tsx`, `VisioViewer`

This package is private source inside the repository. It is not a published, standalone npm artifact. It imports the canonical viewer through relative source paths; a distribution build and packed-consumer verification are required before publication. Import a specific framework entry point to avoid loading unrelated framework runtimes.

## Shared contract

All adapters accept the same properties:

- `document`: a `VisioDocument` or `null` to clear it
- `pageIndex`: zero-based page index
- `zoom`: scale, where `1` means 100%
- `showToolbar`: toolbar visibility
- `events`: typed callback map using literal event names

The event map currently contains `document-load`, `document-error`, `page-change`, `zoom-change`, and `shape-select`. It is derived from the canonical contract. All adapters pass every event; none parse documents or own another renderer.

Property updates are partial. Native framework snapshots are diffed so unrelated rerenders do not undo an imperative load or a user-changed viewport. Omitted or `undefined` properties preserve current viewer state. `null` explicitly clears the document. Removing the `events` prop clears callback-map handlers. Supply a new document object for an external replacement; in-place mutation is not a supported reactivity mechanism. Page and zoom interactions may update internal viewer state; the props request changes when their framework updates them.

The imperative `ViewerHandle` exposes `element`, `controller`, `load(bytesOrBlob)` and `fit()`. Read access and operations through a framework handle fail clearly before mounting or after unmounting. A retained controller or element is also disposed when the owner unmounts. Vanilla returns the full shared binding with `update()` and idempotent `destroy()`.

The editor itself mounts on the client. Separate Node tests verify all six entry points import without browser globals; React, Vue, Svelte and Solid also render empty native hosts on the server. React and Vue have tests hydrating their actual server markup before mounting the shared element. Angular server rendering and Solid/Svelte hydration remain unverified, as does cross-document mounting. These checks use the installed versions and do not imply full SSR-framework integration coverage.

## Native usage

React uses a forwarded ref and ordinary props:

```tsx
const ref = useRef<ViewerHandle>(null);
<VisioViewer ref={ref} document={diagram} events={{ 'document-error': reportError }} />;
```

Vue exposes the same handle on its component ref and also emits the native kebab-case events:

```vue
<VisioViewer ref="viewer" :document="diagram" @document-error="reportError" />
```

Angular exposes native inputs and outputs. Add the standalone component to the application's imports:

```html
<visio-viewer-host [document]="diagram" (documentError)="reportError($event)" />
```

Angular outputs are `documentLoad`, `documentError`, `pageChange`, `zoomChange` and `shapeSelect`. The complete `events` callback map is also accepted as an input. Vue and Angular invoke the callback-map handler and then emit their native output; use either API unless both notifications are intended.

For large immutable document models, use Vue `shallowRef`/`markRaw` and Svelte `$state.raw` rather than deeply proxying a newly constructed options object on every update. Rewrapping a document in a new reactive proxy changes its identity and requests an external replacement. In-place mutation is not supported.

Svelte uses native callback props through `events`. Its component exports `load`, `fit`, and `getHandle` so the complete typed handle is accessible without exposing private DOM:

```svelte
<VisioViewer bind:this={viewer} document={diagram}
  events={{ 'document-error': reportError }} />
```

Solid supplies a ref callback with the handle on mount and `undefined` on unmount or ref replacement:

```tsx
<VisioViewer
	document={diagram()}
	viewerRef={setViewer}
	events={{ 'document-error': reportError }}
/>
```

Vanilla owns only the child it creates:

```ts
const viewer = mountViewer(host, { document: diagram, zoom: 1 });
viewer.update({ events: { 'document-error': reportError } });
// Later:
viewer.destroy();
```

Style hooks are native to each framework: React `className`/`style`, Vue inherited attributes, Angular host attributes, Svelte `class`/`style`, and Solid `class`/`style`. The application must give the host a usable height. Shared viewer styling and accessibility live in the custom element.

## Checks

Install this private package's development dependencies separately from the root:

```sh
npm ci --prefix packages/bindings --ignore-scripts
npm --prefix packages/bindings run check
```

The root viewer and the sibling `ooxml-core/visio` must first be resolvable according to the root setup instructions. This package does not replace or fork the parser.

Checks include TypeScript, `svelte-check`, native runtime lifecycle tests and real-element integration:

- React StrictMode mount, cleanup and remount
- Native prop diffing preserves imperative documents/user zoom across unrelated updates
- Late Blob success and failure after unmount/remount cannot affect the new viewer
- Throwing ref-release callbacks cannot prevent Solid resource cleanup
- No-DOM SSR imports and server rendering; React/Vue server-markup hydration
- Native Vue, Angular zoneless, Svelte and Solid mounting and reactive updates
- All property names, every event key, latest callbacks and callback removal
- Retained handle rejection after unmount
- Actual shared custom-element rendering and malformed-load rejection through all six adapters
- Disposal clears the real element, listeners and owned DOM

The focused lifecycle tests mock only the shared mount to isolate framework wiring. The integration matrix uses the real controller, parser and custom element. These tests run in jsdom and do not replace real-browser visual, accessibility, layout or framework-version compatibility tests. Only the installed versions are currently verified.
