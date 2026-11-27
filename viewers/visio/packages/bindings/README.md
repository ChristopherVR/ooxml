# @christophervr/visio-viewer-bindings

**Private source for the six native framework adapters of the Visio viewer. It is not published to npm.**

`bun run build:packages` (in `viewers/visio`) bundles each entry point into its framework distribution under `packages/{react,vue,angular,svelte,solid,vanilla}`, and those packages are what is published. Install one of them instead; see the [Visio viewer README](../../README.md) and the [bindings guide](../../docs/bindings.md).

## Entry points

| Export path | File                     | Native surface                                                                  |
| ----------- | ------------------------ | ------------------------------------------------------------------------------- |
| `./vanilla` | `src/vanilla.ts`         | `mountViewer`, `createViewerBuildingBlocks`                                     |
| `./react`   | `src/react.ts`           | `VisioViewer` (`forwardRef`), `useVisioViewerState`                             |
| `./vue`     | `src/vue.ts`             | `VisioViewer` (props, emits, exposed handle), `useVisioViewerState`             |
| `./angular` | `src/angular.ts`         | `VisioViewerComponent` (`visio-viewer-host`: inputs, outputs, `state()` signal) |
| `./svelte`  | `src/VisioViewer.svelte` | `VisioViewer` (exported handle methods, `getHandle()`, `getState()` store)      |
| `./solid`   | `src/solid.tsx`          | `VisioViewer` (`viewerRef`), `createVisioViewerState`                           |

Import a specific entry point so a host never loads an unrelated framework runtime. `src/building-blocks.ts`, `src/vanilla-building-blocks.ts` and `src/toolbar-controls.ts` hold the framework-neutral parts of the composable building blocks.

## How the adapters fit together

- `src/common.ts` holds the Visio contract: `VIEWER_CONTRACT` (every property is forwarded manually, the handle members and their kinds), `mountFrameworkViewer` (sends only changed properties as revision-aware patches, so a re-render never undoes an imperative load or the user's zoom), `viewerStateSource` (the subscription every state hook builds on) and `withEventEmitter` (calls the `events` callback first, so `false` can cancel `file-command`, then emits the native output).
- The property and event keys come from `src/core/visio/ui/contract.ts` (`propertyKeys`, `eventKeys`); the viewer itself is `<visio-viewer>` in `ooxml-ui/visio` (`src/ui/src/visio`).
- React and Vue are built from the shared factories in `ooxml-ui/bindings`; Angular installs its handle members from the contract; Svelte and Solid mount the shared element directly.
- A fix in one adapter almost always belongs in all six: change `common.ts` or the shared element where possible, and add the regression to each adapter's tests.

## Checks

From `viewers/visio` (after `bun install` at the repository root):

```sh
bun run check:bindings   # TypeScript, svelte-check, jsdom and SSR tests in this package
```

`tests/contract-parity.test.ts` checks that every adapter receives every property, reports every event through the callback map and its native output, and exposes every handle member. Other tests cover React StrictMode, state subscriptions across remounts, prop diffing, retained handles rejecting after unmount, server imports and React/Vue hydration. They run in jsdom; real-browser behaviour is covered by the Playwright suites in `e2e/visio/` (`bun run test:browser`). Angular server rendering and Solid/Svelte hydration are not verified.
