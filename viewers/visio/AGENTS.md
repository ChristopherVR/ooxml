# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others). This file is
canonical; `CLAUDE.md` imports it. Never maintain a second set of instructions.

## READ FIRST: this viewer is the framework-binding layer

`viewers/visio` lives inside the `ooxml` monorepo; the root `AGENTS.md` applies and describes the
layout, the verify/push flow and the commit rules. All VSDX logic belongs in `src/core/visio`
(exposed as `ooxml-core/visio`): package loading, XML, models, ShapeSheet interpretation,
inheritance, geometry evaluation, editing commands, collaboration mapping, preservation and
serialization. Shared Office logic belongs in the core's shared areas. Never copy or fork it into
this viewer.

The browser UI (one controller, one SVG renderer, one `<visio-viewer>` custom element) lives in
`src/ui/src/visio` and ships as `ooxml-ui/visio`. This folder owns six thin framework adapters
(`packages/bindings` and the `packages/*` distributions), the composable building blocks,
documentation and the package build. The demos are in `demos/visio/` and the browser tests in
`e2e/visio/`. Microsoft Visio equivalence is a long-term target, not a description of the current
beta.

### Shared Office UI lives in `ooxml-ui`

Any UI control another Office product could use belongs in `ooxml-ui` (`src/ui`) as an
`office-ui-*` web component, never as a Visio-only copy. Examples: the ribbon tab row, ribbon
buttons and groups, toolbars, the File backstage and print preview, Find bar, rulers, status bar,
zoom slider, bottom page/sheet tab strip, dialogs (Options), Account profile, presence, the chart
elements (`office-ui-chart-picker`, `office-ui-chart-data-sheet`, `office-ui-chart-format-pane`) and
neutral icons (undo, redo, delete, full screen, fit). Word, PowerPoint and Excel consume the same elements.
Before building any control here, check whether `src/ui/src/pptx`, `docx`, `xlsx` or the shared
elements already have one; if so it belongs in `src/ui`.

- Visio's own UI code (`src/ui/src/visio`) keeps only Visio product content: which commands the ribbon
  shows, Visio shortcuts, canvas tools, page/shape semantics and SVG rendering.
- Compose the ribbon and status bar from `office-ui-*` elements and react to
  their events (`office-command`, `office-status-activate` and so on). Register
  product-only glyphs with `registerIcon`; add neutral glyphs to `ooxml-ui`.
- If a needed control is missing, add it to `src/ui` first and use it from there, in
  the same change. Do not ship a local stand-in.
- All UI ships inside the one `<visio-viewer>` element, so the six framework
  bindings get it without per-framework code. Never build controls in an adapter.

### Build the element the docx way

The docx editor (`src/ui/src/docx`) is the reference for how the viewer element is
built. New and touched UI code follows it; older template-string code migrates when
it is next changed.

- **Package layout.** The element, controller and renderer live in
  `ooxml-ui/visio` (`src/ui/src/visio`); every framework package depends on it.
  `viewers/visio/src/index.ts` only re-exports it, and
  `packages/bindings` (framework lifecycle adapters) is the only private input
  bundled into each framework package.
- **DOM is built by typed builder functions**, never HTML template strings.
  Builders live in `ribbon-parts.ts` (`group`, `tool` and similar) and create
  elements with `createElement`, composing `office-ui-*` controls. Document
  text is only ever assigned with `textContent`.
- **One typed action union.** Every ribbon control carries a
  `VisioRibbonAction` (`ribbon-action.ts`) and emits a `ribbon-action` event.
  `routeRibbonAction` (`ribbon-router.ts`) sends it to the controller that owns
  it. Keyboard shortcuts route through the same router. Stable string ids are
  for state sync and tests only, never for dispatch logic.
- **One module per ribbon tab** (`ribbon-home.ts`, `ribbon-view.ts`), assembled
  by `createRibbon()` in `ribbon.ts`. The status bar and page tabs have their
  own builder modules.
- **Styles are CSS files.** Write them under `styles/*.css`, import them with
  `?inline` and join them in `styles/index.ts`. Theme tokens come from a typed
  `theme/` module (defaults to CSS variables, with `theme` set to
  `light`, `dark` or `auto`), and shared controls read them through
  `--office-*` aliases.
- **Vite builds all JavaScript; `tsc` only type-checks and emits declarations.**
  The root package (`vite.lib.config.ts`) and the framework packages
  (`scripts/build-packages.mjs`) are Vite library builds; workers land in
  `assets/` and `?inline` CSS is inlined. Angular is not an exception: its
  binding uses `inject()` and legacy decorators, which Vite's Oxc transform
  lowers (`oxc.decorator.legacy`). VoidZero's Oxc Angular Compiler
  (`@oxc-angular/vite`) is still experimental; adopt it when it is stable or
  when the binding needs AOT templates. Never add a `tsc` JavaScript emit step.
- **Typed element surface.** Keep a typed event map with a runtime list
  (`events.ts`), observed attributes (`editor-attributes`-style) and the public
  API (`element-api`-style) in separate small modules. Tests sit next to the
  code.

### One viewer, six thin bindings

React, Vue, Angular, Svelte, Solid and vanilla use the same viewer. Adapters in
`packages/bindings/src` own lifecycle, properties, events and handle forwarding.

The UI is one web component, but every binding must feel native, as
the pptx bindings do. A React developer gets a `forwardRef` component with an
imperative handle, `on*` callback props and hooks; Angular gets `@Input`,
`@Output`, lifecycle hooks and signals; Vue gets props, emits, `expose` and
composables; Svelte and Solid get their own reactive primitives. Viewer state
(page, zoom, selection, edit and undo state) is exposed through each framework's
reactivity (for example `useSyncExternalStore` in React and signals in Angular
and Solid), built on the controller's `subscribe`, never by polling or by
reaching into the shadow DOM. Hosts never touch `<visio-viewer>` internals to
integrate.

1. Diagnose whether a bug comes from document logic, shared view behavior or
   framework wiring before choosing a file to change.
2. Fix rendering, selection, controls, keyboard and view state once in `src/ui/src/visio/` (shipped as `ooxml-ui/visio`).
3. For adapter changes, search all six for the same pattern and fix every
   affected adapter together. Visio-specific helpers (the handle contract
   `VIEWER_CONTRACT`, the revision-aware patches of `mountFrameworkViewer`,
   `withEventEmitter`) belong in `packages/bindings/src/common.ts`; binding
   machinery every product shares (the deferred handle, prop picking, the React
   and Vue component factories) lives in `ooxml-ui/bindings`
   (`src/ui/src/bindings`). `packages/bindings/tests/contract-parity.test.ts`
   checks every adapter against the contract.
4. Add regression coverage at the owning layer. Check actual native adapter
   lifecycle and browser behavior; compilation alone is insufficient.
5. Keep parsing and document mutations out of adapters. Editing UI delegates
   XML/package changes to `ooxml-core/visio` through the editing worker.

## Where things live

The root `AGENTS.md` has the repository table. For Visio: logic in `src/core/visio`, the viewer UI
in `src/ui/src/visio`, adapters and packages here, demos in `demos/visio/`, browser tests in
`e2e/visio/`. Machine-local checkout paths belong in an untracked, git-ignored `CLAUDE.local.md`,
never in this file.

### Where does my change go?

| The change is about...                                                          | Make it in                                        |
| ------------------------------------------------------------------------------- | ------------------------------------------------- |
| VSDX parsing, ShapeSheet, model, editing, saving or preservation                | `src/core/visio/`, with core regression tests     |
| XML, OPC, units, color, geometry, DrawingML or SmartArt reusable across formats | The appropriate shared area in `src/core`         |
| DOM-only controls shared by Office viewers                                      | `src/ui` (`ooxml-ui`)                             |
| Legacy binary codecs or CFB                                                     | `ole2`                                            |
| Metafile conversion                                                             | `emf-converter`                                   |
| SVG presentation, browser text measurement, controls or view state              | `src/ui/src/visio/`                               |
| Framework properties, events and lifecycle                                      | `packages/bindings/src`, across affected adapters |
| Docs, packaging and build scripts                                               | Here                                              |
| Demos and browser tests                                                         | `demos/visio/` and `e2e/visio/`                   |
| Office-suite launch routes                                                      | `site/apps.js`                                    |

Extract reusable logic when touching it instead of introducing another copy.
Separate browser rendering decisions from document semantics. Record source
repository, path, revision and adaptations when moving code; core extractions
must update `PROVENANCE.md` and retain licenses.

## Current structure and package status

- `src/index.ts`: a re-export of `ooxml-ui/visio`, where the viewer lives. In `src/ui/src/visio`:
- `controller.ts`: shared view state, selection, events and edit orchestration.
- `binding.ts`: the lifecycle shared by every adapter; `viewer-element.ts`: the shared browser surface.
- `render-*.ts` and browser text layout: SVG presentation.
- In the core (`src/core/visio/ui`): `contract.ts` (common properties and events), `document-history.ts`
  (bounded source-backed history orchestration; XML/package mutation remains in core through an
  isolated editing worker) and `document-text-search.ts`.
- Cancellation and replacement must prevent stale worker results updating UI.
- `packages/bindings` is private source. Use explicit framework entry points
  so hosts do not load unrelated framework runtimes.
- Seven distribution packages live in `packages/{core,react,vue,angular,svelte,solid,vanilla}`. `bun run build:packages` bundles shared UI and adapters, declarations and browser workers. `bun run test:packages` installs their tarballs in a clean consumer with the registry's `ooxml-core` and `ooxml-ui`; `bun run test:packages:workspace` installs them with this checkout's core and UI, which is what CI runs (`--workspace-runtime`), as it is the combination a release publishes together.
- The root implementation and bindings remain private build inputs; never put them in a published manifest.
- The shared contract (`propertyKeys` and `eventKeys` in `src/core/visio/ui/contract.ts`) has the
  properties `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`,
  `collaboration`, `storage` and `customization`, and the events listed there, including
  `collaboration-change` and `file-command`. A contract change updates every adapter, the
  contract-parity test and `docs/bindings.md`. The
  building blocks are read-only unless the host passes `canEdit: true` (`docs/building-blocks.md`),
  `markClean()` declares exported bytes saved, and collaboration merges per page and shape
  (`docs/collaboration.md`).

### Workspace

This folder is a member of the root Bun workspace: one `bun install` at the repository root, one
lockfile. `ooxml-core` and `ooxml-ui` resolve to the packages in the same checkout (`"*"` ranges), so
there is no bridge to a separately released core. A change that needs the core and the viewer lands
in one push. (`scripts/link-core.mjs` and `VISIO_CORE_DIR` only matter for pointing this folder at a
different core checkout.)

## Working agreements

- `mcp/` owns `visio-viewer-mcp`: schemas, MCP registration and its stdio CLI.
  It delegates inspection, experimental editing and filesystem execution to
  `ooxml-core/automation` and `/automation/node`. It is separate from the viewer
  distribution packages. The combined `ooxml-mcp` imports its `registerTools`.
  The MCP package and the core automation entry it uses are released by the same flow.

- Node.js 22.12+, Bun workspace (one root lockfile), strict TypeScript with exact optional properties
  and unchecked indexed access, Vitest and Playwright. Keep modules under
  300 lines where practical.
- Parsing, preservation and package round-trip tests belong in core. Viewer
  tests cover controller, rendering, controls and lifecycle. Exercise
  load/edit/export/reload when preservation is involved.
- Documents stay local: no uploads, telemetry or remote-document fetches.
  Never inject document XML/HTML into the DOM. Preserve package, scene, resource
  and worker limits when adding features.
- Report unsupported features honestly. Plain-text editing does not establish
  rich-text, formula recalculation or native Visio save/reopen fidelity.
  Generated fixtures alone do not establish native visual parity.
- `docs/parity.md` is the canonical capability table (the documentation site is
  VitePress, like the other viewers'). `bun run test:docs` builds the site and checks
  the ledger's inventory and the honesty statements. Update evidence and limits
  when behavior changes.
- Do not commit generated output, credentials or user documents.

## Commands

Install once at the repository root with `bun install`, then run these from `viewers/visio`:

```sh
bun run dev                 # docs landing and /demo/ workspace
bun run typecheck           # viewer and demo TypeScript
bun run test                # viewer unit and DOM tests
bun run check:bindings      # native adapter types, lifecycle and SSR
bun run check:core          # core TypeScript and focused Visio tests
bun run test:docs           # VitePress docs build, ledger and theme tests
bun run docs:build          # VitePress docs plus every demo into docs/.vitepress/dist
bun run build               # ESM/declarations and the playground in site-dist/
bun run check               # all non-browser checks and production build
bun run build:packages      # the seven framework and core distributions
bun run test:packages:workspace  # packed distributions, with this checkout's core and ui (CI)
bun run test:browser        # production browser interaction and landing tests (ends with the packed-package browser check)
bun run fmt                 # oxfmt
```

Run `bun run check` and `bun run test:browser` for implementation changes, and the root
`bun run verify` before pushing (it runs what the CI planner says your commits reach). Focused docs and
formatting checks suffice for guidance-only changes. CI runs the full Linux suite. Record local
environment failures and CI results instead of silently skipping checks. Browser tests accept
`CHROMIUM_PATH`; otherwise install Playwright Chromium.

## GitHub Pages

`https://christophervr.github.io/ooxml/visio/` is the public beta docs site and
`/demo/` is the vanilla playground (also at `/demo-vanilla/`); `/demo-react/`, `/demo-vue/`, `/demo-angular/`,
`/demo-svelte/` and `/demo-solid/` mount the same workspace (`demos/visio/demo/workspace.ts`) through each
framework binding (`demos/visio/demo-*`, built by `scripts/build-demos.mjs`). The root `scripts/build-pages.mjs`
runs this viewer's `scripts/build-pages.mjs` (VitePress docs plus every demo, into `docs/.vitepress/dist/`) and
assembles it under `/ooxml/visio/` of the one Pages site. The Office launcher embeds the demo; update its
registry (`site/apps.js`) when a public demo route changes. Embedded theme follows the shared
`vitepress-theme-appearance` preference on the GitHub Pages origin.

## Releasing

See `docs/releasing.md` and the root `docs/releasing.md`. The root release flow
(`scripts/release-plan.mjs`, `.github/workflows/release.yml`, the Visio package table in
`scripts/viewer-packages.mjs`) versions the packages from conventional commits, writes changelogs
and publishes with npm trusted publishing. Never run `npm publish` or push release tags manually.

## Branching and git workflow

Follow the root `AGENTS.md`: trunk-based development on `main`, no feature branches unless
explicitly requested, small releasable commits, and your own worktree when the checkout is shared
(`git push origin HEAD:main`). Run the root `bun run verify` before pushing.

## Commit conventions and style

- Conventional Commits: `<type>(<scope>): <subject>`, scoped to the affected
  area. Imperative lower-case subject, no trailing period, at most 72 characters.
  Types determine version bumps; touched paths determine package scope.
- Never add a Codex co-author trailer. Use `git commit -F` for multiline
  messages and never include assistant chat share links.
- No U+2014 in source, comments, docs, commits or UI except intentional content
  or assertions. Use punctuation or a spaced hyphen.
- Run oxfmt before committing; avoid unrelated formatting in shared trees.
