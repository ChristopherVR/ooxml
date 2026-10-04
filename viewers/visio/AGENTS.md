# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others). This file is
canonical; `CLAUDE.md` imports it. Never maintain a second set of instructions.

## READ FIRST: this repository is UI only

All VSDX logic belongs in `ChristopherVR/ooxml`, exposed as `ooxml-core/visio`:
package loading, XML, models, ShapeSheet interpretation, inheritance, geometry
evaluation, editing commands, preservation and serialization. Shared Office
logic belongs in the core's shared areas. Never copy or fork it into this viewer.

This repository owns the browser UI: one controller, one SVG renderer, one
`<visio-viewer>` custom element, six thin framework adapters, demos, browser
tests and documentation. Microsoft Visio equivalence is a long-term target,
not a description of the current beta.

### Shared Office UI lives in `ooxml-ui`

Any UI control another Office product could use belongs in `ooxml-ui`
(`packages/ui` in `ooxml`) as an `office-ui-*` web component, never as a
Visio-only copy. Examples: the ribbon tab row, ribbon buttons and groups,
toolbars, the File backstage and print preview, Find bar, rulers, status bar,
zoom slider, bottom page/sheet tab strip, dialogs (Options), Account profile,
presence and neutral icons (undo, redo, delete, full screen, fit). Word,
PowerPoint and Excel consume the same elements; docx-viewer follows the same
web-component model. Before building any control here, check whether
pptx-viewer, docx-viewer or xlsx-viewer has one; if so it belongs in ooxml-ui.

- This repository keeps only Visio product content: which commands the ribbon
  shows, Visio shortcuts, canvas tools, page/shape semantics and SVG rendering.
- Compose the ribbon and status bar from `office-ui-*` elements and react to
  their events (`office-command`, `office-status-activate` and so on). Register
  product-only glyphs with `registerIcon`; add neutral glyphs to `ooxml-ui`.
- If a needed control is missing, add it to `ooxml-ui` first, release it, then
  adopt the released version here. Do not ship a local stand-in.
- All UI ships inside the one `<visio-viewer>` element, so the six framework
  bindings get it without per-framework code. Never build controls in an adapter.

### Build the element the docx-viewer way

`docx-viewer` (`packages/web-component`) is the reference for how the viewer
element is built. New and touched UI code follows it; older template-string
code migrates when it is next changed.

- **Package layout (target).** The element lives in a private workspace
  package, `packages/web-component` (`visio-web-component`). It is bundled into
  every framework package beside `packages/bindings`, as `docx-web-component`
  is. Root `src/` is the current location; move modules there instead of adding
  new top-level layers.
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
pptx-viewer's do. A React developer gets a `forwardRef` component with an
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
2. Fix rendering, selection, controls, keyboard and view state once in `src/`.
3. For adapter changes, search all six for the same pattern and fix every
   affected adapter together. Shared helpers belong in `common.ts`.
4. Add regression coverage at the owning layer. Check actual native adapter
   lifecycle and browser behavior; compilation alone is insufficient.
5. Keep parsing and document mutations out of adapters. Editing UI delegates
   XML/package changes to `ooxml-core/visio` through the editing worker.

## Where things live

These are GitHub names under `ChristopherVR/`. Machine-local checkout paths
belong in an untracked, git-ignored `CLAUDE.local.md`, never in this file.

| Repository            | Packages                                 | Owns                                                                 |
| --------------------- | ---------------------------------------- | -------------------------------------------------------------------- |
| `visio-viewer` (this) | `visio-core`, `visio-<framework>-viewer` | Visio UI, adapters, demos, browser tests and docs                    |
| `ooxml`               | `ooxml-core`, `ooxml-ui`                 | All OOXML logic by format/shared area, plus DOM-only Office controls |
| `pptx-viewer`         | `pptx-viewer-core`, `pptx-*-viewer`      | PowerPoint UI and framework-neutral rendering                        |
| `docx-viewer`         | `docx-core`, `docx-<framework>-viewer`   | Word editor UI and six thin adapters                                 |
| `xlsx-viewer`         | XLSX core and framework viewer packages  | Excel UI and adapters                                                |
| `ole2`                | `@christophervr/ole2`                    | Legacy compound-file and binary formats                              |
| `emf-converter`       | `emf-converter`                          | EMF/WMF conversion and rendering                                     |
| `mtx-decompressor`    | `mtx-decompressor`                       | MicroType Express font decompression                                 |

### Where does my change go?

| The change is about...                                                          | Make it in                                          |
| ------------------------------------------------------------------------------- | --------------------------------------------------- |
| VSDX parsing, ShapeSheet, model, editing, saving or preservation                | `src/visio/` in `ooxml`, with core regression tests |
| XML, OPC, units, color, geometry, DrawingML or SmartArt reusable across formats | The appropriate shared area in `ooxml`              |
| DOM-only controls shared by Office viewers                                      | `packages/ui` in `ooxml` (`ooxml-ui`)               |
| Legacy binary codecs or CFB                                                     | `ole2`                                              |
| Metafile conversion                                                             | `emf-converter`                                     |
| SVG presentation, browser text measurement, controls or view state              | `src/` here                                         |
| Framework properties, events and lifecycle                                      | `packages/bindings/src`, across affected adapters   |
| Docs, demos, packaging, release scripts and browser tests                       | Here                                                |
| Office-suite launch routes                                                      | `site/apps.js` in `ooxml`                           |

Extract reusable logic when touching it instead of introducing another copy.
Separate browser rendering decisions from document semantics. Record source
repository, path, revision and adaptations when moving code; core extractions
must update `PROVENANCE.md` and retain licenses.

## Current structure and package status

- `src/controller.ts`: shared view state, selection, events and edit orchestration.
- `src/contract.ts` and `src/binding.ts`: common properties/events and lifecycle.
- `src/viewer-element.ts`: shared browser surface.
- `src/render-*.ts` and browser text layout: SVG presentation.
- `src/document-history.ts`: bounded source-backed history orchestration.
  XML/package mutation remains in core through an isolated editing worker.
- Cancellation and replacement must prevent stale worker results updating UI.
- `packages/bindings` is private source. Use explicit framework entry points
  so hosts do not load unrelated framework runtimes.
- Seven distribution packages live in `packages/{core,react,vue,angular,svelte,solid,vanilla}`. `npm run build:packages` bundles shared UI and adapters, declarations and browser workers. `npm run test:packages` installs their tarballs in a clean registry-only consumer.
- The root implementation and bindings remain private build inputs; never put them in a published manifest.

### Core development bridge

Normal builds use the released `ooxml-core/visio` dependency. Format conformance tests run in `ooxml`; viewer and packed-consumer tests exercise the released API here.

For intentional core development, the legacy `scripts/setup-core.mjs` prepares `integration/core-revision.txt`, its lock and any explicit patch in an isolated checkout. `VISIO_CORE_DIR` and `npm run link:core` support local overrides. Never overwrite sibling edits. Restore published dependency ranges before committing.

## Working agreements

- `mcp/` owns `visio-viewer-mcp`: schemas, MCP registration and its stdio CLI.
  It delegates inspection, experimental editing and filesystem execution to
  `ooxml-core/automation` and `/automation/node`. It is separate from the viewer
  distribution packages. The combined `ooxml-mcp` imports its `registerTools`.
  Release the core automation entry before publishing the MCP package.

- Node.js 22.12+, npm lockfiles, strict TypeScript with exact optional properties
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
- `docs/parity.html` is the canonical capability table; run `npm run docs:sync`
  for `docs/parity.md`. Update evidence and limits when behavior changes.
- Do not commit generated output, credentials or user documents.

## Commands

```sh
node scripts/setup-core.mjs
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run dev                 # docs landing and /demo/ workspace
npm run typecheck           # viewer and demo TypeScript
npm test                    # viewer unit and DOM tests
npm run check:bindings      # native adapter types, lifecycle and SSR
npm run check:core          # core TypeScript and focused Visio tests
npm run test:docs           # parity synchronization and site tests
npm run docs:sync           # regenerate the Markdown capability table
npm run build              # ESM/declarations and site-dist/
npm run check              # all non-browser checks and production build
npm run test:browser        # production browser interaction and landing tests
npm run fmt                # oxfmt
node --test scripts/release-plan.test.mjs
```

Run `npm run check` and `npm run test:browser` for implementation changes before
pushing. Focused docs and formatting checks suffice for guidance-only changes.
CI runs the full Linux suite. Record local environment failures and CI results
instead of silently skipping checks. Browser tests accept `CHROMIUM_PATH`;
otherwise install Playwright Chromium.

## GitHub Pages

`https://christophervr.github.io/visio-viewer/` is the public beta docs site and
`/demo/` is the vanilla playground; `/demo-react/`, `/demo-vue/`, `/demo-angular/`,
`/demo-svelte/` and `/demo-solid/` mount the same workspace (`demo/workspace.ts`) through each
framework binding (`packages/bindings/demos`, built by `scripts/build-demos.mjs`). `.github/workflows/pages.yml` builds the pinned
core and viewer, runs full checks and browser tests, and deploys `site-dist/`
on pushes to `main`. The OOXML launcher embeds the demo; update its registry
when a public demo route changes. Embedded theme follows the shared
`vitepress-theme-appearance` preference on the GitHub Pages origin.

## Releasing

See `docs/releasing.md` for the functional package release flow.

- Hourly/manual `release.yml` independently versions packages from conventional
  commits since each `<npm-name>@<version>` tag, writes changelogs, commits
  versions to `main`, creates GitHub releases and publishes with provenance.
- Shared viewer or adapter changes release all six viewer packages. Core releases also release dependents. Build scripts trigger all packages; docs and test-only changes do not release unless a shipped package README changes.
- npm trust names `ChristopherVR/visio-viewer`, `release.yml`, environment
  `npm`; the job requests `id-token: write`. `NPM_PUBLISH=true` enables it.
  No npm token is required.
- Never run `npm publish` or push release tags manually after bootstrap.
  Retry with `gh workflow run release.yml -f tag=<npm-name>@<version>`.

## Branching and git workflow

Use trunk-based development: commit directly to `main`. Do not create feature
branches unless explicitly requested. Keep commits small and releasable.
Follow each sibling repository's own guidance.

The checkout may be shared. Before committing, check `git branch --show-current`
and `git status`, and stage only your changes. Fetch and rebase onto `origin/main`
before pushing because release automation also writes there. Do not switch
a shared checkout underneath another session; push `HEAD:main` or use an
isolated worktree when needed.

## Commit conventions and style

- Conventional Commits: `<type>(<scope>): <subject>`, scoped to the affected
  area. Imperative lower-case subject, no trailing period, at most 72 characters.
  Types determine version bumps; touched paths determine package scope.
- Never add a Codex co-author trailer. Use `git commit -F` for multiline
  messages and never include assistant chat share links.
- No U+2014 in source, comments, docs, commits or UI except intentional content
  or assertions. Use punctuation or a spaced hyphen.
- Run oxfmt before committing; avoid unrelated formatting in shared trees.
