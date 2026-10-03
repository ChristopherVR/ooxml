# Office UI package plan

`ooxml-ui` (`packages/ui`) is the second published package of this repository
(section 6 of `agnostic-core-plan.md` proposed the idea; this file is the working plan and
supersedes its layout). It holds the web components every Office product needs, and nothing
product-specific.

## Decisions

- **Layout.** The repository root stays the published `ooxml-core`; `src/` and
  its paths do not move. `package.json` lists `"workspaces": ["packages/*"]`, so `packages/ui` is
  a Bun workspace next to it. The root is not a workspace member (a package cannot be its own
  workspace), so the UI declares `"ooxml-core": "^0.1.0"` as an ordinary semver
  dependency and development resolves the core **from source** (tsconfig `paths` and a vitest
  alias to `../../src/<area>/index.ts`; declaration builds read `../../dist`). Keep that range
  in step with the core when it reaches 0.2.0 (release tooling: see "Release").
- **Direction.** UI may import core types and pure helpers; core never imports the UI (nothing in
  `src/` references `packages/`). Add a CI check next to the other release checks.
- **Style.** Lit-free vanilla custom elements built lazily inside `define*` functions, so
  importing a module never touches `HTMLElement` (SSR-safe). Shadow-root CSS with `--office-*`
  tokens that all have fallbacks; each control ships forced-colors and coarse-pointer rules.
  Typed events: `office-command`, `office-dialog-close`, `office-presence-select`,
  `office-status-activate`, `office-smartart-render`, plus native `input`/`change` for value
  controls. ESM only, strict TypeScript (`exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`).
- **Product rule.** Viewers list `ooxml-ui` as a normal `dependency` of every
  published editor package, so users never install it by hand. The README and install docs of
  the viewers never mention it as a separate install step.
- **Honesty.** `office-ui-smartart` shows the cached drawing and reports approximated
  geometry, fills and flattened 3D in its render event. It makes no layout-parity claim.

## Classification

Sources: `pptx` = `pptx-viewer/packages/shared/src/web-components` (and `render/`), `docx` =
`docx-viewer/packages/web-component/src`.

| Piece                                                                                                    | Where today             | Decision                           | Notes                                                                                                                           |
| -------------------------------------------------------------------------------------------------------- | ----------------------- | ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| Style attachment (`attachControlStyles`), contract stamping                                              | pptx                    | **moved** (`styles`, `registry`)   |                                                                                                                                 |
| Theme tokens, dark mode, forced-colors, touch targets                                                    | pptx host CSS           | **moved** (`theme`)                | `--pptx-*` maps to `--office-*` (table below)                                                                                   |
| Icon registry + neutral glyphs, `office-ui-icon`                                                         | pptx ribbon-icons       | **moved** (`icons`)                | 16 neutral glyphs; view/insert/animation icon sets stay                                                                         |
| Button (command), checkbox, switch                                                                       | pptx                    | **moved**                          | `ribbon-command` becomes `office-ui-button` with a generic `command` id                                                         |
| Select                                                                                                   | pptx                    | **moved**                          | `options` property now; `<option>` child parsing is a later item                                                                |
| Ribbon group, toolbar                                                                                    | pptx                    | **moved**                          | pptx's compact-row ResizeObserver centring is product layout and stays                                                          |
| Dialog shell                                                                                             | docx dialogs            | **moved** (new)                    | Focus trap, Escape, restore focus; the docx dialogs keep their content                                                          |
| Status-bar primitives                                                                                    | docx / pptx             | **moved** (new)                    | Container + item                                                                                                                |
| Status-bar zoom slider, bottom tab strip (pages/sheets)                                                  | visio / xlsx            | **moved** (new)                    | Office zoom (step to next 10%, optional fit) and document tabs with roving focus; first consumer is visio-viewer                |
| Presence avatars                                                                                         | both (collab)           | **moved** (new)                    | Renders a minimal participant view; adapter `participantsFromAwareness`; aligns to core `collab` awareness when that area lands |
| SmartArt SVG renderer                                                                                    | pptx `smartart-drawing` | **moved** (new, on core `diagram`) | Presets without outlines are rectangles and reported; arc commands in custom paths are dropped and reported                     |
| Split button, menu/popover, colour picker, tooltip                                                       | pptx/docx               | **later (wave 2)**                 | Need a shared popover (anchoring, outside click, focus return) first; build once, then the three consumers                      |
| Search field, ribbon gallery, ribbon section, ribbon toggle                                              | pptx                    | **later (wave 2)**                 | Gallery and section depend on the popover                                                                                       |
| Theme editor                                                                                             | pptx                    | **later (wave 3)**                 | Format-neutral in intent but bound to the pptx theme model; move after the shared theme model is in core                        |
| i18n plumbing (`shared/src/i18n`, `packages/locales`)                                                    | pptx                    | **later (wave 2)**                 | Catalogues stay data (`ooxml-ui/i18n/<locale>`); keys namespaced `ui.*`                                                         |
| Collaboration UI beyond avatars (cursors, follow, comments)                                              | both                    | **later**                          | Depends on the core `collab` session API                                                                                        |
| Chart and 3D SmartArt renderers (`three-view`, `smartart-3d`)                                            | pptx                    | **later / stays**                  | Consume pptx runtime models, not neutral core output; revisit when core owns `chart` and the 3D drawing model                   |
| Presentation Home/Insert/Draw/Transitions/Animations/View strips, slide-show options, subtitle settings  | pptx                    | **stays**                          | Product content                                                                                                                 |
| Word ribbon content, backstage, contextual tabs, dialogs (borders, columns, bookmark...), comments panel | docx                    | **stays**                          | Product content                                                                                                                 |
| Framework bindings (React, Vue, Angular, Svelte, Solid, vanilla)                                         | both                    | **stays**                          | Lifecycle adapters of the viewer's own element                                                                                  |

## Token map

| pptx                            | ooxml-ui                     |
| ------------------------------- | ---------------------------- |
| `--pptx-foreground`             | `--office-foreground`        |
| `--pptx-muted-foreground`       | `--office-muted-foreground`  |
| `--pptx-background`             | `--office-background`        |
| `--pptx-border`, `--pptx-input` | `--office-border`            |
| `--pptx-primary`                | `--office-accent`            |
| `--pptx-primary-foreground`     | `--office-accent-foreground` |
| `--pptx-ring`                   | `--office-ring`              |

The pptx editor maps its existing variables to the `--office-*` ones on its host element, so
its theming API does not change.

## Tag migration (pptx-ui-* to office-ui-*)

pptx-viewer must not break, so migration is two-step and the public pptx tags never disappear.

1. **Wave 1 (this PR).** ooxml-ui exists and is released; pptx is untouched.
2. **Wave 2 (pptx-viewer).** Add the dependency, call `registerOfficeUi({ theme: false })` and map
   tokens. For each moved element the pptx tag becomes a thin alias:
   `class PptxUiCheckbox extends customElements.get('office-ui-checkbox')` (a trivial subclass,
   because one constructor cannot be registered under two names), defined under the old tag, so
   `<pptx-ui-checkbox>` keeps its attributes, properties and events. Differences are translated
   inside the alias, not in callers: `pptx-ui-ribbon-command` maps `data-ribbon-control` to
   `command` and re-emits `office-command` as `command-request` `{ id }`; `pptx-ui-ribbon-toggle`
   keeps its `toggle-request` event on top of `office-ui-checkbox`; `pptx-ui-ribbon-group` is a
   subclass that keeps the compact-row centring. The old class bodies and styles are deleted;
   the pptx registration-contract and SSR tests run against the aliases unchanged.
3. **Wave 3.** Product code uses `office-ui-*` directly; the aliases are marked deprecated in
   the pptx changelog and removed in the next major.

## Viewer adoption steps

docx-viewer:

1. Add `ooxml-ui` to the `dependencies` of each published editor package
   (not to peers, not to the install docs) and register through `registerOfficeUi()` from the
   editor's own entry (SSR-safe, idempotent).
2. Replace the local button/checkbox/select/dialog/status primitives with `office-ui-*`
   elements one family at a time; map the existing editor CSS variables to `--office-*` on the
   host. Keep the ribbon content and dialogs' bodies in docx.
3. Use `office-ui-smartart` for the diagram drawings the `DocxDiagram` model exposes
   (`InlineImage.diagram`), supplying `schemeColors` from the document theme.
4. Render collaboration presence with `office-ui-presence` fed by
   `participantsFromAwareness` (docx identity model).

pptx-viewer: wave 2 above; additionally `office-ui-smartart` can replace the 2D SVG path for
diagrams that have a cached drawing, behind a parity test.

## Release

Independent versions per package (as proposed in `agnostic-core-plan.md` 6.2): tag
`ooxml-ui@<version>`, changelog from commits touching `packages/ui`. The release
scripts and CI are owned by the release-tooling work; they must (a) plan and publish
`packages/ui` separately, (b) rewrite nothing (the dependency range is already a plain `^`
range), (c) run `bun run --cwd packages/ui typecheck|test`, then root `build`, `packages/ui`
`build` and `test:package` before publishing. Core must be built before the UI (declarations
read `../../dist`).

## First publish of `ooxml-ui`

A brand-new scoped name cannot use trusted publishing yet: the package does not exist, so there
is nothing to attach a trusted publisher to. The first publish is manual, with a logged-in
maintainer account (`npm login`, 2FA), then trusted publishing takes over:

```
bun install
bun run build                       # core declarations/bundles (the UI reads ../../dist)
bun run --cwd packages/ui typecheck && bun run --cwd packages/ui test
bun run --cwd packages/ui build
bun run --cwd packages/ui test:package   # packs core + UI, installs, imports every entry (Node + jsdom)
cd packages/ui
npm pack --dry-run                  # inspect: dist/, LICENSE, NOTICE, README.md, package.json only
npm login                           # manual, with 2FA
npm publish --access public
```

Afterwards, on npmjs.com, package Settings, Trusted Publisher: GitHub Actions, repository
`ChristopherVR/ooxml`, workflow `release.yml`, environment `npm`. From then on only the
workflow publishes it (OIDC, provenance, no token).
Do not tag by hand. The core `^0.1.0` range must resolve on the registry at first publish
(core 0.1.0 is published).

## Tests

`packages/ui` runs Vitest in jsdom: registration (idempotence, contract mismatch, prefix),
SSR import under Node, property/attribute contracts, events and their order, ARIA roles and
states, keyboard (select, toolbar, switch, dialog Tab trap and Escape), forced-colors,
dark-mode and coarse-pointer CSS presence, presence overflow, and the SmartArt render report.
`test:package` packs the tarballs and imports every entry both under Node and in a jsdom window.
Not covered (needs a real browser, planned with the viewers' browser contract tests): layout,
focus-visible painting, real forced-colors rendering, and `delegatesFocus` behaviour.
