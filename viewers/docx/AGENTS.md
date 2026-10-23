# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others) working in the Word
viewer, `viewers/docx` in the `ooxml` repository. The repository's root
`AGENTS.md` applies as well and wins on anything repository-wide (releases,
`bun run verify`, branching, commits). This file is canonical for the viewer:
`CLAUDE.md` only imports it, so edit this file and never fork the two.

## READ FIRST: this viewer is UI only

All Word logic lives in the `docx` area of `ooxml-core` (`src/core/docx` at the
repository root): the document model, parser, serializer, editing commands,
layout (`ooxml-core/docx/layout`), document loading and the legacy `.doc` reader
(`ooxml-core/docx/load`) and collaboration (`ooxml-core/collab`). The
`<docx-editor>` element itself, with the ribbon, dialogs and panels, lives in
`ooxml-ui/docx` (`src/ui/src/docx`). This directory holds the **framework
bindings**, the published package shells, demos' build wiring, browser tests and
the docs site. The packages here resolve `ooxml-core` and `ooxml-ui` to the
workspace (`"*"` ranges), so a change to the core or the UI is visible here at
once and ships in the same push.

Before writing code, decide where it belongs (see
[Where does my change go?](#where-does-my-change-go)). Never add, copy or fork
logic here that belongs in the core or the UI package.

### One editor, six thin bindings

There is **one** editor: the `<docx-editor>` web component, which lives in `ooxml-ui/docx` (`src/ui/src/docx`); `packages/web-component` only re-exports it. The
bindings in `packages/bindings` (`react.tsx`, `vue.ts`, `angular.ts`, `solid.ts`,
`WordEditor.svelte`, and the framework-neutral `index.ts` that Vanilla and the
others build on) are lifecycle and event adapters only: they create
the element, forward props and options, and re-emit its events.

- A behaviour bug (ribbon, dialog, editing, rendering, keyboard) is fixed in
  `src/ui/src/docx`, once, and reaches all six bindings automatically.
- A binding-only bug is almost always prop, event or lifecycle wiring. When you
  fix one, grep the other adapters in `packages/bindings/src` for the same
  pattern and fix them in the same change; the binding contract tests in
  `packages/bindings/src/*.test.ts` and `packages/vanilla/src/index.test.ts`
  should cover all of them.
- Never put editor logic in an adapter. The props, events and handle every
  adapter shares are one contract, `EDITOR_CONTRACT` in
  `packages/bindings/src/contract.ts`, and `contract-parity.test.ts` checks each
  adapter against it. Binding machinery that is not Word-specific (prop
  tracking, event mapping, the deferred handle, the React, Vue and Solid
  component factories) lives in the shared kit `ooxml-ui/bindings`
  (`src/ui/src/bindings`, see its README section); a helper only Word needs goes
  in `packages/bindings/src/index.ts` (or the editor in `src/ui/src/docx`).
  `common.ts` is only the re-export surface every package shares.

## Working agreements

- `mcp/` owns `docx-viewer-mcp`: schemas, MCP registration and its stdio CLI.
  It delegates headless document operations and filesystem execution to
  `ooxml-core/automation` and `/automation/node`. The combined `ooxml-mcp`
  server imports this package's `registerTools`; never duplicate document logic here.

- Published packages: `docx-core` and one
  self-contained `docx-<framework>-viewer` per framework (react, vue, angular,
  svelte, solid, vanilla; vanilla also owns the plain `<docx-editor>` entry and
  `mountEditor`). Every framework package re-exports the model API
  (`createDocument`, `loadDocument`, ...) so an application needs one install.
- `web-component` (`docx-web-component`) and `bindings` (`docx-bindings`) are
  `private: true` internal packages, bundled into every framework package by
  `scripts/build-packages.mjs`. A published tarball may import only
  `docx-core`, `ooxml-core`, `ooxml-ui`, ProseMirror and its framework peer;
  `bun run check:published` and `bun run pack:smoke` enforce it.
- `docx-core` (`packages/core`) is only a thin re-export of `ooxml-core/docx`
  and `ooxml-core/docx/embedded`; `bun run check:shared` keeps it logic-free.
- `ooxml-ui` (the shared `office-ui-*` custom elements) is a real registry
  dependency of every framework package, never bundled. Its element tag names
  keep the `office-ui-` prefix. Generic chrome is drawn by it: the context menu,
  status bar and zoom slider, title bar, File view (backstage) and the modal
  shell of every format dialog (`dialog-shell.ts`), plus SmartArt. The docx
  content inside them (dialog fields, File pages, ribbon commands) lives with
  the editor in `src/ui/src/docx` and is fed translated state. The shared
  elements read `--office-*` tokens; the editor's `--dve-*` theme is mapped onto
  them by the theme bridge in `src/ui/src` (`theme-bridge.ts`,
  `building-block-styles.ts`), so never hard-code colours for them. A generic
  element that is missing something (a translatable label, a tooltip) gets an
  additive option in `src/ui`, not a copy.
- Still hand-built in `src/ui/src/docx`, pending a move to shared `ooxml-ui`
  controls: the Word ribbon (native controls and its own overflow folding, see
  `ribbon-tabs.ts` and `ribbon-home.ts`), the indent ruler (`ruler.ts`;
  draggable markers; the shared ruler is decorative) and the comments panel
  (`comments-panel.ts`). Move one only when `ooxml-ui` can carry its behaviour
  without a test-visible regression.
- `ole2` (legacy compound-file codecs) reaches this repository only through
  `ooxml-core/docx/load`, which inlines it. Never add `@christophervr/ole2` or an
  internal package to a manifest here, and never fork its code.
- The browser `createCanvasMeasurer` stays in
  `src/ui/src/docx/canvas-measurer.ts` and is injected into the core's layout
  engine; it is the one piece of layout that needs a DOM.
- Unsupported document features must be reported honestly (for example edited
  SmartArt is drawn with an approximate layout, not Word's, and its styles and
  colours are not editable; charts paint from their cached values for the common families (including bubble and
  open-high-low-close stock), bar/line combinations, waterfalls, histograms, treemaps and
  sunbursts, and surface, volume-stock and other chartex charts show a labelled frame; classic charts can be
  inserted and edited (type, legend, data) through the shared chart core, chartex charts stay read-only). Never claim Word layout
  parity or lossless export.
- Bun, TypeScript strict mode, Vitest (tests next to the code) and Playwright
  browser tests (`e2e/docx/*.spec.ts`). Keep source modules under 300 lines where
  practical. Add regression tests for editing, preservation and binding
  contracts; parsing and round-trip tests belong in the core.
- The Word fixtures used by the browser specs are copies in `e2e/docx/support`.

## Packages

| Directory                | npm name              | Published | What it is                                                                                                           |
| ------------------------ | --------------------- | --------- | -------------------------------------------------------------------------------------------------------------------- |
| `packages/core`          | `docx-core`           | yes       | Re-export of `ooxml-core/docx` and `/docx/embedded`.                                                                 |
| `packages/web-component` | `docx-web-component`  | internal  | One-line re-export of `ooxml-ui/docx`, the `<docx-editor>` element.                                                  |
| `packages/bindings`      | `docx-bindings`       | internal  | Framework adapters over one framework-neutral core (`contract.ts`, `index.ts`, `common.ts`) and `ooxml-ui/bindings`. |
| `packages/react`         | `docx-react-viewer`   | yes       | React entry (`WordEditor`).                                                                                          |
| `packages/vue`           | `docx-vue-viewer`     | yes       | Vue 3 entry (`WordEditor`).                                                                                          |
| `packages/angular`       | `docx-angular-viewer` | yes       | Angular entry (`WordEditorComponent`).                                                                               |
| `packages/svelte`        | `docx-svelte-viewer`  | yes       | Svelte 5 entry (`/runtime`, `mountEditor`).                                                                          |
| `packages/solid`         | `docx-solid-viewer`   | yes       | Solid entry (`WordEditor`).                                                                                          |
| `packages/vanilla`       | `docx-vanilla-viewer` | yes       | Plain `<docx-editor>` registration and `mountEditor`.                                                                |

`demos/docx/demo-vanilla` is the single demo app; `?framework=<name>` mounts the
editor through that framework's adapter (`demos/docx/demo-vanilla/framework.ts`), and
the browser tests run against its production build.

## Where things live

This viewer is a member of the `ooxml` repository. Other repositories are named
by their GitHub name under `ChristopherVR/`; where each one is checked out on a
given machine is local knowledge: keep it in an untracked `CLAUDE.local.md`
(git-ignored), never in this file.

| Location                | npm package                            | Owns                                                                                                    |
| ----------------------- | -------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| `viewers/docx` (this)   | `docx-core`, `docx-<framework>-viewer` | Word framework bindings, package shells, docs site. Demos in `demos/docx`, browser tests in `e2e/docx`. |
| `src/core` (`ooxml`)    | `ooxml-core`                           | **All OOXML logic** by area (`docx`, `pptx`, `xml`, `opc`, `collab`, ...).                              |
| `src/ui` (`ooxml`)      | `ooxml-ui`                             | The shared `office-ui-*` elements and the `<docx-editor>` element (`src/ui/src/docx`).                  |
| `viewers/pptx` etc.     | `pptx-*-viewer`, ...                   | The other viewers, over the same core.                                                                  |
| `ole2` (own repository) | `@christophervr/ole2`                  | Legacy binary formats (DOC, XLS, PPT, CFB, RC4/MD4), inlined into `ooxml-core`.                         |

`ooxml-core` and `ooxml-ui` resolve to the workspace, so there is no `file:`
linking: edit the core or the UI and rebuild what the demo reads. Some packages
read another package's `dist`, so build it (or `bun run build:packages` here)
and suspect a stale `node_modules/.vite` before your code when one demo
disagrees with another. The release flow rewrites the workspace ranges to caret
ranges on the versions released together, so a change that needs both the core
or UI and this viewer lands in one push (see the root `AGENTS.md`).

### Where does my change go?

| The change is about...                                                              | Make it in                                        |
| ----------------------------------------------------------------------------------- | ------------------------------------------------- |
| Parsing, the document model, editing commands, saving or round-trip loss            | `src/core/docx/`, with a round-trip test          |
| Pagination, line breaking, page layout, `.doc` loading, collaboration sync          | `src/core` (`docx/layout`, `docx/load`, `collab`) |
| A control shared by Word and PowerPoint (SmartArt drawing, shared ribbon controls)  | `ooxml-ui` (`src/ui`)                             |
| `.doc` / `.xls` / `.ppt` binary codecs, CFB containers                              | `ole2`                                            |
| Ribbon, dialogs, panels, keyboard, how the editor shows the model, styling, locales | `src/ui/src/docx`                                 |
| Framework wiring (props, events, lifecycle)                                         | `packages/bindings` here (all adapters at once)   |
| Docs site, packaging scripts, build wiring for the demo                             | here (`demos/docx` and `e2e/docx` at the root)    |

### GitHub Pages

`https://christophervr.github.io/ooxml/docx/` is the docs site (VitePress in
`docs/`), deployed by the repository's `.github/workflows/pages.yml` on pushes to
`main` that touch `viewers/`, `demos/`, `e2e/`, `src/ui/` or `src/core/`.
`scripts/build-pages.mjs` here also builds the demo once per framework at `/demo/`
(React), `/demo-vue/`, `/demo-angular/`, `/demo-vanilla/`, `/demo-svelte/` and
`/demo-solid/`. The Office-suite launcher at `https://christophervr.github.io/ooxml/`
embeds those URLs, so a renamed demo route must be renamed there too. The demo
follows the VitePress `vitepress-theme-appearance` key in `localStorage`.

## Commands

```bash
bun install
bun run typecheck        # tsc + svelte-check
bun run test             # vitest (unit and binding contract tests)
bun run test:browser     # Playwright against the built demo (PLAYWRIGHT_PORT, default 4180)
bun run test:scripts     # release planner, publish guards, commit checks
bun run demo             # demo dev server (add ?framework=react|vue|angular|svelte|solid)
bun run build            # production build of the demo
bun run build:packages   # build the seven publishable packages into packages/*/dist
bun run check:published  # tarballs import only allowed packages
bun run pack:smoke       # pack, install into a clean consumer, load DOCX and .doc through every package
bun run check:shared     # docx-core stays a logic-free re-export; package classification
bun run fmt              # oxfmt (tabs, single quotes, width 100)
```

These run from `viewers/docx`. From the repository root, `bun run verify` runs
what CI would for your commits (it includes this viewer's own checks when you
touch it). The docs site builds with `bun install --cwd docs && bun run --cwd docs docs:build`.

## Releasing

Releases are automated from conventional commits by the repository's shared flow
(root `scripts/release-plan.mjs`, `scripts/viewer-packages.mjs` and
`.github/workflows/release.yml`); `docs/releasing.md` here summarises what this
viewer publishes and what triggers a release, and the root `docs/releasing.md`
has the whole flow. Never run `npm publish` or push release tags by hand.

- Each published package is versioned from the commits since its own
  `<npm-name>@<version>` tag (for example `docx-core@0.3.0`). A change in
  `web-component` or `bindings` releases all six framework packages; a `docx-core`
  release re-releases them too.
- Dry-run the planner from the repository root: `node scripts/release-plan.mjs --no-npm`.

## Branching and git workflow

The repository uses **trunk-based development**: commit directly to `main`.
**Do not create feature branches unless the user explicitly asks for one**;
this overrides any default "branch before committing" assumption. Keep each
commit small and releasable, since every push to `main` can be released by the
next scheduled run. Run `bun run verify` from your own worktree before pushing
(add this viewer's `test:browser`, `build:packages`, `check:published` and
`pack:smoke` when you touch what they cover); do not push untested code.

> The working tree is sometimes shared by parallel agent sessions, and another
> session may switch the checkout underneath you. Before committing, run
> `git branch --show-current` and `git status`. To land work without moving a
> shared checkout, push `HEAD:main` from an isolated `git worktree` rather than
> `git checkout main`.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org);
the root `CONTRIBUTING.md` has the full rules and `scripts/check-conventional-commits.mjs`
enforces them. The format is load-bearing: the type sets the version bump
(`feat` minor, `!` or `BREAKING CHANGE:` major, anything else patch), and the
**paths** a commit touches decide which packages release (tests, docs, demos
and CI never release anything).

- Scope is the package or area: `core`, `web-component`, `bindings`, `react`,
  `vue`, `angular`, `svelte`, `solid`, `vanilla`, `demo`, `release`, `ci`,
  `deps`, `docs`.
- Subject: imperative, lower-case, no trailing period, header at most 72
  characters.
- Write multi-line messages with a real heredoc or `git commit -F <file>`;
  never wrap them in a PowerShell here-string under bash, where the stray `@`
  leaks into the subject.
- Never add a Codex co-author trailer. Never include an
  AI chat share link in a commit, PR, changelog, comment or doc.

## Style

- **No em-dashes.** Never write U+2014 in source, comments, docs, commit
  messages or UI copy; use a colon, comma, semicolon, parentheses or a spaced
  hyphen. The only exception is content that intentionally renders or asserts
  that character.
- Formatting is `oxfmt`; run `bun run fmt` before committing.
- Record extraction provenance (source repository, path, commit) when code moves
  between repositories.
