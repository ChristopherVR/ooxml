# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others) working in this
repository. This file is canonical: `CLAUDE.md` only imports it, so edit this
file and never fork the two (the viewer repositories drifted that way once).

## Working agreements

- This repository is the **single published package `ooxml-core`**. It owns **all the logic** of the Office products: OOXML (ECMA-376 / ISO 29500) packaging, XML, WordprocessingML, PresentationML and later SpreadsheetML models, parsing and serialization, DrawingML, charts, diagrams (SmartArt), maths, geometry, layout, editing commands, validation, and collaboration (Yjs and the sync protocol). It is a sibling of `ole2`.
- The viewer repositories (`docx-viewer`, `pptx-viewer`, later `xlsx-viewer`) own **only the UI**: web components, framework bindings, ribbon and dialog views, styling, demos and end-to-end tests. They consume this package and must not keep, copy or fork logic that belongs here. Where a viewer still holds logic (today: the Word model, parser, serializer, editing and layout live here in `docx`, and the PowerPoint engine in `pptx`; the pptx-viewer `shared` render logic and the Word editor's view code still live in the viewers), `docx-viewer/docs/ooxml-core-plan.md` says when it moves.
- `ole2` owns the legacy compound-file and binary formats (DOC, XLS, PPT, CFB, RC4/MD4). Never move modern OOXML into `ole2`, and never move binary codecs here. The encrypted-package container is CFB (ole2); the encryption primitives are modern OOXML (here).
- One shared XML model and one structure for every Office type. Parsers and writers are written once against it; type-specific code lives in its own area (`docx`, `pptx`, `xlsx`) of this package, shared code in `xml`, `opc`, `drawingml`, `chart`, `diagram`, and so on.
- Layout: `src/<area>/` with its own `index.ts`; each area is a subpath export (`ooxml-core/xml`) and the root entry groups them by namespace. Do not add a second logic package; add an area. The only other package is the DOM-only `ooxml-ui` in `packages/ui` (see `docs/ooxml-ui-plan.md`); core never imports it.
- `site/` is the static Office-suite launcher deployed to GitHub Pages by `.github/workflows/pages.yml` (no build step, not published to npm). It embeds the demos the viewer repositories deploy to their own Pages sites and must hold no Office logic; add an app there by editing `site/apps.js`.
- Bun, TypeScript strict (including `exactOptionalPropertyTypes` and `noUncheckedIndexedAccess`), Vitest with tests next to the code, ESM output (add CJS when a consumer needs it).
- Keep source modules under 300 lines where practical. Add regression tests for parsing, preservation, round-trip and editing behaviour; move tests with the code they cover.
- Unsupported features must be reported honestly. Never claim Office parity or lossless export without evidence.
- Record extraction provenance (source repository, path, commit, what changed) in `PROVENANCE.md` for every module moved here.
- Releases are automated from conventional commits: `scripts/release-plan.mjs` versions the package from the commits since its last `ooxml-core@<version>` tag, and `.github/workflows/release.yml` tags, writes the changelog and publishes via npm trusted publishing (OIDC, provenance, no token). See `docs/releasing.md`. Never run `npm publish` or push release tags by hand; no commits to published history; never publish fixtures or stale output.
- Status: areas `units`, `color`, `geometry`, `xml`, `opc`, `diagram`, `docx` and `collab` are populated (`diagram` is the format-neutral SmartArt part model; see `docs/agnostic-core-plan.md` for the migration order of `drawingml`, `chart` and the rest of the diagram code). `@christophervr/docx-core` is a thin re-export of `ooxml-core/docx` (and `/docx/embedded`); thin viewer-side packages reference only their own area.
- **The `pptx` area is compiled with relaxed flags (temporary).** It was moved from `pptx-viewer` and predates this repository's strict settings. `tsconfig.pptx.json` extends the base but turns off `exactOptionalPropertyTypes`, `noUncheckedIndexedAccess` and `noImplicitOverride`; `src/pptx` is excluded from the strict base project (`tsconfig.json`, `tsconfig.build.json`). `bun run typecheck` runs both. Tighten the pptx flags gradually, one directory at a time; new code elsewhere stays strict, and code moved out of `src/pptx` into a shared area must compile under the strict project.
- The pptx area is bundled with tsup (ESM `.mjs` + CJS `.cjs`, legacy codecs from `@christophervr/ole2` inlined) and declarations from tsdown (`bun run build` runs the tsc declarations, both tsup bundles and the tsdown declarations in parallel via `scripts/build.mjs`; `prepack` only builds when `dist` is missing). Its subpaths are `./pptx`, `./pptx/converter`, `./pptx/cli`, `./pptx/signature-node`; the root entry does not re-export it, so importing another area never pulls the pptx bundle into the strict project. The pptx area keeps its own XML model (`fast-xml-parser` object trees); unifying it with the shared `xml` area is the next step.
- pptx tests read decks from `src/pptx/__tests__/fixtures`, including a committed snapshot of `pptx-viewer/e2e/fixtures` under `fixtures/e2e` (the generator scripts stay in pptx-viewer). CI never reads from pptx-viewer. The viewer's Playwright setup regenerates its own copies; refresh the snapshot here deliberately, never automatically.
- `collab` is the format-neutral collaboration area (Yjs, `y-protocols` and `lib0` are real dependencies). It must stay free of DOM, UI frameworks, ProseMirror and product document mapping: products implement `DocumentAdapter` and keep their model-to-Yjs code in the viewers until it moves into the product area. See `docs/collab-area.md`.

## Where things live

Repositories are named by their GitHub name under `ChristopherVR/`. Where each
one is checked out on a given machine is local knowledge: keep it in an
untracked `CLAUDE.local.md` (git-ignored), never in this file.

| Repository         | npm package                                                          | Owns                                                                                                                                                                                      |
| ------------------ | -------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ooxml` (this one) | `ooxml-core`, `ooxml-ui`              | **All OOXML logic**, by area under `src/`, plus the DOM-only `packages/ui` and the Office-suite launcher in `site/`.                                                                      |
| `pptx-viewer`      | `pptx-*-viewer`, `pptx-viewer-core`                                  | PowerPoint UI: five bindings (react, vue, angular, svelte, vanilla), the internal `shared` render logic, locales, MCP tools, demos, e2e and the docs site. Consumes `src/pptx` from here. |
| `docx-viewer`      | `@christophervr/docx-core`, `@christophervr/docx-<framework>-viewer` | Word UI: the `<docx-editor>` web component and six bindings (react, vue, angular, svelte, solid, vanilla), demos, browser tests and the docs site. Consumes `src/docx` from here.         |
| `ole2`             | `@christophervr/ole2`                                                | Legacy binary formats (CFB, DOC, XLS, PPT, RC4/MD4). A pinned dev dependency here whose codecs are inlined into the bundles.                                                              |
| `emf-converter`    | `emf-converter`                                                      | EMF/WMF rendering.                                                                                                                                                                        |
| `mtx-decompressor` | `mtx-decompressor`                                                   | MicroType Express (embedded EOT font) decompression.                                                                                                                                      |

The repository was renamed from `ooxml-core` to `ooxml`, and the packages
from `@christophervr/ooxml-core` and `@christophervr/office-ui` to the unscoped
`ooxml-core` and `ooxml-ui` (tags `ooxml-core@<version>` and
`ooxml-ui@<version>`). The `office-ui-*` custom element tags and the
`office-ui.web-control-contract` symbol are public API and keep their names.

The viewers depend on the **published** version of this package. To try a
change in a viewer before it is released, build here (`bun run build`), point
the viewer's core package at this checkout with a `file:` dependency, run
`bun install --force` there, and restore the version range before committing
in that repository. Commits to another repository follow that repository's own
`AGENTS.md`.

### Where does my change go?

| The change is about...                                                                | Make it in                                                      |
| ------------------------------------------------------------------------------------- | --------------------------------------------------------------- |
| Parsing, the document model, editing commands, saving or round-trip loss (any format) | `src/<format>/` here, with a round-trip test                    |
| Units, colour, geometry, XML, OPC, SmartArt or collaboration shared by formats        | the shared area here (`units`, `color`, `geometry`, `xml`, ...) |
| `.doc` / `.xls` / `.ppt` binary codecs, CFB containers, RC4                           | `ole2`                                                          |
| How a correctly parsed element is drawn or laid out in the PowerPoint UI              | `packages/shared` in `pptx-viewer`                              |
| Ribbons, dialogs, framework wiring, styling, demos, browser tests                     | the viewer repository                                           |
| The launcher page at christophervr.github.io/ooxml                                    | `site/` here (no logic; it embeds the viewers' deployed demos)  |

### GitHub Pages

- `https://christophervr.github.io/ooxml/` is the Office-suite launcher
  (`site/`, deployed by `.github/workflows/pages.yml` on pushes to `main` that
  touch `site/`).
- `https://christophervr.github.io/pptx-viewer/` and
  `https://christophervr.github.io/docx-viewer/` are the viewers' docs sites;
  each serves its framework demos at `/demo/` (React), `/demo-vue/`,
  `/demo-angular/`, `/demo-svelte/`, `/demo-vanilla/` and, for Word,
  `/demo-solid/`. The launcher embeds those URLs (see `site/apps.js`), so a demo
  route renamed in a viewer must be renamed there too.
- All three share the `christophervr.github.io` origin and therefore
  `localStorage`. The launcher's theme uses VitePress's
  `vitepress-theme-appearance` key, which the docx demo follows live; the pptx
  demos read `pptx-demo-theme` and `pptx-viewer-prefs` at start-up (see
  `site/theme.js`).

## Commands

```bash
bun install
bun run typecheck      # strict project and the relaxed pptx project
bun run test           # vitest
bun run build          # declarations, tsup bundles, tsdown declarations
bun run test:package   # pack and import every entry point from a clean install
bun run test:scripts   # release planner, publish guards, commit checks
bun run fmt            # oxfmt (tabs, single quotes, width 100)
```

Preview the launcher with any static server from `site/` (for example
`python -m http.server` in that folder). Locally the embedded demos are
cross-origin, so theme sync into the frames only works on the deployed site.

## Branching and git workflow

This repository uses **trunk-based development**: commit directly to `main`.
**Do not create feature branches unless the user explicitly asks for one**;
this overrides any default "branch before committing" assumption. Keep each
commit small and releasable, since every push to `main` can be released by the
next scheduled run.

> The working tree is sometimes shared by parallel agent sessions, and another
> session may switch the checkout underneath you. Before committing, run
> `git branch --show-current` and `git status`. To land work on `main` without
> moving a shared checkout, push `HEAD:main` (or use an isolated
> `git worktree`) rather than `git checkout main`.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org);
`CONTRIBUTING.md` has the full rules and `scripts/check-conventional-commits.mjs`
enforces them on pull requests. The format is load-bearing: the type sets the
version bump (`feat` minor, `!` or `BREAKING CHANGE:` major, anything else
patch), and the **paths** a commit touches decide whether it releases at all
(only published files under `src/`, the bundler configs, the manifest's
shipping fields and the licence files do; `site/`, docs, tests and CI never do).

- Scope is the area: `units`, `color`, `geometry`, `xml`, `opc`, `docx`,
  `pptx`, `collab`, `ui`, `site`, `ci`, `deps`, `docs`.
- Subject: imperative, lower-case, no trailing period, header at most 72
  characters.
- Write multi-line messages with a real heredoc or `git commit -F <file>`;
  never wrap them in a PowerShell here-string under bash, where the stray `@`
  leaks into the subject.
- End every commit message with the `Co-Authored-By:` trailer. Never include an
  AI chat share link in a commit, PR, changelog, comment or doc.

## Style

- **No em-dashes.** Never write U+2014 in source, comments, docs, commit
  messages or UI copy; use a colon, comma, semicolon, parentheses or a spaced
  hyphen. The only exception is content that intentionally renders or asserts
  that character.
- Formatting is `oxfmt`; run `bun run fmt` before committing.
