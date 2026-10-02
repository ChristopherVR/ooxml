# pptx-viewer Pages design review and Visio implementation requirements

Reviewed 2026-10-02. Reference: `ChristopherVR/pptx-viewer`, branch `main`, commit `b4712950838e00ac81a09fb9595554db7d4f8158` (`fix(shared): parse gallery preview SVG instead of assigning innerHTML`).

## Recommendation

Reuse the reference's product structure: an editorial documentation home, a real framework-switchable playground, task-oriented developer/user guides, and one shared editor design system. Adapt the workspace to diagrams rather than reproducing presentation-specific panels. Make capability evidence visible from the beginning. Do not claim Microsoft Visio parity from a good-looking ribbon or a list of package exports.

The most important improvements over the reference are fail-closed demo deployment, one canonical framework registry, local/base-aware preview URLs, proper keyboard tab behavior, a real app-readiness/error state, and an explicit fidelity report. Keep the new repository private and its deployment untriggered until publication is separately authorized.

## Review scope and evidence

- Read the repository's root `AGENTS.md` and `DESIGN.md`. There were no additional `AGENTS.md` files or `.agents/skills` packages in the reference checkout.
- Inspected the VitePress config/theme, all landing components, installation/code examples, live-demo state management, shared upload screen, demo build stamps, Vite configuration, Pages workflow, known-limitations documentation, editor layout contracts, and relevant responsive/parity tests.
- Viewed the published home in the cloud browser at 1188 x 761. Switched the embedded playground through React, Vue, Angular, Svelte, and Vanilla. All five rendered the sample deck. This is a mount-and-layout check, not a claim that every editor command works.
- Verified a keyboard accessibility defect: after selecting the Vue framework tab, ArrowRight neither moved focus nor changed the selected tab. DOM inspection showed no `tabindex` or `aria-controls` on any framework tab. The landing page had no `main` element.
- Responsive conclusions below are source-based. A physical phone, touch input, narrow viewport, screen reader, reduced-motion preference, and every locale were not exercised in this review.
- Did not install dependencies, run the reference test suite, change the reference checkout, modify remotes, or deploy anything. The reference working tree remained clean.

Primary entry points:

- [Repository](https://github.com/ChristopherVR/pptx-viewer)
- [Published documentation home](https://christophervr.github.io/pptx-viewer/)
- [Published React demo](https://christophervr.github.io/pptx-viewer/demo/)
- [Source at reviewed commit](https://github.com/ChristopherVR/pptx-viewer/tree/b4712950838e00ac81a09fb9595554db7d4f8158)

## 1. What makes the reference effective

### A. It separates the marketing surface from the work surface

`docs/index.md` selects a custom `LandingHome` VitePress layout. The actual editor runs in an iframe as an independently built application. This avoids loading an entire editor into the default documentation shell and allows a visitor to open it full-screen without losing the documentation context.

The page sequence is deliberate:

1. Split hero: short product promise, explanatory paragraph, Get Started / Live Demo actions, framework install command, and a real editor recording.
2. Embedded live demo: framework picker, standalone-app link, and an optional two-app collaboration demonstration.
3. Eight compact capability cards, each linking to relevant documentation.
4. Agent/headless capability explanation with code.
5. Framework-specific quickstart code and deep links.
6. Native expandable FAQ.
7. Final adoption action and grouped product/docs/community links.

This converts a reader from "what is it?" to "can I see it?" to "how do I use it?" without forcing an install or file upload first. The sample document is crucial: a first-time visitor can verify real rendering immediately.

### B. It gives all five bindings equal visibility

React, Vue, Angular, Svelte, and Vanilla appear in the hero installer, playground, code examples, package navigation, and package-specific documentation. Visitors can compare the same sample in each implementation. A shared sample and shared chrome are much stronger evidence than five logos.

The repository rules reinforce this: format/model work belongs to `ooxml-core`; framework-neutral decisions and style descriptors belong to the internal shared package; bindings own only rendering and reactive wiring. The same rule should apply to Visio before framework adapters multiply.

### C. It uses a recognizable but restrained visual system

The documentation feels distinct from the editor. The home is warm, spacious, and editorial. The editor is dense, dark, and task-oriented. A consistent accent, monospaced utility labels, subtle border treatment, and real screenshots connect them.

Reference tokens from `docs/.vitepress/theme/landing/landing.css`:

| Role           | Light     | Dark      |
| -------------- | --------- | --------- |
| Canvas         | `#fbfaf7` | `#0f1113` |
| Surface        | `#ffffff` | `#171a1e` |
| Primary text   | `#1a1d21` | `#f0efec` |
| Secondary text | `#5a626e` | `#9aa1ab` |
| Border         | `#e6e2d9` | `#272c33` |
| Accent         | `#c2431f` | `#e86a40` |

Display font: Bricolage Grotesque. Utility/code font: IBM Plex Mono. Body uses the VitePress system stack. The Visio project can keep the typography/spacing strategy while using a blue accent and its own identity. Do not merely replace the word PowerPoint in a cloned screenshot.

### D. Documentation navigation matches two audiences

The top navigation distinguishes Developer Guide, User Guide, Packages, and Releases. Contextual sidebars then provide local navigation, while the right outline handles long pages. Local search avoids requiring a search service. Package guides separate first use from props/handles, customization, and advanced features.

For Visio, distinguish user tasks (open, navigate, inspect, edit, save) from developer tasks (install, load model, embed, customize, export). A single giant README will not scale to the requested scope.

## 2. Layout measurements worth reusing

### Documentation/landing page

- Content width: maximum 84rem; horizontal gutters `clamp(1.4rem, 5vw, 4.5rem)`.
- Hero: approximately equal columns (1.05fr / 1fr), headline `clamp(2.7rem, 6.2vw, 5.2rem)`, compact leading, maximum 33rem explanatory text.
- Hero collapses to one column below 900px; the decorative scroll cue disappears.
- Feature grid: four columns, two at 1100px, one at 560px. The 1px grid line serves as separation rather than a forest of floating cards.
- Playground: one main pane at `clamp(480px, 72vh, 760px)`. Split collaboration becomes vertically stacked below 1080px.
- Card radius: generally 6-8px. Buttons are more squared, around 3px.
- Long-form documentation uses 1.75 line height and darker code cards in both color modes.

### Editor chrome

`DESIGN.md` and `e2e/editor-first-screen-parity.spec.ts` define real geometry contracts:

- 36px title bar.
- Ribbon primary row 32px, tabs 35px, content at least 82px, 1px lower border.
- Default slide rail 180px; inspector 288px.
- 6px icon-to-label spacing; fixed 120px font-family and 64px font-size fields to avoid layout shifts.
- 4px rail row gap and 3px selected-row accent marker.
- Collapsed notes strip 25.5px; status bar 29px.

These measurements demonstrate how to prevent drift, not a requirement to force a Visio stencil browser into a PowerPoint slide rail. Visio needs its own documented measurements and shared testable hooks.

## 3. Findings and improvements to carry forward

### P0: A failed demo can still be published as a missing path

Evidence: `.github/workflows/docs.yml`, lines 62-123, 130-192. Binding and demo builds use `continue-on-error`. The bundling step omits missing outputs. `deploy` and `demo-build-status` both depend only on `build`, so the deployment can succeed while a sibling status job turns the workflow red. The comments explicitly acknowledge that skipped paths will 404.

Requirement: build and smoke-test every advertised binding before uploading the Pages artifact. Make deployment depend on an aggregate all-required-artifacts success gate. A demo link must either work or be absent with an honest status; never silently publish a 404. For this private project, do not add an automatic publication trigger as a side effect of ordinary code pushes.

### P1: Playground tabs use incomplete tab semantics

Evidence: `LandingLiveDemo.vue:43-56`, `code/InstallPicker.vue:30-43`, `LandingQuickstart.vue:23-36`. They declare `role=tab` and `aria-selected` but have no panel association, roving tab stops, or keyboard handling. The live playground's ArrowRight failure was verified.

Requirement: either implement the complete tab pattern (labelled tablist, one tab stop, ArrowLeft/Right, Home/End, `aria-controls`, corresponding labelled tabpanel) or use ordinary buttons with `aria-pressed`. The quickstart and install controls need the same treatment. Add visible focus styles and a polite copy-success/error message.

### P1: Iframe load is not application readiness

Evidence: `LiveDemoPane.vue:59` removes the loading overlay on the iframe's load event. There is no app-ready handshake, error view, retry control, or deadline state in the component. During browsing, the shell/build stamp appeared before the editor mounted; Vanilla also showed a later internal loading overlay.

Requirement: render an explicit state machine: idle, loading application, opening sample, ready, and failed. The app should send a versioned, origin-checked ready message after document rendering succeeds. Show a useful retry and standalone-app link if loading fails. Do not use arbitrary waits as proof of readiness.

### P1: Preview URLs point at production

Evidence: `useLiveDemo.ts:14` hardcodes the published production root; all framework URLs derive from it. A developer reviewing the local docs can therefore be testing a different deployed editor version.

Requirement: derive demo paths from the documentation base and current origin, with an explicit override only for documented external hosting. Build and preview the assembled artifact under `/visio-viewer/`, not just `/`. Test direct navigation and refresh for docs and demo paths.

### P1: Claims need more visible, granular evidence

The reference does have unusually detailed limitations and runtime warnings. However, the homepage's broad round-trip/accessibility statements require a much more nuanced page to understand. Visio's requested "full parity" particularly needs an evidence model, because parsing a property does not establish accurate rendering, editing, or preservation on save.

Requirement: maintain a machine-readable capability inventory with at least Parse, Render, Edit, and Round-trip states. Include Unsupported, Preserved-only, Partial, and Verified values rather than a single supported checkmark. Link each verified claim to representative fixtures and tests. Separate browser-platform limits, known format gaps, and untested claims. Surface relevant per-document warnings in the viewer and before export.

### P2: Framework metadata and examples can drift

The reference has a framework list in `useLiveDemo.ts` and another in `code/samples.ts`, alongside repeated navigation metadata. Installer, quickstart, and live demo each own independent selection state. The code samples also do not represent identical workflows: React/Vue/Angular examples request editable mode while the Svelte sample omits it; API inputs differ across bindings.

Requirement: one framework registry should drive labels, install commands, public API examples, demo routes, availability, supported versions, and docs links. Use one shared selected framework where helpful, or explicitly label controls as independent. Compile every displayed quickstart against the built public package. Show the same user workflow in every binding, including CSS import, bounded container, loading/error handling, open, change callback, and cleanup where required.

### P2: Accessibility requires explicit page and diagram structure

The published landing page had no `main` element. It has a skip link and sensible heading sequence, and the FAQ's native `details/summary` is a strength. The iframe has a descriptive title. The custom landing should keep those strengths and add a real main landmark.

For Visio, SVG/text presence alone is not sufficient diagram accessibility. Provide named pages, a keyboard-navigable shape list or tree, shape labels and data, meaningful connection summaries, selection announcements, and a text alternative for the diagram's structure. Avoid duplicate hidden toolbar controls remaining focusable when a responsive toolbar replaces them.

### P2: Animation and typography should fail safely

The hero GIF is about 1.1MB. Reduced-motion CSS suppresses entrance animations and the scroll cue, but cannot pause the GIF. `useReveals.ts` hides below-fold content until an IntersectionObserver adds classes, with no fallback in that composable if the observer is unavailable or JavaScript fails. Fonts are loaded from Google Fonts in the site head.

Requirement: use an optimized still screenshot/SVG as the default hero, optionally add user-controlled video, and respect reduced motion. Content must remain visible without JavaScript; only apply hidden reveal styles after enhancement is initialized. Self-host appropriately licensed fonts or use system fonts for a private/offline-friendly demo. Keep the statement "files stay local" distinct from a claim of zero third-party network requests.

### P2: Build-time link checking is weakened

Evidence: `docs/.vitepress/config.ts:18` sets `ignoreDeadLinks: true`. This allows a documentation build to pass despite broken internal links.

Requirement: keep dead-link validation enabled and add checks for all package/demo routes. Only use narrowly justified allowlists for intentional external or generated links. Cross-link source, compatibility, and current package version so readers can determine exactly what is deployed.

### P2: Mobile tables and embedded editor density need real checks

The source has thoughtful breakpoints, flex wrapping, and stacked demo panes. But `custom.css:202` forces documentation tables to `display: table`, removing the default block-scroll strategy. Very wide API tables could overflow small screens. At the reviewed desktop width the embedded ribbon also requires horizontal scrolling; Svelte and Vanilla exposed bright native scrollbars unlike the subdued React/Angular appearance.

Requirement: wrap wide documentation tables in a labelled horizontal scroller, and verify 320px/375px widths and 200% text zoom. On phones, offer an immediate full-app action and collapse nonessential explanation. Use a mobile-specific editor toolbar and bottom sheets; do not squeeze the desktop ribbon and both side rails into the phone viewport.

### P2: Collaboration should be an intentional, well-labelled optional mode

The reference's split view is compelling and its teardown helper sends a leave signal before destroying frames. It uses a fresh room when the host framework changes, avoiding stale document duplication. These lifecycle details are worth preserving if collaboration is implemented.

Do not ship the feature cosmetically before the transport, ownership, data sharing, teardown, and rejoin behaviors exist. Use `crypto.randomUUID()` for room identifiers, exact `postMessage` target origins, validated message sources, explicit connection state, and truthful network disclosure. Client-side rendering does not mean a collaborative document never leaves the browser. A TLS URL alone is not a trust decision for a relay.

## 4. Proposed Visio information architecture

### Home

1. Hero: concise statement of currently implemented Visio capability; two actions, Open Playground and Developer Guide. Include real diagram preview generated by the current implementation.
2. Sample playground: open a known `.vsdx` sample without uploading a file; show diagram name, active page, framework, and build identity.
3. Evidence strip: formats implemented, API/environment availability, and a link to the capability matrix. Avoid an unsupported 100%/full-parity badge.
4. Feature groups: document navigation, rendering, shape editing, data/layers, round-trip/export, embedding. Only describe available functions as available.
5. Framework selector and compiled quickstart.
6. Compatibility and privacy FAQ.
7. Documentation/source links and clear project status.

### Main navigation

- Playground
- User Guide: Open files, Pages and backgrounds, Pan/zoom, Shapes and connectors, Layers, Shape data, Editing, Save/export, Keyboard shortcuts
- Developer Guide: Install, Quickstart, Architecture, Document model, Events, Customization, Framework bindings, Worker/headless usage
- Compatibility: formats, feature matrix, runtime warnings, fixture evidence, known limits
- API and Releases

Do not create empty documentation sections for unsupported capabilities merely to make the navigation look complete. Add them with clear planned status only when useful.

## 5. Concrete Visio workspace contract

### Desktop

- Top title/status strip: document name, dirty/save state, open/save, undo/redo, command search.
- Compact command groups: File, Home, Insert, Design, Data, Review, View, plus contextual shape/text commands only when implemented.
- Left rail: searchable stencils/masters and a shape tree. Distinguish a source stencil master from an instance already placed on the page.
- Center: crisp SVG diagram stage with pan/zoom, fit page/selection, rulers/grid/guides, clear selection/resize/rotation/connection handles, visible page bounds, and predictable hit targets at small zoom.
- Right inspector: Shape, Data, and Layers tabs; position/size/rotation, text/style, layer visibility/lock, shape properties, and an optional read-only ShapeSheet/formula diagnostic panel.
- Bottom: horizontal page tabs with current page and background-page distinction, add/rename/reorder when supported, zoom controls, and compatibility-warning count.
- Unsupported content must remain identifiable and preserved where possible. Selecting it should show what is known and what editing would risk.

### Mobile and tablet

- Under an explicitly tested breakpoint, replace the ribbon with a compact toolbar and action menu.
- Page navigation, stencils, inspector, and layers become mutually exclusive sheets/drawers. Keep part of the diagram visible above a bottom sheet where practical.
- Minimum recommended touch targets 44px; handles may be visually smaller only with expanded hit areas.
- Support pinch zoom and pan without confusing them with shape drag. An explicit Select/Hand tool is useful on touch.
- Escape/Close dismisses the active panel and restores focus to its opener. Opening and closing an inspector must not reset document, selection, page, or zoom.

### Common interactions and fidelity

- A single command/state model powers all bindings. A command has a stable ID, enabled/checked state, accessible label, shortcut, and implementation.
- Expose stable `data-visio-*` hooks for viewer root, viewport, selected shape, page tabs, stencil rail, inspector, and command IDs.
- Framework switching in a sample showcase can reset the sample, but state this clearly. In the real app, protect unsaved changes before replacing a document.
- File-open errors must identify unsupported format versus malformed package versus partial feature support. Cancel/retry must preserve the previous document.
- Save/export must report whether it is native Visio serialization or a view-only export, and whether approximated/unsupported data is preserved.

## 6. Build and deployment structure

The reference builds docs separately with VitePress, builds foundation packages first, then bindings, then one Vite application per framework, and copies demo outputs into the documentation artifact. Each demo receives an explicit deployment base and a package-version/commit/date build stamp. Those are sound techniques.

Suggested structure when the project grows into multiple packages:

```text
packages/core/       Public facade over the Visio engine in ooxml-core
packages/shared/     Commands, state decisions, geometry-to-view descriptors, shared chrome
packages/react/      Thin React binding
packages/vue/        Thin Vue binding
packages/angular/    Thin Angular binding
packages/svelte/     Thin Svelte binding
packages/vanilla/    Plain DOM binding
demos/demo-*/        Independent apps using built package exports
docs/.vitepress/     Docs shell and landing components
docs/compatibility/  Capability inventory and evidence
e2e/                 Framework-neutral interaction and visual tests
```

Do not reorganize a working small implementation purely to mimic the reference. Establish clear engine/shared/binding boundaries now, and extract packages when justified. The important contract is shared behavior, public-API testing, and equal advertised capability, not directory count.

For future authorized Pages publication:

1. Install from a committed lockfile with a pinned toolchain.
2. Lint, type-check, and unit-test the final code.
3. Build all required packages and public API examples.
4. Build documentation with link validation and all advertised demos using repository-prefixed bases.
5. Assemble one clean artifact, with no private fixture files, secrets, or unintended source maps.
6. Serve that artifact locally under the exact deployment prefix and run route, runtime, accessibility, and responsive smoke checks.
7. Only upload/deploy after the aggregate gate succeeds and publication is authorized.
8. Verify the deployed build identity and each advertised deep link before calling it complete.

Private repository status is not itself proof of private website access. Confirm hosting visibility and audience separately before enabling a Pages environment.

## 7. Acceptance checklist

### Documentation and landing

- One main landmark, one H1, logical heading order, functional skip link.
- All tabs keyboard-operable; copy action announces success/failure.
- 320/375/768/1024/1440px checks, 200% text zoom, light/dark, reduced motion.
- No whole-page horizontal overflow; wide API tables/code scroll locally.
- No invisible content when enhancement fails; diagram preview has meaningful alt text.
- Correct install/example/docs link for every framework, with each example compiled.
- All internal links checked; no falsely advertised published npm package or unavailable demo.

### Playground and viewer

- Each advertised binding opens the same sample with the same page/shape counts.
- Visible build identity matches the tested commit/package version.
- Load, error, retry, cancel, second open, empty document, malformed document, large diagram.
- Page switch, pan, zoom, fit, select, inspect, layer toggle, undo/redo and supported edit/save operations work consistently.
- Rapid framework switches never allow a stale load to win.
- Closing/reopening a panel preserves page, selection and zoom; focus returns correctly.
- Keyboard shortcuts do not fire while editing text or using a dialog.
- Hidden responsive controls are removed from the accessibility/focus tree.
- Connection handles remain operable at different zooms; long shape data and labels wrap/scroll without destroying the workspace.

### Evidence and delivery

- Capability status distinguishes implemented, preserved, approximated, unsupported, and untested.
- A corpus of genuine Visio-authored documents is compared with exported/reference output before fidelity claims.
- Round-trip assertions verify package structure, relationships, unknown parts and edited values, plus real Visio reopen evidence when available.
- No deployment or sharing change is implied by local implementation completion.

## Source map

- Overall app and repo boundaries: `AGENTS.md`, `DESIGN.md`, `CONTRIBUTING.md`.
- Site navigation/base/search/locales: `docs/.vitepress/config.ts`.
- Page composition: `docs/.vitepress/theme/landing/LandingHome.vue`.
- Hero/layout/tokens: `LandingHero.vue`, `landing.css`, `docs/.vitepress/theme/custom.css`.
- Playground/state/lifecycle: `LandingLiveDemo.vue`, `LiveDemoPane.vue`, `useLiveDemo.ts`.
- Installer and examples: `landing/code/InstallPicker.vue`, `landing/code/samples.ts`, `LandingQuickstart.vue`.
- Accessibility/reveals/FAQ: `useReveals.ts`, `LandingFaq.vue`.
- Demo upload screen and identity: `demos/shared/dropzone.css`, `demos/build-stamp.ts`, `demos/demo-react/vite.config.ts`.
- Deployment: `.github/workflows/docs.yml`.
- Honest capability model: `docs/guide/limitations.md`, `docs/guide/customization.md`.
- Useful regression-test patterns: `e2e/editor-first-screen-parity.spec.ts`, `e2e/chrome-shell-parity.spec.ts`, `e2e/inspector-responsiveness.spec.ts`, `e2e/mobile-audit.spec.ts`.

## Later CI health check

A later [main CI run](https://github.com/ChristopherVR/pptx-viewer/actions/runs/37014286048) failed on 2026-10-02 while Pages deployment succeeded independently. Reported failures included a Vanilla heap exhaustion case and browser-shard failures across all five bindings, with print/share/custom-control interaction gaps. Deployment success is therefore not a substitute for application health. The Visio project keeps independent common-state, cross-binding, resource and accessibility checks, and has no deployment workflow enabled.
