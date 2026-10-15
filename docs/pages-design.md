# GitHub Pages design system

The one visual system for everything served from `https://christophervr.github.io/ooxml/`: the
OOXML Office suite (`site/`, `apps/office-suite/`), the documentation hub (`site/docs/`) and the
five product documentation sites (`viewers/<product>/docs/`). The suite and the docs are different
kinds of page (an application shell and reading pages), so they share tokens, components and rules
rather than one stylesheet.

## Principles

- **Colour comes from the product.** The neutrals are the same everywhere; the only saturated
  colour on a page is its product's accent (or the suite's blue). Product colours never mix on one
  surface except in the app tiles that stand for the products.
- **Structure is information.** Hairline-joined grids, header bars and tags encode grouping, kind
  and status. Nothing is numbered unless it is a sequence (steps, releases).
- **Text is sentence case.** Labels, buttons, table headers, kickers and tags are set in the UI
  face in sentence case. Monospace is for code, file paths and commands only. No label is set in
  tracked-out capitals, and no arrow glyph is appended to link or button text.
- **Motion answers an action.** Hover and open states transition in 140 to 220 ms. Pages do not
  animate on scroll. `prefers-reduced-motion` turns transitions off.
- **Dark first in the app, system in the docs.** The suite opens in dark mode on a first visit
  (light stays one click away in Appearance). The docs follow the reader's system setting.

## Tokens

Light and dark share the structure: a canvas, a raised surface, a sunken field, one line colour,
the ink and two softer inks. The values are the suite's own neutrals.

| Token    | Light     | Dark      | Used for                                   |
| -------- | --------- | --------- | ------------------------------------------ |
| canvas   | `#f7f8fa` | `#181a1e` | page and app background                    |
| surface  | `#ffffff` | `#202228` | cards, panels, the sidebar, the editor     |
| sunken   | `#eceef2` | `#121418` | search fields, code backgrounds in the app |
| line     | `#e3e6eb` | `#34383f` | every border and divider (1 px, no shadow) |
| ink      | `#242424` | `#f0f0f0` | text                                       |
| muted    | `#616161` | `#b5b5b5` | secondary text                             |
| faint    | `#8a8a8a` | `#8d8d8d` | placeholders, disabled, decorative strokes |
| ok       | `#16856a` | `#4fc08a` | the storage dot, saved states              |
| danger   | `#c4314b` | `#ff8585` | errors                                     |
| selected | `#e7eef8` | `#29384c` | the active nav item, selected rows (blue)  |

Accents, light then dark. The docs sites use their product's accent for `--vp-c-brand-*` and for
the landing tokens; the suite uses the Office blue and lets the user pick another.

| Product      | Light     | Dark      |
| ------------ | --------- | --------- |
| Office/suite | `#0f6cbd` | `#479ef5` |
| Word         | `#185abd` | `#5b8fe6` |
| Excel        | `#107c41` | `#3fb36b` |
| PowerPoint   | `#c43e1c` | `#e86a40` |
| Visio        | `#3955a3` | `#7b93dc` |
| Teams        | `#6264a7` | `#9a9cf0` |

Radii: 4 px on controls, 8 px on cards, panels and framed grids, 10 px on floating menus, 50% on
avatars. Shadows only on floating layers (menus, dialogs): `0 12px 45px #0003, 0 2px 8px #0001`.

## Type

- **UI and body:** `'Segoe UI Variable Text', 'Segoe UI', system-ui, Arial, sans-serif` in the
  suite; the VitePress default stack in the docs. 13 px in the app chrome, 16 px in reading pages.
- **Display:** the suite uses the same family at 600 (`Segoe UI Variable Display`), 30 px for the
  page title. The docs keep `Bricolage Grotesque` for `h1` to `h3` and landing headlines, weight
  640 to 700, letter-spacing -0.02 em, line-height 1.05 for headlines.
- **Code:** `IBM Plex Mono` in the docs, `ui-monospace, 'Cascadia Mono', Consolas` in the suite.
  Code blocks sit on a dark card in both modes.

## Components

**Joined grid.** Tiles share one hairline: `display: grid; gap: 1px; background: line; border: 1px
solid line; border-radius: 8px; overflow: hidden`, each cell `background: surface`. Hover tints the
cell with the accent at 8%. Used for the suite's create tiles, the docs' feature cards, the app
launcher and the start chooser.

**Tool card.** One cell of a joined grid that stands for an app, a package or a feature: a 36 px
icon tile top-left (the product colour at 14% as background, at 42% as a 1 px border, the glyph in
the product colour; in the suite the app tiles stay the solid product-colour squares with a white
letter), a 13.5 to 16 px semibold title, one line of muted description, an optional tag beside the
title and a 22 px hairline box with a plus or arrow glyph top-right that takes the accent on hover.

**Section bar.** A section opens with a row: the section title left (display face), a short muted
note right, a hairline below. No kicker above it.

**Tag.** `Alpha`, `Beta`, `New`: 10 px semibold, 1 px border in the tone colour at 40%, text in
the tone colour, 3 px radius, 1 px 6 px padding. Tones: alpha amber `#b07a17`, beta blue (the
accent), new green (`ok`).

**Buttons.** Primary: accent background, white text (dark ink on dark-mode accents), 12 px
semibold, 8 px 14 px padding, 4 px radius. Secondary: transparent with a `line-strong` border.
Text link: accent colour, underline on hover.

**Header bar.** 52 px, canvas background, hairline below. Brand left, search centred (sunken field,
`/` focuses it), actions right. The docs keep the VitePress nav and add two links: "OOXML Office"
(the suite) and "All docs" (the hub).

**Sidebar.** 228 px, surface background, hairline right. Group captions are 11 px semibold muted
sentence case. The active item has the selected tint, accent text and a 3 px accent bar on the
left. Items that are previews carry a tag.

## The pages

- **Start chooser** (`site/`, before the suite loads): one joined list with a row per product, a
  row for the whole suite and a row for the documentation hub.
- **Suite home**: a headline that asks what to work on, the create tiles as tool cards (one per
  app plus "Open a file"), the library below it in a framed table. The sidebar groups Workspace,
  Apps and Resources (Documentation, Product sites and demos; hidden in an installed app, which has
  no browser chrome to come back through).
- **Documentation hub** (`site/docs/`): one page that links every manual: the core library, the UI
  package, the MCP server, the suite, the desktop build and each product site, as tool cards.
- **Product docs landing**: hero with the product frame, features as tool cards, status, quick
  start, FAQ, finale. The accent is the product colour.

## Do not

- Do not introduce a second palette, a paper tint, a scanline or a grain texture.
- Do not set labels in monospace capitals, add kickers above headings, or append arrows to text.
- Do not add drop shadows to cards or screenshots; frame them with the line colour.
- Do not name the external sites this direction was compared against, in code, docs or commits.
