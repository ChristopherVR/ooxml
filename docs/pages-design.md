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
- **Motion is fluid and purposeful.** One easing, four durations, entrances that settle once and
  hover states that lift; see Motion below. `prefers-reduced-motion` turns it all off.
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

**Tag.** `Alpha`, `Beta`, `Stable`: 10 px semibold, 1 px border in the tone colour at 40%, text
in the tone colour, 3 px radius, 1 px 6 px padding. Tones: alpha amber `#b07a17`, beta blue (the
accent), stable green (`#12715a` light, `ok` dark). A product's tag is the label in
`docs/stability.md` and nothing else; a product without a label there carries no tag. Pages do
not announce versions or label things as new: release history lives in the changelogs, reached
through a plain "Releases" link. The one exception is the release post of a coordinated major
(below), which is reached through a "Release notes" link.

**Product strip.** A joined row of the five products under the hub's opening buttons: each cell
is a link to the product's site with its solid 30 px tile, the name, one muted word for what it
makes and its tag in the top-right corner. Cells grow to fill each row, so a narrow window wraps
them without empty cells.

**Hero facts.** Under a product landing's install line, the status, licence and language as a
row of small hairline chips (12 px, 4 px radius, surface background); the status chip takes its
tag tone. The copy keeps them in one `kicker` string separated by `·`. The hero frame's bar
carries the product's solid 18 px tile where a window would have its controls.

**Buttons.** Primary: accent background, white text (dark ink on dark-mode accents), 12 px
semibold, 8 px 14 px padding, 4 px radius. Secondary: transparent with a `line-strong` border.
Text link: accent colour, underline on hover.

**Header bar.** 52 px, canvas background, hairline below. Brand left, search centred (sunken field,
`/` focuses it), actions right.

**Docs nav.** Every docs site has the same three entries and nothing else: `Guide` (the site's
getting-started page, or `/guide/` on PowerPoint), `OOXML Office` (the suite, opened in the same
tab) and `Releases`. Everything else (user guide, demos, packages, API, theming, roadmap) lives in
the sidebar. Locales translate the labels but keep the three entries. The appearance switch and
the GitHub link stay visible at 1280 px; nothing falls into the overflow menu.

**Docs sidebar head.** From 960 px, a docs page with a sidebar has no top bar: the head of the
sidebar (`DocsSidebarTop.vue` in each site's theme, mounted through the `sidebar-nav-before` slot)
carries what the bar held, in the hub's order: the product's letter on its accent tile and the
site name, the source link and the appearance switch as 30 px icon buttons, a search field that
opens the site's own search (`/` and Ctrl+K still work), the three nav entries as shortcut rows
and the language select. It stays pinned while the page list scrolls under it. Narrower screens
and the landing page keep the default header, so the phone menu is unchanged. The five themes
hold the same component and the same CSS block; change them together.

**Wide screens.** The sidebar stays at the left edge (272 px, 300 px from 1800 px, 320 px from
2300 px). The article and its outline are one group, centred in what the sidebar leaves: the
article grows from 688 px to 960 px with the window, the outline from 256 px to 288 px, and the
body text steps up to 17 px and 18 px at the same two widths. The hub follows the same rule, so
no docs page leaves half a large monitor empty.

**Languages.** Every docs site has the same five languages (English, French, Spanish, German,
Simplified Chinese) in the language menu (the sidebar head on wide screens, the nav elsewhere).
Each locale translates the landing copy, the nav, the sidebar and the site strings, and every
page a reader works from: getting started, the user-facing guide, architecture, API,
collaboration, framework bindings and the per-framework guides, building blocks, theming,
OneDrive, demos, preservation and releasing. These stay English in every locale, and the locale
sidebar links to the English page: generated lists (Excel's function list, PowerPoint's
customization reference) and the project trackers that change with every parity commit (parity
and roadmap pages, verification, outstanding work, research notes, the ribbon parity and UI
migration logs). On the PowerPoint site, French, Spanish and German cover the guide and the user
guide (`translationScope` in `scripts/check-docs.config.json`); its binding, core and package
reference sections are translated into Chinese only. A translated page keeps the English heading
ids (`## Titre {#english-slug}`) so links into it work from every language, links to the locale's
copy of a page when one exists (`/fr/bindings`), and keeps code blocks identical to English,
comments included. A language-menu link from an untranslated page leads to a missing locale URL,
which the site's one `404.html` (`site/404.html`, served by GitHub Pages for every missing path)
sends to the English page.

**Docs pages move too.** Each docs site mounts `Layout.vue` with `PageMotion.vue`: on every page
change the article and its outline rise 12 px and fade in; sidebar and nav links, pager links and
images answer hover; anchors scroll smoothly. All of it is off under reduced motion.

**Live demo language.** The embedded demo follows the site language: a demo that accepts
`?locale=` (Word, Excel, PowerPoint) is loaded in the locale the reader chose from the nav's
language menu. There is no second language control beside the framework switcher.

**Sidebar.** 228 px, surface background, hairline right. Group captions are 11 px semibold muted
sentence case. The active item has the selected tint, accent text and a 3 px accent bar on the
left. Items that are previews carry a tag.

## Motion

One easing and four durations everywhere; nothing else.

- **Easing:** `cubic-bezier(0.22, 1, 0.36, 1)` (a fast start that settles) for everything that
  enters, opens or moves; `cubic-bezier(0.65, 0, 0.35, 1)` only for things that leave.
- **Durations:** 120 ms press, 200 ms hover and toggles, 320 ms open, close, tab change and theme
  cross-fade, 650 ms entrances and reveals.
- **Entrances.** Above the fold, the headline, copy, actions and frame rise 16 px and fade in,
  staggered 80 ms apart. Below the fold, each section reveals once as it enters the viewport
  (IntersectionObserver, threshold 0.15): 24 px rise and fade over 650 ms, and the cells of a
  joined grid stagger 60 ms apart. Nothing re-animates when scrolling back up.
- **Hover.** Cards and tiles lift 2 px (`translateY(-2px)`) and take the accent on their border
  or corner box; links underline; primary buttons brighten. Pressing scales to 0.985 for 120 ms.
- **Open and close.** Dialogs, popovers and menus scale from 0.98 and fade over 320 ms; closing
  takes 200 ms. Tabs and the sidebar's active bar move, they do not blink.
- **Theme.** Changing the appearance cross-fades colours over 320 ms.
- **Never:** parallax, continuous loops (the start screen's slow outer ring and the one scroll cue
  on a landing hero are the exceptions),
  bounce, blur-in, or anything that moves on scroll position.

The suite implements this in `site/styles.css`; the docs in `landing.css` plus a `useReveals`
composable per site; the hub in `site/docs/docs.css`; the release post in
`site/releases/releases.css` and `releases.js`.

## The pages

- **Start screen** (`site/`, before the suite loads): a full page, not a dialog. The workspace
  sits in the centre of a ring with the five products around it, and the documentation and the
  sample files side by side beneath;
  the rings draw in, the centre settles, the products arrive one after another, and the outer
  ring turns once every two minutes (the site's one ambient motion). Narrow or short screens show
  the same buttons as a list.
- **Suite home**: a headline that asks what to work on, the create tiles as tool cards (one per
  app plus "Open a file"), the library below it in a framed table. The sidebar groups Workspace,
  Apps and Resources (Documentation, Product sites and demos; hidden in an installed app, which has
  no browser chrome to come back through).
- **Document commands in the ribbon layout**: there is no command row. The saved state,
  **Simple view**, the assistant and the More menu (Save, Rename, Download, Embedded files,
  Versions, Share in Teams) sit at the right end of the tab strip, so the tabs are the only suite
  line above an app's own title bar. On a phone the cluster takes a line of its own.
- **Simple view** (the suite's default for an open document): one 44 px bar on the surface colour, under the
  tabs, replaces each app's ribbon. Left to right: the app tile and file name (the extension
  muted), Save, undo and redo, bold, italic and underline drawn as the letters themselves (Word
  and Excel; a pressed one takes the selected tint), the slide or page in view in a hairline box,
  then the saved state, **Show ribbon** (the one outlined button), the assistant, the More menu
  and focus. Icon buttons are 30 px and muted until hovered. Nothing is added inside an editor:
  Word's and Excel's own title bars are hidden through their `title-bar` part while the bar
  stands in for them.
- **Documentation hub** (`site/docs/`): one page in three columns. A 272 px sidebar on the rail
  colour holds the brand, a search field (`/` focuses it; it searches every guide the page links
  to), three shortcuts and every manual grouped under sentence-case captions; the active item has
  the selected tint and the 3 px accent bar. The article opens with a breadcrumb, the title, a
  lead, two buttons and the product strip, then section bars: the three ways in as a joined grid, Simple view with a
  drawn (never screenshotted) app frame, an "I want to / Use" table in a framed panel, and the
  manuals as tool cards. The right column lists the sections and marks the one in view. The
  article and the outline grow with the window up to 1560 px together (1840 px from 2300 px) and
  stay centred in what the sidebar leaves. Under 1180 px the outline goes; under 820 px the
  sidebar becomes a header whose menu button opens the list.
- **Samples gallery** (`site/samples/`, built by `scripts/build-samples-page.mjs` from
  `samples/manifest.json`): the hub's page with its sidebar, search and outline, and one section
  per product. A sample is a hub card in a joined grid with a preview above it: the thumbnail
  bleeds to the card's edges at 16:9 behind a hairline, then the product's solid tile, the title,
  one or two sentences, the features as a plain list under a hairline, the file name and size, and two 30 px
  buttons (the accent "Open in Word", the outlined "Download"). The accent stays the suite's blue;
  the product colours appear only on the tiles. The page ships no script of its own.
- **Product docs landing**: hero with the product frame and the hero facts, the live demo (with its language select
  where the demo supports one), features as tool cards, status, quick start, FAQ, finale. The
  accent is the product colour. Every site keeps this order so the five landings read as one
  family.
- **Release post** (`/ooxml/releases/<name>/`, from `docs/releases/<name>.md`): the launch
  article of a coordinated major release, linked as "Release notes" from the hub (sidebar, the
  Contributing card and the footer) and from the suite's Resources group. See Release post below.

## Release post

One Markdown file per coordinated major, `docs/releases/<name>.md`, is both the post GitHub
shows and the source of the Pages article. `scripts/build-release-notes.mjs` renders it (no
dependencies; `release-markdown.mjs` is the Markdown subset, `release-highlight.mjs` the code
colouring, `release-media.mjs` the pictures) on `site/releases/template.html`, and
`build-suite.mjs` and `build-pages.mjs` call it after copying `site/`, so the page is
`releases/<name>/index.html` with its pictures in `releases/<name>/assets/` and the shared
`releases/releases.css` and `releases.js` beside it. The page is not precached by the PWA (only
the shell and `apps/` are) and needs no `ooxml-root`.

**The Markdown contract.** The renderer handles ATX headings with GitHub's slugs, paragraphs,
nested bullet and numbered lists, bold, italics, inline code, links, images, tables, fenced code,
blockquotes, GitHub alerts (`> [!NOTE]`, `TIP`, `IMPORTANT`, `WARNING`, `CAUTION`), HTML
comments (dropped) and raw HTML blocks (a line that starts with a tag, passed through to the next
blank line). Inline HTML inside a paragraph is escaped, as is every other character. The page
reads these shapes:

- `# Title`, then `**Published:** <date> | **Author:** [name](url)` (the metadata line; the date
  becomes a `<time>`), then an optional image alone in a paragraph (the hero).
- The lede is everything before the first `##`: a paragraph, a list of 4 to 6 pillars whose
  items all start with `**Lead-in:**` (drawn as a joined grid of cards), a fenced install block
  and a `> [!NOTE]` for upgraders. A release folder's optional `showcase.json` can supply a workspace
  tour with install commands and an index of named sections. When a section uses a
  `<div data-showcase="section-id"></div>` marker, the site renders the matching interactive scene.
  Each tab's `notes` names the Markdown headings it displays underneath. Keep benchmark facts in
  the Markdown for GitHub.
- Each `##` is a section (sentence case) in the table of contents, with its `###` features under
  it. A feature is short paragraphs, an optional code block and picture, and a closing
  paragraph that starts with `Learn more`.
- Code: `**file.ts**` alone on the line before a fence names it (otherwise the bar shows the
  language), `_Caption._` alone on the line after the fence captions it. TypeScript and
  JavaScript, JSON, HTML and shell get a light token colouring at build time. A `Before:` or
  `After:` line gives a code sample its labelled red or green header. Adjacent pairs sit side by
  side on wide screens.
- Pictures: `![alt](path)` alone in a paragraph, then an optional `_Caption._` line, then an
  optional stat line `**Before:** 4.8 s · **After:** 0.6 s` (any `**Label:** value` pairs joined
  by `·`; `After` takes the accent). Paths resolve relative to the Markdown file. In a release post, a path not found
  there is looked for in `docs/releases/<name>/`, so write `<name>/assets/clip.gif` (which GitHub
  resolves too). Pictures are copied beside the page with their width and height set. A GIF
  with an MP4 of the same name beside it is a recording: the GIF plays, and under reduced motion
  the MP4 shows instead as a paused `<video controls>` (with `<stem>.poster.webp|png|jpg` as its
  poster when one exists). A still image with an MP4 of the same name shows the paused video with
  controls and uses the image as its poster. A one-pass GIF with a same-name poster and no MP4 starts
  when in view for readers who allow motion. Interactive product previews can loop while visible. Their Enlarge control opens the full-size capture.
- A companion `docs/migration-<name>.md` is published as `/ooxml/releases/<name>/migration.html`.
  Relative links between the release post and its migration guide stay on Pages. Other repository
  links point to the file on GitHub.
- `## Every change, by product` holds one `<a id="changes-<area>"></a>` line followed by a
  `<details>` and `<summary>Name</summary>` and a blank line, then the area's chapter in Markdown
  (`###`/`####` and bullets) and `</details>`. The areas are `changes-core`, `changes-ui-mcp`,
  `changes-word`, `changes-excel`, `changes-powerpoint`, `changes-visio` and
  `changes-openteams`: changelogs link to these anchors, and on Pages a link to one opens the
  list and scrolls to it. Headings inside a list stay out of the table of contents.

**Article.** A 52 px header bar (brand, Docs, Releases, GitHub, the appearance switch, which
shares `ooxml-suite-theme` with the suite and the hub, dark first). The hero is centred on a
wash of the accent fading into the canvas: the five solid product tiles, the title at up to
96 px with the version in the accent, the metadata line and the hero picture in a hairline
frame. Below, the contents column (220 px, sticky, the current section marked with the
selected tint and the accent bar) sits left of the article; prose is 720 px wide and code,
pictures, the pillars and the change lists 880 px. Under 1180 px the contents fold into a box
above the article.

**Code card.** Dark in both modes: the file name as a tab with an accent underline, the
language and a Copy button (shown by the script; it announces the copy in a live region and
falls back to selecting the code).

**Callout.** The alert tone (note accent, tip `ok`, important Teams violet, warning alpha amber,
caution `danger`) as a 3 px left bar, a 32% border and an 8% tint, with a dot and the label.

**Change list.** One card per area: the area's solid tile (the suite's four-square symbol for the
core, `U` for shared UI and MCP), the name, a muted count of its bullets and a 22 px hairline
box with a chevron that turns. Open all and Close all sit above the first card.

**Motion.** The hero tiles, title, metadata and picture rise on load, staggered; every section,
figure, code card and change list reveals once (its top 15% into the view); the pillar cards
stagger 60 ms apart; an opened list's body rises over 320 ms. All of it is off under reduced
motion, which also swaps recordings for their paused videos.

## Do not

- Do not introduce a second palette, a paper tint, a scanline or a grain texture.
- Do not set labels in monospace capitals, add kickers above headings, or append arrows to text.
- Do not add drop shadows to cards or screenshots; frame them with the line colour.
- Do not name the external sites this direction was compared against, in code, docs or commits.
