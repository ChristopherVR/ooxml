# Font loading

How the products choose the face for each run of text, when they download one, and how a host or
a reader configures that. The network side is summarised in [PRIVACY.md](../PRIVACY.md).

## Word, Excel and Visio

These use the same webfont fallback as PowerPoint (below), from the same shared code, but it is
**off unless the host turns it on** with the `googleFonts` option. With it off they use the fonts
the device has and the file embeds. Either way, a family the document names is written first in
the CSS `font-family` list, followed by its verified metric-compatible clone and the substitutes
in `ooxml-core/text` and a generic family, in every place that measures or paints document text:

- **Word** (`<docx-editor>`): `cssFontStack` (`metricFontFamily`) in Print Layout, Draft and Web
  Layout runs, headers, footers and the watermark. Before the first layout, `load()` waits
  (bounded) for the faces the document draws; faces that arrive later relay the pages out.
  Word's loader does not register embedded fonts, so none count as available.
- **Excel** (`<xlsx-editor>`): `workbookFontFamily` (`cssFontFamily`) for cells and rich runs,
  autofit, automatic row heights, shapes and text boxes, Page Layout, print, headers and
  footers. `load()` waits before showing the workbook; faces that arrive later re-measure and
  repaint the grid without changing stored row heights or marking the workbook dirty.
- **Visio** (`<visio-viewer>`): `getSubstituteFontFamily` for shape text, bullets, the in-place
  text editor and the wrap measurement; exported SVG keeps the original family first. PDF export
  is unchanged. `load()` waits before showing the drawing; later faces re-render the page.

The fonts a file uses are collected by `collectDocxWebfontRuns` (`ooxml-core/docx/layout`),
`collectWorkbookWebfontRuns` (`ooxml-core/xlsx`) and `collectVisioWebfontRuns`
(`ooxml-core/visio`). A font chosen while editing is only requested on the next load.

Why off by default: unlike PowerPoint, which has always downloaded fonts, these products made no
network request at all, and a new request that names fonts to a third party must be chosen by
the host or the user (`AGENTS.md`). The OOXML Office suite turns it on from Settings > Fonts.

## PowerPoint

PowerPoint decks lean on fonts the reader often lacks (Aptos and other Microsoft 365 cloud fonts,
template display faces), and slide layout is sensitive to their metrics, so the PowerPoint viewer
resolves each family in this order:

1. **Embedded in the deck.** Embedded fonts (obfuscated `.fntdata` / `.odttf`, or EOT, which
   `mtx-decompressor` unpacks) are de-obfuscated by the core loader and registered as
   `@font-face` rules from `data:`/`blob:` URLs. A family whose embedded files do not match their
   slot (a SemiBold file in the Light slot) is checked against the file's own weight and italic
   flags (`embedded-font-trust.ts`) and, when Google Fonts serves that family, downloaded instead.
2. **Installed on the device.** Detected by measuring text on a canvas
   (`isFontFamilyInstalledLocally` in `google-webfonts.ts`); `document.fonts.check` is not used
   because Chromium answers `true` for unknown families. An installed family is never requested.
3. **Added by the reader.** File > Options > Custom fonts registers a `.ttf`, `.otf`, `.woff` or
   `.woff2` file for the session (`ooxml-ui/fonts/custom-fonts.ts`). It is not stored.
4. **Downloaded** (the webfont fallback, below), when allowed.
5. **Substituted.** The CSS list continues with the metric-compatible clone and the substitutes
   from `ooxml-core/text`, then a PANOSE-based generic family.

### The webfont fallback

`ooxml-core/text` (`text/webfonts/`) builds one Google Fonts CSS2 stylesheet URL for the
families a document needs, and `ooxml-ui`'s `fonts/webfonts.ts` injects it and waits for the
faces. PowerPoint reaches them through `src/ui/src/pptx/render/google-webfonts.ts`:

- Which families Google serves comes from a bundled catalogue (`google-fonts-catalogue.ts`,
  regenerated with `bun run fonts:catalogue`), never from probing the API. So only known families
  are requested, and the names of private or corporate fonts never leave the device.
- A family Google does not serve under its own name is requested as its verified
  metric-compatible clone (`metric-clones.ts`: Calibri and Aptos as Carlito,
  Cambria as Caladea, Arial as Arimo, Times New Roman as Tinos, and so on), so line breaks match.
- A weight named as its own family ("DM Sans Light") requests the base family with that weight.
- Fragments are sorted, so the same deck always produces the same URL and the stylesheet is
  fetched once.

### A short wait before showing a document

**Starting early.** As soon as a file's bytes arrive, before it is parsed, every product reads
the font families its package declares from a few small parts (themes, Word's font table and
styles, Excel's style sheet, PowerPoint's presentation part, masters and first slides, Visio's face
names; `readDeclaredFonts` in `ooxml-core/opc`, inflating only those parts) and injects an early
stylesheet for the ones neither embedded nor installed (`prefetchWebfonts` in
`ooxml-ui/fonts/webfonts.ts`, PowerPoint's `prefetchDeckWebfonts`, link id `<id>-prefetch`). The
CSS downloads while the parse runs, so the faces are usually ready by the time the short wait
below begins. The browser only downloads font files for faces the page draws. The early link is
kept for the open file and removed with the product's own link when downloads are turned off.

Each PowerPoint binding calls `prepareDeckWebfonts` (`deck-webfont-readiness.ts`), and Word,
Excel and Visio call `prepareWebfonts` (`ooxml-ui/fonts/webfonts.ts`), between parsing a file and
showing it. The wait is deliberately short (`WEBFONT_TIMEOUT_MS`, 150 ms, about the browser's own
block period for `font-display: swap`): cached faces arrive inside it, and a first download on a
slow network does not hold the document back:

- the managed `<link>` is injected as soon as the deck is parsed (for PowerPoint, as soon as core
  has parsed the slides, `PptxHandlerLoadOptions.onSlidesParsed`, while the rest of the deck still
  loads and its media resolves);
- the loading state stays up until each face the slides draw has loaded through `document.fonts`
  (with the characters they use, so the right `unicode-range` subset loads);
- the wait ends after 150 ms, or at once when the stylesheet request fails (offline, blocked,
  refused by a Content Security Policy); the file is then shown with its metric-compatible
  substitutes, and a face that arrives later is laid out again in place (each product listens
  for `document.fonts` `loadingdone` and keeps the selection, scroll and any edit).

PowerPoint also says what it is waiting for. A download still running when the deck is shown
keeps being reported (`followMs`, up to 8 seconds): from 400 ms the status bar shows
"Downloading fonts (n of m)", with the families in its tooltip, and the item goes away when
every face has loaded or failed, or the download is given up.

Installed and embedded fonts need no wait, and with downloads off there is none.

### Other uses of downloaded fonts

- **WordArt warps** (inflate, can, waves and the like) bend real glyph outlines, so for those runs
  only the font files are also fetched with `fetch()` (`deck-webfont-outlines.ts`). Without them
  the viewer traces the browser's own rendering.
- **Slide image export** inlines the faces into the exported SVG so the canvas is not tainted
  (`foreign-object-font-embed.ts`).

Both reuse the stylesheet URL above, so they follow the same setting. Both use `fetch()`, so a
mirror must send `Access-Control-Allow-Origin`.

## Configuration

### Host: the `googleFonts` option

`<docx-editor>`, `<xlsx-editor>` and `<visio-viewer>` take the same values as a property; unset
means off for them. Every PowerPoint binding (React, Vue, Svelte, Angular, vanilla) takes `googleFonts`:

| Value                                           | Effect                                                                                                        |
| ----------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| unset (default)                                 | The reader's Trust Center choice decides; it is on by default and requests `fonts.googleapis.com`.            |
| `false`                                         | Nothing is downloaded. The Trust Center toggle is shown off and locked.                                       |
| `{ baseUrl: 'https://fonts.example.com/css2' }` | The same requests go to a self-hosted mirror of the CSS2 API. The Trust Center toggle is shown on and locked. |

A `baseUrl` must be an absolute `http(s)` URL without credentials, query or fragment; anything
else turns downloads off and logs a warning, rather than falling back to Google.

```ts
createPptxViewer(container, { googleFonts: false });
createPptxViewer(container, { googleFonts: { baseUrl: 'https://fonts.example.com/css2' } });
```

### Reader: File > Options > Trust Center

"Download fonts that are not installed or embedded (Google Fonts)" (`trust.allowFontDownloads`, on by default)
lets the reader turn downloads off. Turning it off removes the open deck's webfont stylesheet at
once and applies to every later load. The choice is saved with the other viewer preferences, and
a host `googleFonts` value overrides it and shows the toggle read-only. The rule lives once in
`src/ui/src/pptx/render/options/font-download-policy.ts`, which every binding uses.

Word, Excel and Visio have the same toggle in their own File > Options > Trust Center, **off by
default**: the reader turning it on is what allows the request when the host set nothing. A host
`googleFonts` value overrides it and shows it read-only, with the "managed by your organization"
reason. Turning it on loads the open file's webfonts and lays it out again; turning it off
removes both of the product's stylesheets at once. The rule is `effectiveFontDownloads` and
`fontDownloadLock` in `ooxml-ui`'s `fonts/font-download-policy.ts`.

### OOXML Office suite: Settings > Fonts

The suite's Settings dialog (the shared `office-ui-options-dialog`, as PowerPoint's File >
Options) chooses for every editor it opens:

- **Download from Google Fonts** (default): PowerPoint gets no `googleFonts`, so its own Trust
  Center choice still applies; Word, Excel and Visio get `{ baseUrl: 'https://fonts.googleapis.com/css2' }`,
  since unset leaves it to their Trust Center toggle, which is off by default (the suite's value
  locks that toggle);
- **Download from my own font server**: passes `{ baseUrl }`;
- **Do not download fonts**: passes `false`.

The choice is kept in `localStorage` (`ooxml-suite-fonts`) and applies to documents opened after
it changes. A custom address that is not an absolute http(s) URL without credentials, query or
fragment is refused in place, and the previous choice holds until it is valid. Settings >
Privacy lists what the current choice contacts. The site's Content Security Policy (`site/index.html`, `site/404.html`) allows
`https://fonts.googleapis.com` in `style-src` and `https://fonts.gstatic.com` in `font-src`; a
deployment that uses its own font server adds it to both (and to the desktop policy in
`desktop/src-tauri/tauri.conf.json`).

### Self-hosting a mirror

The mirror must answer the CSS2 query format (`?family=Name:ital,wght@0,400;0,700&display=swap`,
one `family` parameter per family) with `@font-face` CSS whose `src` URLs it also serves, and
send CORS headers for the outline and export fetches. Serving the font files from the same
mirror keeps every request off Google. Families outside Google's catalogue (Selawik for Segoe UI,
for example) can be served by a mirror, since the viewer still lists them first in the
`font-family` chain.

## Documentation sites

The VitePress sites bundle their own fonts (`@fontsource-variable/bricolage-grotesque` and
`@fontsource/ibm-plex-mono`, imported in each `.vitepress/theme/index.ts`), so a page view
contacts no font service.
