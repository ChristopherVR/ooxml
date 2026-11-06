# Font loading

How the products choose the face for each run of text, when they download one, and how a host or
a reader configures that. The network side is summarised in [PRIVACY.md](../PRIVACY.md).

## Word, Excel and Visio

These use the fonts the device has. A family the document names is written first in the CSS
`font-family` list, followed by the substitutes in `ooxml-core/text` (`font-substitution-map.ts`,
for example Calibri, then Carlito, Liberation Sans, Arial) and a generic family. Nothing is
downloaded, so a missing font is drawn by the first installed substitute and line breaks can
differ from Office.

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

`src/ui/src/pptx/render/google-webfonts.ts` builds one Google Fonts CSS2 stylesheet URL for the
families a deck needs:

- Which families Google serves comes from a bundled catalogue (`google-fonts-catalogue.ts`,
  regenerated with `bun run fonts:catalogue`), never from probing the API. So only known families
  are requested, and the names of private or corporate fonts never leave the device.
- A family Google does not serve under its own name is requested as its verified
  metric-compatible clone (`google-webfonts-metric-clones.ts`: Calibri and Aptos as Carlito,
  Cambria as Caladea, Arial as Arimo, Times New Roman as Tinos, and so on), so line breaks match.
- A weight named as its own family ("DM Sans Light") requests the base family with that weight.
- Fragments are sorted, so the same deck always produces the same URL and the stylesheet is
  fetched once.

### Waiting for fonts before showing a deck

Each binding calls `prepareDeckWebfonts` (`deck-webfont-readiness.ts`) between parsing a deck and
showing it, so slides do not appear in a stand-in face and then visibly change:

- the managed `<link>` is injected as soon as the deck is parsed;
- the loading state stays up until each face the slides draw has loaded through `document.fonts`
  (with the characters they use, so the right `unicode-range` subset loads);
- after 400 ms the loading message shows "Downloading fonts (n of m)";
- the wait ends after 8 seconds, or at once when the stylesheet request fails (offline,
  blocked, refused by a Content Security Policy); the deck is then shown with substitutes and a
  face that arrives later swaps in.

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

Every PowerPoint binding (React, Vue, Svelte, Angular, vanilla) takes `googleFonts`:

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

### OOXML Office suite: Settings > Fonts and privacy

The suite's Settings dialog chooses for every editor it opens:

- **Download from Google Fonts** (default): passes no `googleFonts`, so each editor's own Trust
  Center choice still applies;
- **Download from my own font server**: passes `{ baseUrl }`;
- **Do not download fonts**: passes `false`.

The choice is kept in `localStorage` (`ooxml-suite-fonts`) and applies to documents opened after
it changes. The site's Content Security Policy (`site/index.html`, `site/404.html`) allows
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
