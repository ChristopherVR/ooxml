# Browser support

The editors and viewers support the last two major versions of Chrome, Edge,
Firefox and Safari. The table gives the oldest versions the shipped code and
styles can run on at all, which sit below that policy:

| Browser                | Minimum | Why                                                                                                                                                |
| ---------------------- | ------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| Safari (macOS, iPadOS) | 16.4    | Tailwind CSS v4 (cascade layers, `@property`, `color-mix()`), regular-expression lookbehind, `CompressionStream` and the unprefixed Fullscreen API |
| Chrome, Edge           | 111     | Tailwind CSS v4                                                                                                                                    |
| Firefox                | 128     | Tailwind CSS v4                                                                                                                                    |

Below Safari 16.4 the bundles fail to parse (lookbehind is a syntax error), so
older iPads load nothing. Supporting them would mean replacing Tailwind v4 as
well as rewriting those expressions, so it is out of scope.

## Safari and iPadOS notes

Behaviour that differs on WebKit and how the code handles it (issue #290):

- **Symbol fonts.** iPadOS has no Wingdings or Symbol. Bullets and text in a
  symbol font the browser cannot paint show their Unicode equivalents
  (`ooxml-core/text` `symbolCharToUnicode`, applied by the pptx paragraph
  builder), painted with a bundled symbol font when the paragraph font lacks
  them (see [font loading](font-loading.md)); devices with the font keep the
  exact glyph.
- **Clipboard.** A clipboard write must start while the click is handled, so
  Copy Slide as Image passes the render to `ClipboardItem` as a promise
  (`copyPngToClipboard`).
- **Downloads.** iPadOS reads a download's object URL only after the user
  confirms the download sheet, so `downloadBlob` keeps the URL for a minute.
- **Canvas limits.** iOS WebKit caps a canvas at 4096 x 4096 pixels of area and
  limits total canvas memory, so raster exports are tiled or scaled under that
  area (`maxCanvasAreaFor`) and each canvas is freed once read (`releaseCanvas`).
- **foreignObject exports.** WebKit decodes pictures and fonts inside an SVG
  image lazily, so the exporter awaits `img.decode()` and draws twice on WebKit.
- **Video.** Safari records MP4 only, so the container is chosen with
  `MediaRecorder.isTypeSupported` (`pickVideoRecordingFormat`).
- **Printing.** iOS prints the parent page when an iframe is printed, so
  `printHtmlDocument` prints through a window there and waits for it to load.
- **Clipboard formats.** Safari accepts no custom `web ` clipboard formats, so
  copied slides are also kept in the page for paste.
- **SVG filters.** WebKit resolves CSS `filter: url()` only within the document,
  so the Pixelate animation's filters live in a hidden `<svg>` in the page.
- **`user-select`.** Safari still needs `-webkit-user-select`; both are set.
- **Touch.** iOS sends no `contextmenu` on long-press and no reliable
  `dblclick`; hover-only controls never show. See the shared touch gestures in
  `src/ui/src/pptx/render/`.
