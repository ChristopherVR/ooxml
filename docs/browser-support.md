# Browser support

The editors and viewers target current evergreen browsers. The floor is set by
what the shipped code and styles already require, not by a policy choice:

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
  builder); devices with the font keep the exact glyph.
- **Clipboard.** A clipboard write must start while the click is handled, so
  Copy Slide as Image passes the render to `ClipboardItem` as a promise
  (`copyPngToClipboard`).
- **Downloads.** iPadOS reads a download's object URL only after the user
  confirms the download sheet, so `downloadBlob` keeps the URL for a minute.
- **SVG filters.** WebKit resolves CSS `filter: url()` only within the document,
  so the Pixelate animation's filters live in a hidden `<svg>` in the page.
- **`user-select`.** Safari still needs `-webkit-user-select`; both are set.
- **Touch.** iOS sends no `contextmenu` on long-press and no reliable
  `dblclick`; hover-only controls never show. See the shared touch gestures in
  `src/ui/src/pptx/render/`.
