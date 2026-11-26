# pptx-vanilla-viewer

Open, edit, present and save PowerPoint (`.pptx`) files in any web page with one function call. No framework, no server, no conversion step, no PowerPoint install.

[![npm version](https://img.shields.io/npm/v/pptx-vanilla-viewer.svg)](https://www.npmjs.com/package/pptx-vanilla-viewer)
[![license](https://img.shields.io/npm/l/pptx-vanilla-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo-vanilla/) · [Docs](https://christophervr.github.io/ooxml/pptx/vanilla/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/vanilla/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![Editing, undoing and rendering a deck with the zero-framework demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/vanilla-demo.webp)

Slides are drawn as HTML, CSS and SVG, not as a canvas bitmap, so text stays sharp at any zoom, can be selected and is read by screen readers. The [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) engine that parses and writes the file is bundled in.

## Install

```bash
npm install pptx-vanilla-viewer
```

There are no required peer dependencies. Optional extras:

| For                                  | Add                                                            |
| ------------------------------------ | -------------------------------------------------------------- |
| Real-time collaboration              | `yjs` (`y-websocket` and `y-webrtc` are optional dependencies) |
| 3D models, 3D SmartArt and 3D charts | `three`                                                        |
| The AI assistant panel               | `ai`                                                           |

Without them, those features stay off and everything else works.

## Quick start

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const viewer = createPptxViewer(document.getElementById('host')!, {
	source: '/decks/quarterly.pptx', // a URL, ArrayBuffer, Uint8Array, Blob or File
	editable: true,
	onLoad: ({ slideCount }) => console.log(`${slideCount} slides`),
	onDirtyChange: (dirty) => console.log('unsaved edits:', dirty),
});

const bytes = await viewer.save(); // a valid .pptx
await viewer.downloadPptx('quarterly-edited.pptx');

viewer.destroy(); // removes the DOM, listeners, blob URLs and the engine handler
```

The viewer fills its container, so give the container a size. Styles are injected once per document; a page with a strict Content Security Policy imports `pptx-vanilla-viewer/styles.css` instead (`getViewerCss()` returns the same text). For a copy-paste Vite project, see [getting started](https://christophervr.github.io/ooxml/pptx/vanilla/getting-started).

## What you get

- **Viewing**: text, all 187 preset shapes, pictures, tables, charts, SmartArt, connectors, audio, video, ink, OLE objects and 3D models, with themes, masters and layouts resolved.
- **Editing**: a PowerPoint-style ribbon, inspector and on-canvas handles; inline text, tables, chart data, SmartArt, masters and layouts; full undo and redo.
- **Presenting**: transitions (including Morph), animations and motion paths, presenter view with notes and timer, rehearse timings, pen, highlighter and laser.
- **Exporting**: PDF, PNG, GIF, video and JSON, slides copied as images, printed handouts and notes pages, and Save As `.pptx`, `.ppsx`, `.pptm` or legacy `.ppt`.
- **Working together**: real-time co-editing and Broadcast over Yjs, through your server or peer to peer, with presence, cursors and threaded comments.
- **Assistant**: an optional AI panel whose edits follow a write policy you choose (apply, stage or approve).
- **Keyboard and screen readers**: F6 between regions, Tab through slide objects, Enter to edit text, and names that match the visible captions.
- **Seven UI languages**: English, French, Spanish, German, Simplified Chinese, Japanese and Korean.

Some effects are approximated or not yet supported; the [limitations page](https://christophervr.github.io/ooxml/pptx/guide/limitations) lists them, and a deck that uses one reports it as a load warning.

## Using it without a framework

**The instance.** `createPptxViewer` returns a handle with the shared slide and element API (`getContent`, `getSlides`, `addSlide`, `addElement`, `updateElement`, undo and redo, zoom, mode and selection) plus loading (`loadFile`, `loadUrl`), navigation (`next`, `prev`, `goToSlide`, `zoomToFit`), presenting (`enterPresentation`), saving and export (`save`, `downloadAs`, `exportPdf`, `exportSlidePng`, `print`), collaboration, autosave, `setTheme`, `setLocale` and `getHandler()` for the live engine. Callbacks are options (`onLoad`, `onError`, `onSlideChange`, `onChange`, `onSelectionChange` and more). See [the API](https://christophervr.github.io/ooxml/pptx/vanilla/api) and [options](https://christophervr.github.io/ooxml/pptx/vanilla/options).

**Custom renderers.** Elements are drawn through an open registry, so you can replace one kind without forking: `viewer.getRegistry().register('table', renderTable)`. See [renderers](https://christophervr.github.io/ooxml/pptx/vanilla/renderers).

**Your own shell.** `createViewerBuildingBlocks` runs the full viewer engine behind a canvas, a compact toolbar and the slide rail that you mount where you like (`mountCanvas`, `mountToolbar`, `mountSlidesPane`), and lets you answer picker and dialog requests with your own UI. `createExternalContentBar`, `createSlidesPane`, `renderSlideThumbnail` and `createRibbon` are exported for deeper custom chrome. See [building blocks](https://christophervr.github.io/ooxml/pptx/vanilla/building-blocks).

**Translations.** English is built in. Pass `messages` (one `pptx.*` dictionary per locale) with `locale`, or call `setLocale`; missing keys fall back to English. Dictionaries for the other languages are at `pptx-vanilla-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`. See [localization](https://christophervr.github.io/ooxml/pptx/guide/localization).

**AI assistant.** Create one assistant with `createAiAssistant` from `pptx-vanilla-viewer/ai` and pass it as `ai`. Its panel and the AI SDK load only when the user opens it. See the [assistant guide](https://christophervr.github.io/ooxml/pptx/guide/ai-assistant).

## Configuration

The most used options; the [options reference](https://christophervr.github.io/ooxml/pptx/vanilla/options) has all of them, and every binding takes the same set.

| Option                                           | What it does                                                                                                     |
| ------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------- |
| `editable`, `readOnly`                           | Turn editing on, or force the deck read-only.                                                                    |
| `initialSlide`, `showToolbar`, `showThumbnails`  | Starting slide and which chrome to show.                                                                         |
| `locale`, `messages`                             | The UI language and your dictionaries.                                                                           |
| `theme`, `defaultThemeKey`, `availableThemes`    | Colours and radius as `--pptx-*` tokens, and the File > Options appearance choices.                              |
| `customization`, `hiddenActions`, `ribbonAddIns` | Hide, lock or remap ribbon tabs, commands, options, menus, shortcuts and panels; add your own tabs.              |
| `fonts`, `googleFonts`                           | Supply licensed fonts, and control or redirect web font downloads.                                               |
| `allowExternalContent`                           | Download a deck's linked pictures and media without asking (`true`) or block them silently (`false`).            |
| `limits`                                         | Bound the work a load may do, for decks you do not trust.                                                        |
| `collaboration`, `shareDefaults`                 | Join or seed a real-time session.                                                                                |
| `storage`, `fileProviders`, `autosave`           | Save to your own storage or OneDrive, and control recovery autosave.                                             |
| `registry`                                       | Your own element renderers.                                                                                      |
| `ai`                                             | The optional assistant.                                                                                          |
| `fitPadding`, `maxFitScale`                      | How the slide fits its viewport ([viewport fit](https://christophervr.github.io/ooxml/pptx/guide/viewport-fit)). |

OneDrive needs a Microsoft sign-in your host provides; see the [OneDrive guide](https://christophervr.github.io/ooxml/pptx/guide/onedrive). It stores files and does not provide Office co-authoring.

## Supported environments

- Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- Any bundler that resolves ES modules, or none: the package has no framework dependency.
- Node.js 22.3 or newer to install, build and test.

## Privacy and network use

The viewer renders and saves decks in the browser and never uploads them. It contacts other servers only in these cases:

- **Fonts that are neither installed nor embedded.** When a deck uses such a font and Google Fonts serves it (or a metric-compatible clone, such as Carlito for Calibri), the viewer downloads it from `fonts.googleapis.com` and `fonts.gstatic.com`, so the reader's IP address and the needed family names reach Google. This is on by default, as PowerPoint downloads its cloud fonts. Turn it off with `googleFonts: false`, or send the same requests to your own mirror with `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }`. Readers can turn it off in File > Options > Trust Center ("Download fonts that are not installed or embedded (Google Fonts)"); a host value locks that toggle. With downloads off, text uses installed, embedded or substitute fonts, so line breaks can differ from PowerPoint. A slow download holds the deck for up to 8 seconds, and the deck opens at once when the font server cannot be reached, so offline use works. See [font loading](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md).
- **Linked pictures and media** are not fetched when a deck opens, as in PowerPoint. A security bar ("References to external pictures and other media have been blocked") offers **Enable Content**, which draws them without reloading: edits and undo history stay, and in a shared session each participant chooses for themselves. The choice lasts until another presentation opens. The Trust Center option "Allow external content" (off by default) fetches them without asking; `allowExternalContent: true` does the same and `false` blocks them without the bar, each locking that option.
- **Hyperlinks** open only when the reader clicks them, always in a new tab with `noopener` and `noreferrer`, so the linked site receives no referrer and cannot script the viewer's page, whatever frame the deck names. `file:` and `ftp:` links stay in the deck but never open. Relative links (no scheme, such as `notes.html`) open, against the host page's address, only after the host calls `setOpenRelativeHyperlinks(true)` from `ooxml-ui/pptx`.
- **Share and Broadcast**, only when started. Updates go to the server you configure. With no server, peers meet through y-webrtc's public signaling server and public STUN servers; the signaling server sees a hash of the secret in the invitation link's `#key=` fragment (never the room name), the participants' IP addresses and encrypted signaling messages.

Nothing collects analytics or telemetry. The full list, including what is kept in local storage, is in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

Draw stores pen preset types, colours and thicknesses in the local viewer preferences. A host can redirect that storage with the pen preference adapter in `configurePptxDraw`. Ink to Math calls a host recognizer only when configured and explicitly requested. It passes the selected strokes as slide-local coordinates. See [PowerPoint privacy](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#ink-to-math).

## Related packages

| Package                                                                    | Use it for                                                     |
| -------------------------------------------------------------------------- | -------------------------------------------------------------- |
| [`pptx-react-viewer`](https://www.npmjs.com/package/pptx-react-viewer)     | The same viewer for React.                                     |
| [`pptx-vue-viewer`](https://www.npmjs.com/package/pptx-vue-viewer)         | The same viewer for Vue 3.                                     |
| [`pptx-angular-viewer`](https://www.npmjs.com/package/pptx-angular-viewer) | The same viewer for Angular.                                   |
| [`pptx-svelte-viewer`](https://www.npmjs.com/package/pptx-svelte-viewer)   | The same viewer for Svelte 5.                                  |
| [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core)       | Reading, building and saving decks in Node or a worker, no UI. |
| [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp)         | PowerPoint editing tools for AI agents, as an MCP server.      |

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), which also has [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) and [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) viewers.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
