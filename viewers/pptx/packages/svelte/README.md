# pptx-svelte-viewer

A Svelte 5 component that opens, edits, presents and saves PowerPoint (`.pptx`) files in the browser. No server, no conversion step, no PowerPoint install.

[![npm version](https://img.shields.io/npm/v/pptx-svelte-viewer.svg)](https://www.npmjs.com/package/pptx-svelte-viewer)
[![license](https://img.shields.io/npm/l/pptx-svelte-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo-svelte/) · [Docs](https://christophervr.github.io/ooxml/pptx/svelte/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/svelte/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![A PowerPoint deck rendered by the Svelte 5 viewer demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/svelte-demo.webp)

Slides are drawn as HTML, CSS and SVG, not as a canvas bitmap, so text stays sharp at any zoom, can be selected and is read by screen readers. The component is written with runes, and the [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) engine that parses and writes the file is bundled in.

## Install

```bash
npm install pptx-svelte-viewer
```

The only peer dependency is Svelte 5.7 or newer. Optional extras:

| For                                  | Add                                                                 |
| ------------------------------------ | ------------------------------------------------------------------- |
| Real-time collaboration              | `yjs` (`y-websocket` and `y-webrtc` are optional dependencies)      |
| 3D models, 3D SmartArt and 3D charts | nothing: `three` is an optional dependency npm installs when it can |
| The AI assistant panel               | `ai` and `@ai-sdk/svelte`                                           |

Without them, those features stay off and everything else works.

## Quick start

```svelte
<script lang="ts">
	import { PowerPointViewer, type PowerPointViewerApi } from 'pptx-svelte-viewer';
	import 'pptx-svelte-viewer/styles';

	let source: Uint8Array | undefined = $state();
	let viewer: PowerPointViewerApi | undefined = $state();

	async function onpick(event: Event) {
		const file = (event.currentTarget as HTMLInputElement).files?.[0];
		if (file) source = new Uint8Array(await file.arrayBuffer());
	}

	async function save() {
		const bytes = await viewer?.getContent(); // a valid .pptx
	}
</script>

<div style="height: 100vh">
	{#if source}
		<PowerPointViewer bind:this={viewer} {source} editable />
	{:else}
		<input type="file" accept=".pptx,.ppt" onchange={onpick} />
	{/if}
</div>
```

`source` takes a `Uint8Array` or an `ArrayBuffer`. The viewer fills its parent, so give the parent a height. The stylesheet ships as a real CSS file rather than injected styles, so it applies under SSR and a strict Content Security Policy; import it once at your app entry.

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

## Using it in Svelte

**Props, callbacks and `bind:this`.** The editing switch is `editable` and the bytes are `source`. Changes arrive as callback props (`oncontentchange`, `ondirtychange`, `onslidechange`, `onselectionchange`, `onload`, `onerror` and more). The bound instance (`PowerPointViewerApi`) covers navigation, zoom, mode, selection, undo and redo, the shared slide and element API (`getSlides`, `addSlide`, `addElement`, `updateElement` and the rest), and saving and export (`save`, `downloadPptx`, `exportPdf`, `exportSlidePng`, `print`). See [props](https://christophervr.github.io/ooxml/pptx/svelte/props) and [the API](https://christophervr.github.io/ooxml/pptx/svelte/api).

**Theming.** Pass a partial `theme` prop; `vermilionLightTheme` and `vermilionDarkTheme` are ready-made presets. See [theming](https://christophervr.github.io/ooxml/pptx/svelte/theming).

**Your own shell.** `createViewerBuildingBlocks` builds the full viewer state behind `BuildingBlocksCanvas`, `BuildingBlocksToolbar` and `BuildingBlocksSlidesPane`, so you can place them in your own layout and answer picker and dialog requests with your own UI. Call it synchronously in the script of the component that renders the blocks. `pptx-svelte-viewer/viewer` has the lower-level `createViewerState`, `SlideCanvas`, `EditorLayer` and presence layers. See [building blocks](https://christophervr.github.io/ooxml/pptx/svelte/building-blocks).

**Translations.** English is built in. Register another language with `registerTranslations(locale, dictionary)` from `pptx-svelte-viewer/i18n`, using the dictionaries at `pptx-svelte-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`, then set the `locale` prop. See [i18n](https://christophervr.github.io/ooxml/pptx/svelte/i18n).

**AI assistant.** Create one assistant with `createAiAssistant` from `pptx-svelte-viewer/ai` and pass it to the `ai` prop. Its panel and the AI SDK load only when the user opens it. See the [assistant guide](https://christophervr.github.io/ooxml/pptx/guide/ai-assistant).

## Configuration

The most used props; the [props reference](https://christophervr.github.io/ooxml/pptx/svelte/props) has all of them, and every binding takes the same set.

| Prop                                             | What it does                                                                                                     |
| ------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------- |
| `editable`, `readOnly`                           | Turn editing on, or force the deck read-only.                                                                    |
| `initialSlide`, `showToolbar`, `showThumbnails`  | Starting slide and which chrome to show.                                                                         |
| `locale`                                         | The UI language.                                                                                                 |
| `theme`, `defaultThemeKey`, `availableThemes`    | Colours and radius as `--pptx-*` tokens, and the File > Options appearance choices.                              |
| `customization`, `hiddenActions`, `ribbonAddIns` | Hide, lock or remap ribbon tabs, commands, options, menus, shortcuts and panels; add your own tabs.              |
| `fonts`, `googleFonts`                           | Supply licensed fonts, and control or redirect web font downloads.                                               |
| `allowExternalContent`                           | Download a deck's linked pictures and media without asking (`true`) or block them silently (`false`).            |
| `limits`                                         | Bound the work a load may do, for decks you do not trust.                                                        |
| `collaboration`, `shareDefaults`                 | Join or seed a real-time session.                                                                                |
| `storage`, `fileProviders`, `autosave`           | Save to your own storage or OneDrive, and control recovery autosave.                                             |
| `ai`                                             | The optional assistant.                                                                                          |
| `fitPadding`, `maxFitScale`                      | How the slide fits its viewport ([viewport fit](https://christophervr.github.io/ooxml/pptx/guide/viewport-fit)). |

OneDrive needs a Microsoft sign-in your host provides; see the [OneDrive guide](https://christophervr.github.io/ooxml/pptx/guide/onedrive). It stores files and does not provide Office co-authoring.

## Supported environments

- Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- Svelte 5.7 or newer.
- Node.js 22.3 or newer to install, build and test.

## Privacy and network use

The viewer renders and saves decks in the browser and never uploads them. It contacts other servers only in these cases:

- **Fonts that are neither installed nor embedded.** When a deck uses such a font and Google Fonts serves it (or a metric-compatible clone, such as Carlito for Calibri), the viewer downloads it from `fonts.googleapis.com` and `fonts.gstatic.com`, so the reader's IP address and the needed family names reach Google. This is on by default, as PowerPoint downloads its cloud fonts. Turn it off with `googleFonts: false`, or send the same requests to your own mirror with `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }`. Readers can turn it off in File > Options > Trust Center ("Download fonts that are not installed or embedded (Google Fonts)"); a host value locks that toggle. With downloads off, text uses installed, embedded or substitute fonts, so line breaks can differ from PowerPoint. A slow download holds the deck for up to 8 seconds, and the deck opens at once when the font server cannot be reached, so offline use works. See [font loading](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md).
- **Linked pictures and media** are not fetched when a deck opens, as in PowerPoint. A security bar offers **Enable Content** for that presentation, which lasts until another presentation opens. The deck keeps the links and saves write them back; Enable Content only redraws, so unsaved edits, undo history, autosave and a collaboration session carry on. The Trust Center option "Allow external content" (off by default) fetches them without asking. `allowExternalContent: true` fetches them without the bar and `false` blocks them without it; either locks the option.
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
| [`pptx-vanilla-viewer`](https://www.npmjs.com/package/pptx-vanilla-viewer) | The same viewer with no framework.                             |
| [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core)       | Reading, building and saving decks in Node or a worker, no UI. |
| [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp)         | PowerPoint editing tools for AI agents, as an MCP server.      |

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), which also has [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) and [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) viewers.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
