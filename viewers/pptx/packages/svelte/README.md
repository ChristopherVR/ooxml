# pptx-svelte-viewer

**A Svelte 5 component that opens, edits, presents and saves PowerPoint files (`.pptx`, `.pptm`, `.ppsx`, `.potx` and legacy `.ppt`) in the browser, with no server.**

[![npm version](https://img.shields.io/npm/v/pptx-svelte-viewer.svg)](https://www.npmjs.com/package/pptx-svelte-viewer)
[![license](https://img.shields.io/npm/l/pptx-svelte-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)
[![types](https://img.shields.io/npm/types/pptx-svelte-viewer.svg)](https://www.npmjs.com/package/pptx-svelte-viewer)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo-svelte/) · [Documentation](https://christophervr.github.io/ooxml/pptx/) · [Svelte guide](https://christophervr.github.io/ooxml/pptx/svelte/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/svelte/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![A PowerPoint deck rendered by the Svelte 5 viewer demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/svelte-demo.webp)

## Install

```bash
npm install pptx-svelte-viewer
```

The only required peer is `svelte` `^5.7.0`. The engine comes with it (`ooxml-core` and `ooxml-ui` are dependencies), and the component is written with runes. Optional extras:

| For                                  | Add                                                                      |
| ------------------------------------ | ------------------------------------------------------------------------ |
| Real-time collaboration              | `yjs` (`y-websocket` and `y-webrtc` are optional dependencies)           |
| 3D models, 3D SmartArt and 3D charts | nothing: `three` is an optional dependency that npm installs when it can |
| The AI assistant panel               | `ai` (6 or 7) and `@ai-sdk/svelte` 5                                     |

Without them those features stay off and everything else works.

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
		const bytes = await viewer?.getContent(); // the edited .pptx
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

`source` takes a `Uint8Array` or an `ArrayBuffer`. The viewer fills its parent, so give the parent a height. The stylesheet is a real CSS file rather than injected styles, so it works under SSR and a strict Content Security Policy; import it once at your app entry.

## Props

The most used props. Every binding takes the same set (here the bytes are `source` and the edit switch is `editable`); the [props reference](https://christophervr.github.io/ooxml/pptx/svelte/props) has all of them.

| Prop                                             | What it does                                                                                    |
| ------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| `source`                                         | The deck's bytes (`Uint8Array` or `ArrayBuffer`).                                               |
| `editable`, `readOnly`                           | Turn editing on, or force the deck read-only (`readOnly` wins).                                 |
| `initialSlide`, `showToolbar`, `showThumbnails`  | Starting slide and which chrome to show.                                                        |
| `locale`, `defaultLocale`, `availableLocales`    | The UI language and the languages offered.                                                      |
| `theme`, `defaultThemeKey`, `availableThemes`    | Viewer chrome colours and radius, and the File > Options appearance choices.                    |
| `customization`, `hiddenActions`, `ribbonAddIns` | Hide, lock or remap ribbon tabs, commands, options, menus, shortcuts and panels; add your tabs. |
| `fonts`, `googleFonts`                           | Supply licensed fonts; turn off or redirect web font downloads.                                 |
| `allowExternalContent`                           | Fetch a deck's linked pictures and media without asking (`true`) or block them (`false`).       |
| `limits`                                         | Bound the work a load may do, for decks you do not trust.                                       |
| `collaboration`, `shareDefaults`                 | Join or seed a real-time session.                                                               |
| `storage`, `fileProviders`, `autosave`           | Save to your own storage or a cloud provider, and control recovery autosave.                    |
| `ai`                                             | The optional assistant, from `createAiAssistant` in `pptx-svelte-viewer/ai`.                    |
| `fitPadding`, `maxFitScale`                      | How the slide fits its viewport.                                                                |

The bound instance (`PowerPointViewerApi`) has `getContent`, navigation (`goTo`, `goNext`, `goPrev`), zoom, `setMode`, `undo`, `redo`, the slide and element API (`getSlides`, `addSlide`, `addElement`, `updateElement`, `deleteElements`), and saving and export (`save`, `downloadPptx`, `exportPdf`, `exportSlidePng`). See [the API](https://christophervr.github.io/ooxml/pptx/svelte/api).

## Events

Callback props, most used first: `oncontentchange` (the edited bytes), `ondirtychange`, `onslidechange`, `onselectionchange`, `onload`, `onerror`, `onchange`, `onmodechange`, `onzoomchange`, `onslidecountchange`, `onautosave`, `oncollaborationchange`, `onhyperlinkclick`, `onopenfile` and `onpdfexportreport`, plus `onThemeChange`, `onLocaleChange` and `onElementRenderError`. See the [props reference](https://christophervr.github.io/ooxml/pptx/svelte/props).

## Your own UI

`createViewerBuildingBlocks` builds the full viewer state behind `BuildingBlocksCanvas`, `BuildingBlocksToolbar` and `BuildingBlocksSlidesPane`, so you can place them in your own layout and answer picker and dialog requests with your own UI. Call it synchronously in the script of the component that renders the blocks; `pptx-svelte-viewer/viewer` has the lower-level pieces. See [building blocks](https://christophervr.github.io/ooxml/pptx/svelte/building-blocks).

## Theming

The chrome inherits the shared `--office-*` tokens: put `data-office-theme="light"`, `"dark"` or `"auto"` on a containing element, or set tokens such as `--office-accent` there. The `theme` prop (a partial `ViewerTheme`) overrides them through the viewer's `--pptx-*` properties; `vermilionLightTheme` and `vermilionDarkTheme` are ready-made. See [theming](https://christophervr.github.io/ooxml/pptx/svelte/theming).

## Localisation

English (built in), French, Spanish, German, Simplified Chinese, Japanese and Korean. Register a language with `registerTranslations(locale, dictionary)` from `pptx-svelte-viewer/i18n`, using the dictionaries at `pptx-svelte-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`, then set the `locale` prop. See [i18n](https://christophervr.github.io/ooxml/pptx/svelte/i18n).

## Status and limits

PowerPoint carries no stability label yet: it has not been assessed against the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md).

- Saving keeps the parts you did not touch, but saving an edited slide can still rewrite minor markup; it is not a lossless round trip.
- Some effects are approximated (3D materials, some WordArt warps, SmartArt without its cached drawing). A deck that uses an approximated construct reports it as a load warning.
- The editor covers a subset of PowerPoint: Remove Background, narration, Notes Master, Handout Master, Notes Page view, Grayscale views and Designer are not there yet.
- OneDrive stores files through an adapter whose sign-in your host provides; there is no Office co-authoring.

See the [limitations page](https://christophervr.github.io/ooxml/pptx/guide/limitations) and the [Office parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md#powerpoint).

## Privacy and network use

The viewer renders and saves decks in the browser and never uploads them. Nothing collects analytics or telemetry. It contacts other servers only in these cases:

- **Web fonts (on by default).** A font that is neither installed nor embedded, when the bundled Google Fonts catalogue serves it or a stand-in, is downloaded from `fonts.googleapis.com` and `fonts.gstatic.com`; Google sees the reader's IP address and the family names. Set `googleFonts` to `false` to turn it off, or to `{ baseUrl }` to use your own mirror; readers can turn it off in File > Options > Trust Center. Offline, the deck opens with substitute fonts.
- **Linked pictures and media (off by default).** Fetched only after the reader chooses Enable Content on the security bar, or always with `allowExternalContent` set to `true`; `false` blocks them without the bar.
- **Share and Broadcast (only when started).** Updates go to the server you configure; without one, peers meet through y-webrtc's public signalling server and public STUN servers. Pass your own `collaboration.signaling` list to avoid them.
- **AI assistant (off unless you pass `ai`).** Requests go to the connection you configure, with the conversation and deck context (`contextStrategy: 'none'` sends no deck context).
- **Links** (hyperlinks, Microsoft Forms responses) open only when clicked, in a new tab without opener or referrer.

Autosave, recovery and preferences stay on the device (IndexedDB and localStorage). The full list, including Ink to Math and the camera, is in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#powerpoint-ooxml-uipptx-and-the-pptx--viewer-packages).

## Browser support

- The last two major versions of Chrome, Edge, Firefox and Safari; Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- Svelte 5.7 or newer.
- Node.js 22.3 or newer to install and build.

## Documentation and demos

- [Documentation](https://christophervr.github.io/ooxml/pptx/), the [Svelte guide](https://christophervr.github.io/ooxml/pptx/svelte/) and the [live demo](https://christophervr.github.io/ooxml/pptx/demo-svelte/).
- The same viewer for [React](https://www.npmjs.com/package/pptx-react-viewer), [Vue](https://www.npmjs.com/package/pptx-vue-viewer), [Angular](https://www.npmjs.com/package/pptx-angular-viewer) and [no framework](https://www.npmjs.com/package/pptx-vanilla-viewer).
- [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) for decks without a UI, and [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp) for AI agents.
- Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), with Word, Excel and Visio editors.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
