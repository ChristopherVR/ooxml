# pptx-react-viewer

**A React component that opens, edits, presents and saves PowerPoint files (`.pptx`, `.pptm`, `.ppsx`, `.potx` and legacy `.ppt`) in the browser, with no server.**

[![npm version](https://img.shields.io/npm/v/pptx-react-viewer.svg)](https://www.npmjs.com/package/pptx-react-viewer)
[![license](https://img.shields.io/npm/l/pptx-react-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)
[![types](https://img.shields.io/npm/types/pptx-react-viewer.svg)](https://www.npmjs.com/package/pptx-react-viewer)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo/) · [Documentation](https://christophervr.github.io/ooxml/pptx/) · [React guide](https://christophervr.github.io/ooxml/pptx/react/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/react/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![Selecting, dragging and resizing a slide element in the React demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/react-demo.webp)

## Install

```bash
npm install pptx-react-viewer
npm install react react-dom framer-motion lucide-react react-icons jspdf jszip fast-xml-parser i18next react-i18next
```

The second line is the required peers: `react` and `react-dom` `^18.2.0 || ^19.0.0`, `framer-motion` 12, 13 or 14, `i18next` 26 or 27, `react-i18next` 17 or 18, `lucide-react` 1, `react-icons` 5, `jspdf` 4, `jszip` 3 and `fast-xml-parser` 5. The engine comes with it (`ooxml-core` and `ooxml-ui` are dependencies). Optional extras:

| For                                  | Add                                                                      |
| ------------------------------------ | ------------------------------------------------------------------------ |
| Real-time collaboration              | `yjs` (`y-websocket` and `y-webrtc` are optional dependencies)           |
| 3D models, 3D SmartArt and 3D charts | nothing: `three` is an optional dependency that npm installs when it can |
| The AI assistant panel               | `ai` (6 or 7) and `@ai-sdk/react` 4                                      |

Without them those features stay off and everything else works.

## Quick start

```tsx
import { PowerPointViewer, type PowerPointViewerHandle } from 'pptx-react-viewer';
import 'pptx-react-viewer/styles';
import { useRef, useState, type ChangeEvent } from 'react';

export default function App() {
	const [content, setContent] = useState<Uint8Array | null>(null);
	const viewer = useRef<PowerPointViewerHandle>(null);

	const onPick = async (event: ChangeEvent<HTMLInputElement>) => {
		const file = event.target.files?.[0];
		if (file) setContent(new Uint8Array(await file.arrayBuffer()));
	};

	const save = async () => {
		const bytes = await viewer.current?.getContent(); // the edited .pptx
	};

	return (
		<div style={{ height: '100vh' }}>
			{content ? (
				<PowerPointViewer ref={viewer} content={content} canEdit />
			) : (
				<input type='file' accept='.pptx,.ppt' onChange={onPick} />
			)}
		</div>
	);
}
```

`content` is a `Uint8Array`. The viewer fills its parent, so give the parent a height. `pptx-react-viewer/styles` is the full stylesheet, scoped to the viewer; an app that builds Tailwind CSS v4 itself can use `pptx-react-viewer/theme.css` instead, and a host that renders only `SlideCanvas` can import `pptx-react-viewer/canvas.css`.

## Props

The most used props. Every binding takes the same set; the [props reference](https://christophervr.github.io/ooxml/pptx/react/props) has all of them.

| Prop                                             | What it does                                                                                    |
| ------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| `content`                                        | The deck's bytes (`Uint8Array`).                                                                |
| `canEdit`, `readOnly`                            | Turn editing on, or force the deck read-only (`readOnly` wins).                                 |
| `initialSlide`, `showToolbar`, `showThumbnails`  | Starting slide and which chrome to show.                                                        |
| `theme`, `defaultThemeKey`, `availableThemes`    | Viewer chrome colours and radius, and the File > Options appearance choices.                    |
| `defaultLocale`, `availableLocales`              | The starting UI language and the languages offered.                                             |
| `customization`, `hiddenActions`, `ribbonAddIns` | Hide, lock or remap ribbon tabs, commands, options, menus, shortcuts and panels; add your tabs. |
| `fonts`, `googleFonts`                           | Supply licensed fonts; turn off or redirect web font downloads.                                 |
| `allowExternalContent`                           | Fetch a deck's linked pictures and media without asking (`true`) or block them (`false`).       |
| `limits`                                         | Bound the work a load may do, for decks you do not trust.                                       |
| `collaboration`, `shareDefaults`                 | Join or seed a real-time session.                                                               |
| `storage`, `fileProviders`, `autosave`           | Save to your own storage or a cloud provider, and control recovery autosave.                    |
| `ai`                                             | The optional assistant, from `createAiAssistant` in `pptx-react-viewer/ai`.                     |
| `fitPadding`, `maxFitScale`                      | How the slide fits its viewport.                                                                |

The `ref` handle (`PowerPointViewerHandle`) has `getContent`, navigation (`goTo`, `goNext`, `goPrev`), zoom (`setZoom`, `zoomIn`, `zoomOut`), `setMode`, `undo`, `redo`, `isDirty`, and the slide and element API (`getSlides`, `addSlide`, `deleteSlides`, `addElement`, `updateElement`, `deleteElements`, `selectElements`). See [the handle](https://christophervr.github.io/ooxml/pptx/react/handle).

## Events

Callback props, most used first: `onContentChange` (the edited bytes), `onDirtyChange`, `onActiveSlideChange`, `onSelectionChange`, `onLoad`, `onError`, `onChange`, `onModeChange`, `onZoomChange`, `onSlideCountChange`, `onAutosave`, `onCollaborationChange`, `onHyperlinkClick`, `onOpenFile`, `onThemeChange`, `onLocaleChange` and `onPdfExportReport`. See the [props reference](https://christophervr.github.io/ooxml/pptx/react/props).

## Your own UI

`useViewerBuildingBlocks` runs the full viewer engine behind `SlideCanvas`, `Toolbar` and `SlidesPane`, so you can place them in your own layout and answer picker and dialog requests with your own UI. See [hooks](https://christophervr.github.io/ooxml/pptx/react/hooks).

## Theming

The chrome inherits the shared `--office-*` tokens: put `data-office-theme="light"`, `"dark"` or `"auto"` on a containing element, or set tokens such as `--office-accent` there. The `theme` prop (a partial `ViewerTheme`) overrides them through the viewer's `--pptx-*` properties; `vermilionLightTheme` and `vermilionDarkTheme` are ready-made, and `ViewerThemeProvider` themes a subtree. See [theming](https://christophervr.github.io/ooxml/pptx/react/theming).

## Localisation

English, French, Spanish, German, Simplified Chinese, Japanese and Korean. Labels go through your `react-i18next` instance (without one the viewer shows its built-in English): register `translationsEn` from `pptx-react-viewer/i18n` and the other dictionaries from `pptx-react-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`. For another library, wrap the viewer in `PptxI18nProvider` with your own `translate(key, params)`, and import from `pptx-react-viewer/no-i18next` to leave i18next out of the bundle. See [localisation](https://christophervr.github.io/ooxml/pptx/guide/localization).

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
- **Linked pictures and media (off by default).** Fetched only after the reader chooses Enable Content on the security bar, or always with `allowExternalContent: true`; `false` blocks them without the bar.
- **Share and Broadcast (only when started).** Updates go to the server you configure; without one, peers meet through y-webrtc's public signalling server and public STUN servers. Pass your own `collaboration.signaling` list to avoid them.
- **AI assistant (off unless you pass `ai`).** Requests go to the connection you configure, with the conversation and deck context (`contextStrategy: 'none'` sends no deck context).
- **Links** (hyperlinks, Microsoft Forms responses) open only when clicked, in a new tab without opener or referrer.

Autosave, recovery and preferences stay on the device (IndexedDB and localStorage). The full list, including Ink to Math and the camera, is in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#powerpoint-ooxml-uipptx-and-the-pptx--viewer-packages).

## Browser support

- The last two major versions of Chrome, Edge, Firefox and Safari; Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- React 18.2 or 19.
- Node.js 22.3 or newer to install and build.

## Documentation and demos

- [Documentation](https://christophervr.github.io/ooxml/pptx/), the [React guide](https://christophervr.github.io/ooxml/pptx/react/) and the [live demo](https://christophervr.github.io/ooxml/pptx/demo/).
- The same viewer for [Vue](https://www.npmjs.com/package/pptx-vue-viewer), [Angular](https://www.npmjs.com/package/pptx-angular-viewer), [Svelte](https://www.npmjs.com/package/pptx-svelte-viewer) and [no framework](https://www.npmjs.com/package/pptx-vanilla-viewer).
- [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) for decks without a UI, and [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp) for AI agents.
- Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), with Word, Excel and Visio editors.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
