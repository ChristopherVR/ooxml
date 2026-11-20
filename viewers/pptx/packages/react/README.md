# pptx-react-viewer

A React component that opens, edits, presents and saves PowerPoint (`.pptx`) files in the browser. No server, no conversion step, no PowerPoint install.

[![npm version](https://img.shields.io/npm/v/pptx-react-viewer.svg)](https://www.npmjs.com/package/pptx-react-viewer)
[![license](https://img.shields.io/npm/l/pptx-react-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)
[![types](https://img.shields.io/npm/types/pptx-react-viewer.svg)](https://www.npmjs.com/package/pptx-react-viewer)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo/) · [Docs](https://christophervr.github.io/ooxml/pptx/react/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/react/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![Selecting, dragging and resizing a slide element in the React demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/react-demo.webp)

Slides are drawn as HTML, CSS and SVG, not as a canvas bitmap, so text stays sharp at any zoom, can be selected and is read by screen readers. The [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) engine that parses and writes the file is bundled in.

## Install

```bash
npm install pptx-react-viewer
npm install react react-dom framer-motion lucide-react react-icons jspdf jszip fast-xml-parser i18next react-i18next
```

The second line is the peer dependencies; `react` and `react-dom` may be 18.2 or 19. Optional extras:

| For                                  | Add                                                                 |
| ------------------------------------ | ------------------------------------------------------------------- |
| Real-time collaboration              | `yjs` (`y-websocket` and `y-webrtc` are optional dependencies)      |
| 3D models, 3D SmartArt and 3D charts | nothing: `three` is an optional dependency npm installs when it can |
| The AI assistant panel               | `ai` and `@ai-sdk/react`                                            |

Without them, those features stay off and everything else works.

## Quick start

```tsx
import { PowerPointViewer } from 'pptx-react-viewer';
import 'pptx-react-viewer/styles'; // skip if your app builds Tailwind CSS v4 itself
import { useState } from 'react';

export default function App() {
	const [content, setContent] = useState<Uint8Array | null>(null);

	const onPick = async (event: React.ChangeEvent<HTMLInputElement>) => {
		const file = event.target.files?.[0];
		if (file) setContent(new Uint8Array(await file.arrayBuffer()));
	};

	return (
		<div style={{ height: '100vh' }}>
			{content ? (
				<PowerPointViewer content={content} canEdit />
			) : (
				<input type='file' accept='.pptx,.ppt' onChange={onPick} />
			)}
		</div>
	);
}
```

`content` is a `Uint8Array` (wrap an `ArrayBuffer` with `new Uint8Array(buffer)`). The viewer fills its parent, so give the parent a height. To read the edited deck back, pass a `ref` and call `getContent()`:

```tsx
const viewer = useRef<PowerPointViewerHandle>(null);
// <PowerPointViewer ref={viewer} content={content} canEdit />
const bytes = await viewer.current?.getContent(); // a valid .pptx
```

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

## Using it in React

**Styles.** `pptx-react-viewer/styles` is the full stylesheet. Its base reset applies only inside the viewer, so your page keeps its own headings and margins. A Tailwind CSS v4 app can build the viewer's utilities itself from `pptx-react-viewer/theme.css`, and a host that renders only `SlideCanvas` can import the smaller `pptx-react-viewer/canvas.css`. See [theming](https://christophervr.github.io/ooxml/pptx/react/theming).

**The ref handle.** `PowerPointViewerHandle` covers navigation, zoom, mode, selection, undo and redo, plus the shared slide and element API (`getSlides`, `addSlide`, `addElement`, `updateElement`, `deleteElements` and the rest). See [the handle](https://christophervr.github.io/ooxml/pptx/react/handle).

**Your own shell.** `useViewerBuildingBlocks` runs the full viewer engine behind `SlideCanvas`, `Toolbar` and `SlidesPane`, so you can place them in your own layout and answer picker and dialog requests with your own UI. See [hooks](https://christophervr.github.io/ooxml/pptx/react/hooks#public-hooks).

**Translations.** Labels go through your `react-i18next` instance (`pptx-react-viewer/i18n` has the English dictionary and `keyToLabel` for a fallback). To use Lingui, FormatJS or another library, wrap the viewer in `PptxI18nProvider` with a `translate(key, params)` function, and import from `pptx-react-viewer/no-i18next` to leave i18next out of the bundle. Dictionaries for the other languages are at `pptx-react-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`. See [localization](https://christophervr.github.io/ooxml/pptx/guide/localization).

**AI assistant.** Create one assistant with `createAiAssistant` from `pptx-react-viewer/ai` and pass it to `ai`. Its panel and the AI SDK load only when the user opens it. See the [assistant guide](https://christophervr.github.io/ooxml/pptx/guide/ai-assistant).

```tsx
import { createAiAssistant } from 'pptx-react-viewer/ai';

const assistant = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/ai/chat' } });
// <PowerPointViewer content={content} ai={assistant} />
```

## Configuration

The most used props; the [props reference](https://christophervr.github.io/ooxml/pptx/react/props) has all of them, and every binding takes the same set.

| Prop                                                      | What it does                                                                                                     |
| --------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| `canEdit`, `readOnly`                                     | Turn editing on, or force the deck read-only.                                                                    |
| `initialSlide`, `showToolbar`, `showThumbnails`           | Starting slide and which chrome to show.                                                                         |
| `theme`, `defaultThemeKey`, `availableThemes`             | Colours and radius as `--pptx-*` tokens, and the File > Options appearance choices.                              |
| `customization`, `hiddenActions`, `ribbonAddIns`          | Hide, lock or remap ribbon tabs, commands, options, menus, shortcuts and panels; add your own tabs.              |
| `fonts`, `googleFonts`                                    | Supply licensed fonts, and control or redirect web font downloads.                                               |
| `allowExternalContent`                                    | Download a deck's linked pictures and media without asking (`true`) or block them silently (`false`).            |
| `limits`                                                  | Bound the work a load may do, for decks you do not trust.                                                        |
| `collaboration`, `shareDefaults`                          | Join or seed a real-time session.                                                                                |
| `storage`, `fileProviders`, `autosave`                    | Save to your own storage or OneDrive, and control recovery autosave.                                             |
| `ai`                                                      | The optional assistant.                                                                                          |
| `fitPadding`, `maxFitScale`                               | How the slide fits its viewport ([viewport fit](https://christophervr.github.io/ooxml/pptx/guide/viewport-fit)). |
| `onContentChange`, `onDirtyChange`, `onActiveSlideChange` | Follow edits and navigation.                                                                                     |

OneDrive needs a Microsoft sign-in your host provides; see the [OneDrive guide](https://christophervr.github.io/ooxml/pptx/guide/onedrive). It stores files and does not provide Office co-authoring.

## Supported environments

- Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- React 18.2 or 19.
- Node.js 22.3 or newer to install, build and test.

## Privacy and network use

The viewer renders and saves decks in the browser and never uploads them. It contacts other servers only in these cases:

- **Fonts that are neither installed nor embedded.** When a deck uses such a font and Google Fonts serves it (or a metric-compatible clone, such as Carlito for Calibri), the viewer downloads it from `fonts.googleapis.com` and `fonts.gstatic.com`, so the reader's IP address and the needed family names reach Google. This is on by default, as PowerPoint downloads its cloud fonts. Turn it off with `googleFonts: false`, or send the same requests to your own mirror with `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }`. Readers can turn it off in File > Options > Trust Center ("Download fonts that are not installed or embedded (Google Fonts)"); a host value locks that toggle. With downloads off, text uses installed, embedded or substitute fonts, so line breaks can differ from PowerPoint. A slow download holds the deck for up to 8 seconds, and the deck opens at once when the font server cannot be reached, so offline use works. See [font loading](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md).
- **Linked pictures and media** are not fetched when a deck opens, as in PowerPoint. A security bar offers **Enable Content**, which downloads them for that presentation only; in a shared session each participant chooses for themselves. The Trust Center option "Allow external content" (off by default) always allows them. `allowExternalContent: true` downloads them without asking, and `false` blocks them without the bar.
- **Hyperlinks** open only when the reader clicks them, always in a new tab with `noopener` and `noreferrer`, so the linked site receives no referrer and cannot script the viewer's page, whatever frame the deck names. `file:` and `ftp:` links stay in the deck but never open. Relative links (no scheme, such as `notes.html`) open, against the host page's address, only after the host calls `setOpenRelativeHyperlinks(true)` from `ooxml-ui/pptx`.
- **Share and Broadcast**, only when started. Updates go to the server you configure. With no server, peers meet through y-webrtc's public signaling server and public STUN servers; the signaling server sees a hash of the secret in the invitation link's `#key=` fragment (never the room name), the participants' IP addresses and encrypted signaling messages.

Nothing collects analytics or telemetry. The full list, including what is kept in local storage, is in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

| Package                                                                                  | Use it for                                                        |
| ---------------------------------------------------------------------------------------- | ----------------------------------------------------------------- |
| [`pptx-vue-viewer`](https://www.npmjs.com/package/pptx-vue-viewer)                       | The same viewer for Vue 3.                                        |
| [`pptx-angular-viewer`](https://www.npmjs.com/package/pptx-angular-viewer)               | The same viewer for Angular.                                      |
| [`pptx-svelte-viewer`](https://www.npmjs.com/package/pptx-svelte-viewer)                 | The same viewer for Svelte 5.                                     |
| [`pptx-vanilla-viewer`](https://www.npmjs.com/package/pptx-vanilla-viewer)               | The same viewer with no framework.                                |
| [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core)                     | Reading, building and saving decks in Node or a worker, no UI.    |
| [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp)                       | PowerPoint editing tools for AI agents, as an MCP server.         |
| [`@christophervr/pptx-viewer`](https://www.npmjs.com/package/@christophervr/pptx-viewer) | This viewer under a scoped name, plus an `npx` project installer. |

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), which also has [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) and [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) viewers.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
