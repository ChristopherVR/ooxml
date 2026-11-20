# pptx-angular-viewer

A standalone Angular component that opens, edits, presents and saves PowerPoint (`.pptx`) files in the browser. No server, no conversion step, no PowerPoint install.

[![npm version](https://img.shields.io/npm/v/pptx-angular-viewer.svg)](https://www.npmjs.com/package/pptx-angular-viewer)
[![license](https://img.shields.io/npm/l/pptx-angular-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo-angular/) · [Docs](https://christophervr.github.io/ooxml/pptx/angular/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/angular/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![Exploring the Insert and View ribbon tabs in the Angular demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/angular-demo.webp)

Slides are drawn as HTML, CSS and SVG, not as a canvas bitmap, so text stays sharp at any zoom, can be selected and is read by screen readers. The component is built on signals and `OnPush`, works without Zone.js, and needs no NgModule. The [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) engine is installed with it.

## Install

```bash
npm install pptx-angular-viewer
npm install @angular/core @angular/common @angular/platform-browser rxjs @ngx-translate/core
```

The second line is the peer dependencies: Angular 19, 20, 21 or 22, `rxjs` 7 and `@ngx-translate/core` 18. Optional peers:

| For                                  | Add                                    |
| ------------------------------------ | -------------------------------------- |
| Real-time collaboration              | `yjs` with `y-websocket` or `y-webrtc` |
| 3D models, 3D SmartArt and 3D charts | `three`                                |
| The AI assistant panel               | `ai`                                   |

Without them, those features stay off and everything else works.

## Quick start

Add the stylesheet to the `styles` array of your build target in `angular.json` (or `@import` it at the top of `src/styles.css`):

```json
"styles": ["pptx-angular-viewer/styles.css", "src/styles.css"]
```

Do not import it from TypeScript: the Angular CLI's application builder writes CSS imported that way to a file `index.html` never links, and the viewer renders unstyled.

Every label goes through `@ngx-translate/core`, so provide it and register the viewer's English dictionary:

```ts
// main.ts
import { bootstrapApplication } from '@angular/platform-browser';
import { provideTranslateService } from '@ngx-translate/core';

import { AppComponent } from './app.component';

bootstrapApplication(AppComponent, {
	providers: [provideTranslateService({ lang: 'en', fallbackLang: 'en' })],
});
```

```ts
// app.component.ts
import { Component, inject, signal, viewChild } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { PowerPointViewerComponent, translationsEn } from 'pptx-angular-viewer';

@Component({
	selector: 'app-root',
	imports: [PowerPointViewerComponent],
	template: `
		<div style="height: 100vh">
			@if (content()) {
				<pptx-viewer [content]="content()" [canEdit]="true" />
			} @else {
				<input type="file" accept=".pptx,.ppt" (change)="onPick($event)" />
			}
		</div>
	`,
})
export class AppComponent {
	readonly content = signal<Uint8Array | null>(null);
	private readonly viewer = viewChild(PowerPointViewerComponent);

	constructor() {
		inject(TranslateService).setTranslation('en', translationsEn);
	}

	async onPick(event: Event) {
		const file = (event.target as HTMLInputElement).files?.[0];
		if (file) this.content.set(new Uint8Array(await file.arrayBuffer()));
	}

	async save() {
		const bytes = await this.viewer()?.getContent(); // a valid .pptx
	}
}
```

`content` takes a `Uint8Array` or an `ArrayBuffer`. The viewer fills its parent, so give the parent a height.

## What you get

- **Viewing**: text, all 187 preset shapes, pictures, tables, charts, SmartArt, connectors, audio, video, ink, OLE objects and 3D models, with themes, masters and layouts resolved.
- **Editing**: a PowerPoint-style ribbon, inspector and on-canvas handles; inline text, tables, chart data, SmartArt, masters and layouts; full undo and redo.
- **Presenting**: transitions (including Morph), animations and motion paths, custom shows, presenter view with notes and timer, rehearse timings, pen, highlighter and laser.
- **Exporting**: PDF, PNG, GIF, video and JSON, slides copied as images, printed handouts and notes pages, and Save As `.pptx`, `.ppsx`, `.pptm` or legacy `.ppt`.
- **Working together**: real-time co-editing and Broadcast over Yjs, through your server or peer to peer, with presence, cursors and threaded comments.
- **Assistant**: an optional AI panel whose edits follow a write policy you choose (apply, stage or approve).
- **Keyboard and screen readers**: F6 between regions, Tab through slide objects, Enter to edit text, and names that match the visible captions.
- **Seven UI languages**: English, French, Spanish, German, Simplified Chinese, Japanese and Korean.

Some effects are approximated or not yet supported; the [limitations page](https://christophervr.github.io/ooxml/pptx/guide/limitations) lists them, and a deck that uses one reports it as a load warning.

## Using it in Angular

**Inputs, outputs and methods.** Changes arrive as outputs (`contentChange`, `dirtyChange`, `activeSlideChange`, `selectionChange` and more). The component instance covers navigation, zoom, mode, selection, undo and redo, plus the shared slide and element API (`getSlides`, `addSlide`, `addElement`, `updateElement` and the rest). Host callbacks such as `onOpenFile` and `onLocaleChange` are inputs. See [inputs and outputs](https://christophervr.github.io/ooxml/pptx/angular/props) and [the API](https://christophervr.github.io/ooxml/pptx/angular/api).

**Theming.** Pass a partial `[theme]`, or provide one for the whole app with `provideViewerTheme(theme)` in `bootstrapApplication`. `vermilionLightTheme` and `vermilionDarkTheme` are ready-made presets. See [theming](https://christophervr.github.io/ooxml/pptx/angular/theming).

**Your own shell.** `createViewerBuildingBlocks` with `VIEWER_BUILDING_BLOCKS_PROVIDERS` runs the viewer's editor services behind `<pptx-building-blocks-canvas>`, `<pptx-building-blocks-toolbar>` and `<pptx-slides-pane>`, so you can place them in your own layout and answer picker and dialog requests with your own UI. See [building blocks](https://christophervr.github.io/ooxml/pptx/angular/building-blocks) and the [services reference](https://christophervr.github.io/ooxml/pptx/angular/services-reference).

**Translations.** `translationsEn`, `keyToLabel` (for a `MissingTranslationHandler`) and `LOCALE_CATALOG` come from the package root. Dictionaries for the other languages are at `pptx-angular-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`; register them with `setTranslation` and switch with `use('fr')`. See [localization](https://christophervr.github.io/ooxml/pptx/guide/localization).

**AI assistant.** Create one assistant with `createAiAssistant` from `pptx-angular-viewer/ai` and pass it to `[ai]`. Its chat engine, tool catalogue and schemas load only when the panel starts, so they stay out of your initial chunks. See the [assistant guide](https://christophervr.github.io/ooxml/pptx/guide/ai-assistant).

## Configuration

The most used inputs; the [inputs reference](https://christophervr.github.io/ooxml/pptx/angular/props) has all of them, and every binding takes the same set.

| Input                                            | What it does                                                                                                     |
| ------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------- |
| `canEdit`, `readOnly`                            | Turn editing on, or force the deck read-only.                                                                    |
| `initialSlide`, `showToolbar`, `showThumbnails`  | Starting slide and which chrome to show.                                                                         |
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
- Angular 19, 20, 21 or 22, with or without Zone.js.
- Node.js 22.3 or newer to install, build and test (the Angular CLI may ask for more).

## Privacy and network use

The viewer renders and saves decks in the browser and never uploads them. It contacts other servers only in these cases:

- **Fonts that are neither installed nor embedded.** When a deck uses such a font and Google Fonts serves it (or a metric-compatible clone, such as Carlito for Calibri), the viewer downloads it from `fonts.googleapis.com` and `fonts.gstatic.com`, so the reader's IP address and the needed family names reach Google. This is on by default, as PowerPoint downloads its cloud fonts. Turn it off with `googleFonts: false`, or send the same requests to your own mirror with `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }`. Readers can turn it off in File > Options > Trust Center ("Download fonts that are not installed or embedded (Google Fonts)"); a host value locks that toggle. With downloads off, text uses installed, embedded or substitute fonts, so line breaks can differ from PowerPoint. A slow download holds the deck for up to 8 seconds, and the deck opens at once when the font server cannot be reached, so offline use works. See [font loading](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md).
- **Linked pictures and media** are not fetched when a deck opens, as in PowerPoint. A security bar ("SECURITY WARNING", Enable Content) offers to download them for that presentation only; in a shared session each participant chooses for themselves. The Trust Center option "Allow external content" (off by default) always allows them. `allowExternalContent: true` downloads them without the bar, and `false` blocks them without it.
- **Hyperlinks** open only when the reader clicks them, always in a new tab with `noopener` and `noreferrer`, so the linked site receives no referrer and cannot script the viewer's page, whatever frame the deck names. `file:` and `ftp:` links stay in the deck but never open. Relative links (no scheme, such as `notes.html`) open, against the host page's address, only after the host calls `setOpenRelativeHyperlinks(true)` from `ooxml-ui/pptx`.
- **Share and Broadcast**, only when started. Updates go to the server you configure. With no server, peers meet through y-webrtc's public signaling server and public STUN servers; the signaling server sees a hash of the secret in the invitation link's `#key=` fragment (never the room name), the participants' IP addresses and encrypted signaling messages.

Nothing collects analytics or telemetry. The full list, including what is kept in local storage, is in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

| Package                                                                    | Use it for                                                     |
| -------------------------------------------------------------------------- | -------------------------------------------------------------- |
| [`pptx-react-viewer`](https://www.npmjs.com/package/pptx-react-viewer)     | The same viewer for React.                                     |
| [`pptx-vue-viewer`](https://www.npmjs.com/package/pptx-vue-viewer)         | The same viewer for Vue 3.                                     |
| [`pptx-svelte-viewer`](https://www.npmjs.com/package/pptx-svelte-viewer)   | The same viewer for Svelte 5.                                  |
| [`pptx-vanilla-viewer`](https://www.npmjs.com/package/pptx-vanilla-viewer) | The same viewer with no framework.                             |
| [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core)       | Reading, building and saving decks in Node or a worker, no UI. |
| [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp)         | PowerPoint editing tools for AI agents, as an MCP server.      |

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), which also has [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) and [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) viewers.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
