# pptx-vue-viewer

**A Vue 3 component that opens, edits, presents and saves PowerPoint files (`.pptx`, `.pptm`, `.ppsx`, `.potx` and legacy `.ppt`) in the browser, with no server.**

[![npm version](https://img.shields.io/npm/v/pptx-vue-viewer.svg)](https://www.npmjs.com/package/pptx-vue-viewer)
[![license](https://img.shields.io/npm/l/pptx-vue-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)
[![types](https://img.shields.io/npm/types/pptx-vue-viewer.svg)](https://www.npmjs.com/package/pptx-vue-viewer)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo-vue/) · [Documentation](https://christophervr.github.io/ooxml/pptx/) · [Vue guide](https://christophervr.github.io/ooxml/pptx/vue/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/vue/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![Navigating between slides in the Vue 3 demo](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/vue-demo.webp)

## Install

```bash
npm install pptx-vue-viewer
npm install vue vue-i18n jszip fast-xml-parser
```

The second line is the required peers: `vue` `^3.5.39`, `vue-i18n` `^11.4.6`, `jszip` 3 and `fast-xml-parser` 5. The engine comes with it (`ooxml-core` and `ooxml-ui` are dependencies). Optional extras:

| For                                  | Add                                                                      |
| ------------------------------------ | ------------------------------------------------------------------------ |
| Real-time collaboration              | `yjs` (`y-websocket` and `y-webrtc` are optional dependencies)           |
| 3D models, 3D SmartArt and 3D charts | nothing: `three` is an optional dependency that npm installs when it can |
| The AI assistant panel               | `ai` (6 or 7) and `@ai-sdk/vue` 4                                        |

Without them those features stay off and everything else works.

## Quick start

Every label goes through `vue-i18n`, so install an instance with the viewer's English dictionary:

```ts
// main.ts
import { keyToLabel, translationsEn } from 'pptx-vue-viewer/i18n';
import 'pptx-vue-viewer/styles';
import { createApp } from 'vue';
import { createI18n } from 'vue-i18n';

import App from './App.vue';

const i18n = createI18n({
	legacy: false,
	locale: 'en',
	fallbackLocale: 'en',
	messages: { en: translationsEn },
	missing: (_locale, key) => keyToLabel(key),
});

createApp(App).use(i18n).mount('#app');
```

```vue
<!-- App.vue -->
<script setup lang="ts">
import { PowerPointViewer, type PowerPointViewerExpose } from 'pptx-vue-viewer';
import { ref } from 'vue';

const content = ref<Uint8Array | null>(null);
const viewer = ref<PowerPointViewerExpose>();

async function onPick(event: Event) {
	const file = (event.target as HTMLInputElement).files?.[0];
	if (file) content.value = new Uint8Array(await file.arrayBuffer());
}

async function save() {
	const bytes = await viewer.value?.getContent(); // the edited .pptx
}
</script>

<template>
	<div style="height: 100vh">
		<PowerPointViewer v-if="content" ref="viewer" :content="content" can-edit />
		<input v-else type="file" accept=".pptx,.ppt" @change="onPick" />
	</div>
</template>
```

`content` takes a `Uint8Array` or an `ArrayBuffer`. The viewer fills its parent, so give the parent a height. Import `pptx-vue-viewer/styles` once at your app entry.

## Props

The most used props. Every binding takes the same set; the [props reference](https://christophervr.github.io/ooxml/pptx/vue/props) has all of them.

| Prop                                             | What it does                                                                                    |
| ------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| `content`                                        | The deck's bytes (`Uint8Array` or `ArrayBuffer`).                                               |
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
| `ai`                                             | The optional assistant, from `createAiAssistant` in `pptx-vue-viewer/ai`.                       |
| `fitPadding`, `maxFitScale`                      | How the slide fits its viewport.                                                                |

The template ref (`PowerPointViewerExpose`) has `getContent`, navigation (`goTo`, `goNext`, `goPrev`), zoom (`setZoom`, `zoomIn`, `zoomOut`), `setMode`, `undo`, `redo`, `isDirty`, and the slide and element API (`getSlides`, `addSlide`, `deleteSlides`, `addElement`, `updateElement`, `deleteElements`, `selectElements`). See [the handle](https://christophervr.github.io/ooxml/pptx/vue/handle).

## Events

Emitted events, most used first: `@content-change` (the edited bytes), `@dirty-change`, `@active-slide-change`, `@selection-change`, `@load`, `@error`, `@change`, `@mode-change`, `@zoom-change`, `@slide-count-change`, `@autosave`, `@collaboration-change` and `@pdf-export-report`. A few host callbacks are props instead: `onOpenFile`, `onHyperlinkClick`, `onThemeChange`, `onLocaleChange` and `onElementRenderError`. See the [props reference](https://christophervr.github.io/ooxml/pptx/vue/props).

## Your own UI

`useViewerBuildingBlocks` runs the full viewer engine behind `BuildingBlocksCanvas`, `BuildingBlocksToolbar` and `BuildingBlocksSlidesPane`, so you can place them in your own layout and answer picker and dialog requests with your own UI; `pptx-vue-viewer/viewer` has lower-level composables. See [building blocks](https://christophervr.github.io/ooxml/pptx/vue/building-blocks) and [composables](https://christophervr.github.io/ooxml/pptx/vue/composables).

## Theming

The chrome inherits the shared `--office-*` tokens: put `data-office-theme="light"`, `"dark"` or `"auto"` on a containing element, or set tokens such as `--office-accent` there. The `theme` prop (a partial `ViewerTheme`) overrides them through the viewer's `--pptx-*` properties, and `provideViewerTheme(theme)` in a parent's `setup()` themes a subtree; `vermilionLightTheme` and `vermilionDarkTheme` are ready-made. See [theming](https://christophervr.github.io/ooxml/pptx/vue/theming).

## Localisation

English, French, Spanish, German, Simplified Chinese, Japanese and Korean. The dictionaries are at `pptx-vue-viewer/i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko`; convert one with `toVueI18nSyntax` from `pptx-vue-viewer/i18n` before passing it to `vue-i18n` (`translationsEn` is already converted), then switch with `i18n.global.locale.value = 'fr'`. See [localisation](https://christophervr.github.io/ooxml/pptx/guide/localization).

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
- **Links** (hyperlinks, Microsoft Forms responses) open only when clicked, in a new tab without opener or referrer. `file:` and `ftp:` links never open, and relative links (`notes.html`, `#part`) open, against your page's address, only after you call `setOpenRelativeHyperlinks(true)` from `ooxml-ui/pptx`.

Autosave, recovery and preferences stay on the device (IndexedDB and localStorage). The full list, including Ink to Math and the camera, is in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#powerpoint-ooxml-uipptx-and-the-pptx--viewer-packages).

## Browser support

- The last two major versions of Chrome, Edge, Firefox and Safari; Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- Vue 3.5 (3.5.39 or newer) with `vue-i18n` 11.
- Node.js 22.3 or newer to install and build.

## Documentation and demos

- [Documentation](https://christophervr.github.io/ooxml/pptx/), the [Vue guide](https://christophervr.github.io/ooxml/pptx/vue/) and the [live demo](https://christophervr.github.io/ooxml/pptx/demo-vue/).
- The same viewer for [React](https://www.npmjs.com/package/pptx-react-viewer), [Angular](https://www.npmjs.com/package/pptx-angular-viewer), [Svelte](https://www.npmjs.com/package/pptx-svelte-viewer) and [no framework](https://www.npmjs.com/package/pptx-vanilla-viewer).
- [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) for decks without a UI, and [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp) for AI agents.
- Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), with Word, Excel and Visio editors.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
