---
title: Installation
description: Installieren Sie die pptx-viewer-Pakete von npm, richten Sie die Peer-Abhängigkeiten für React, Vue 3, Angular, Svelte oder Vanilla JS ein, und führen Sie das Monorepo für die Entwicklung lokal aus.
---

# Installation

Die `pptx-viewer`-Pakete werden unabhängig voneinander auf npm veröffentlicht. Installieren Sie nur, was Sie benötigen: die framework-unabhängige [Core-Engine](/core/) oder eines der Framework-Viewer-Pakete.

::: tip Node-Version
Node.js **18 oder neuer** ist für die TypeScript-Kompilierung und für die Ausführung der Pakete außerhalb des Browsers erforderlich.
:::

## Framework wählen

| Framework                 | Paket                 | Hinweise                                                           |
| ------------------------- | --------------------- | ------------------------------------------------------------------ |
| React                     | `pptx-react-viewer`   | Viewer, Editor, Präsentator, Export, Zusammenarbeit                |
| Vue 3                     | `pptx-vue-viewer`     | Gleiche Engine und gleicher Funktionsumfang wie das React-Binding  |
| Angular                   | `pptx-angular-viewer` | Gleiche Engine und gleicher Funktionsumfang wie das React-Binding  |
| Svelte 5                  | `pptx-svelte-viewer`  | Gleiche Engine und gleicher Funktionsumfang wie das React-Binding  |
| Ohne Framework            | `pptx-vanilla-viewer` | Gleiche Engine, reines DOM: überhaupt keine Framework-Abhängigkeit |
| Headless (Node / Browser) | `pptx-viewer-core`    | Keine UI, keine Framework-Abhängigkeit                             |
| KI / MCP-Werkzeuge        | `pptx-viewer-mcp`     | 73 MCP-Tools + CLI + Y.Doc-Codec                                   |

## Von npm installieren

### React-Viewer

Die React-Viewer/Editor-Komponente, veröffentlicht als **`pptx-react-viewer`**. Die Core-Engine ist **enthalten**, Sie müssen sie nicht separat installieren.

::: code-group

```bash [npm]
npm install pptx-react-viewer react react-dom
```

```bash [pnpm]
pnpm add pptx-react-viewer react react-dom
```

```bash [yarn]
yarn add pptx-react-viewer react react-dom
```

```bash [bun]
bun add pptx-react-viewer react react-dom
```

:::

::: tip Weitere Peer-Abhängigkeiten
Der Viewer erwartet außerdem `framer-motion`, `lucide-react`, `react-icons`, `jspdf`, `jszip`, `fast-xml-parser` und `i18next`/`react-i18next` - installieren Sie die, die Ihr Einsatz benötigt.
:::

### Vue-3-Viewer

Die Vue-3-Viewer-Komponente, veröffentlicht als **`pptx-vue-viewer`**. Die Core-Engine ist enthalten.

::: code-group

```bash [npm]
npm install pptx-vue-viewer vue
```

```bash [pnpm]
pnpm add pptx-vue-viewer vue
```

```bash [yarn]
yarn add pptx-vue-viewer vue
```

```bash [bun]
bun add pptx-vue-viewer vue
```

:::

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { PowerPointViewer } from 'pptx-vue-viewer';

const content = ref<ArrayBuffer | null>(null);
</script>

<template>
	<PowerPointViewer :content="content" />
</template>
```

### Angular-Viewer

Die Angular-Viewer-Komponente, veröffentlicht als **`pptx-angular-viewer`**. Die Core-Engine ist enthalten.

::: code-group

```bash [npm]
npm install pptx-angular-viewer @angular/core @angular/common
```

```bash [pnpm]
pnpm add pptx-angular-viewer @angular/core @angular/common
```

```bash [yarn]
yarn add pptx-angular-viewer @angular/core @angular/common
```

```bash [bun]
bun add pptx-angular-viewer @angular/core @angular/common
```

:::

```typescript
// app.component.ts (standalone component, no NgModule)
import { Component } from '@angular/core';
import { PowerPointViewerComponent } from 'pptx-angular-viewer';

@Component({
	selector: 'app-root',
	standalone: true,
	imports: [PowerPointViewerComponent],
	templateUrl: './app.component.html',
})
export class AppComponent {}
```

```html
<!-- app.component.html -->
<pptx-viewer [content]="content"></pptx-viewer>
```

### Svelte-5-Viewer

Die Svelte-5-Viewer-Komponente, veröffentlicht als **`pptx-svelte-viewer`**. Die Core-Engine ist
enthalten - die vollständige Komponentenreferenz steht in der [Svelte-Dokumentation](/svelte/).

::: code-group

```bash [npm]
npm install pptx-svelte-viewer svelte
```

```bash [pnpm]
pnpm add pptx-svelte-viewer svelte
```

```bash [yarn]
yarn add pptx-svelte-viewer svelte
```

```bash [bun]
bun add pptx-svelte-viewer svelte
```

:::

```svelte
<script lang="ts">
	import { PowerPointViewer } from 'pptx-svelte-viewer';

	let bytes = $state<Uint8Array | null>(null);
</script>

{#if bytes}
	<PowerPointViewer source={bytes} />
{/if}
```

### Vanilla-JS-Viewer (ohne Framework)

Der Viewer ohne Framework, veröffentlicht als **`pptx-vanilla-viewer`**. Die Core-Engine ist enthalten;
es gibt keine Framework-Peer-Abhängigkeiten. Anzeigen, Präsentieren und Bearbeiten (hinter der Option
`editable`) funktionieren alle über eine einzige Factory-Funktion - die vollständige API-Referenz
steht in der [Vanilla-JS-Dokumentation](/vanilla/).

::: code-group

```bash [npm]
npm install pptx-vanilla-viewer
```

```bash [pnpm]
pnpm add pptx-vanilla-viewer
```

```bash [yarn]
yarn add pptx-vanilla-viewer
```

```bash [bun]
bun add pptx-vanilla-viewer
```

:::

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const viewer = createPptxViewer(document.getElementById('host')!, {
	source: '/decks/quarterly.pptx',
	editable: true,
});
```

### Core-Engine

Die framework-unabhängige Engine zum Analysieren, Bearbeiten, Serialisieren und Konvertieren von PPTX-Dateien. Verwenden Sie sie für Headless-Automatisierung, Build-Skripte oder Node.js-Pipelines ohne UI-Abhängigkeit. Die oben genannten UI-Pakete enthalten die Core-Engine bereits, Sie müssen sie also nicht separat installieren, wenn Sie schon eines davon nutzen.

::: code-group

```bash [npm]
npm install pptx-viewer-core
```

```bash [pnpm]
pnpm add pptx-viewer-core
```

```bash [yarn]
yarn add pptx-viewer-core
```

```bash [bun]
bun add pptx-viewer-core
```

:::

### MCP-Server und Werkzeuge

73 PPTX-Manipulations-Tool-Funktionen, ein MCP-Server für KI-Agenten und der Y.Doc-Zusammenarbeit-Codec - alles auf der Core-Engine aufgebaut.

::: code-group

```bash [npm]
npm install pptx-viewer-mcp
```

```bash [pnpm]
pnpm add pptx-viewer-mcp
```

```bash [yarn]
yarn add pptx-viewer-mcp
```

```bash [bun]
bun add pptx-viewer-mcp
```

:::

## Optionale Peer-Abhängigkeiten

Einige Funktionen im React-Paket werden nur aktiviert, wenn die optionalen Peers vorhanden sind.

| Funktion                    | Optionale Peers      | Hinweise                                                |
| --------------------------- | -------------------- | ------------------------------------------------------- |
| **3D-Modelle** (GLB/GLTF)   | `three`              | Ohne sie greifen 3D-Elemente auf ihr Posterbild zurück. |
| **Echtzeit-Zusammenarbeit** | `yjs`, `y-websocket` | Yjs-CRDT mit Anwesenheitsanzeige.                       |

::: code-group

```bash [npm]
npm install three yjs y-websocket
```

```bash [pnpm]
pnpm add three yjs y-websocket
```

```bash [yarn]
yarn add three yjs y-websocket
```

```bash [bun]
bun add three yjs y-websocket
```

:::

## Lokale Entwicklung (Monorepo klonen)

Der PowerPoint-Viewer liegt im Monorepo `ooxml` unter `viewers/pptx`. Das Monorepo verwendet **Bun** als Paketmanager und Workspace-Runner.

```bash
# Clone the repository and install every workspace dependency from the root
git clone https://github.com/ChristopherVR/ooxml.git
cd ooxml
bun install

# Build ooxml-core and ooxml-ui first, then work from the viewer
bun run build
bun run --cwd src/ui build
cd viewers/pptx

# Build the viewer packages, then test / type-check
bun run build
bun run test
bun run typecheck
```

::: warning Build-Reihenfolge ist wichtig
Die Viewer-Pakete verwenden die gebauten Pakete `ooxml-core` und `ooxml-ui`; bauen Sie diese also zuerst im Repository-Stamm (wie oben). Innerhalb von `viewers/pptx` werden die Pakete in Abhängigkeitsreihenfolge gebaut:

```
core -> shared -> locales -> tools -> react / vue / angular / vanilla / svelte -> cli
```

`bun run build` in `viewers/pptx` übernimmt das für Sie. Wenn Sie ein einzelnes Paket manuell bauen (`cd packages/<pkg> && bun run build`), stellen Sie sicher, dass seine Abhängigkeiten zuerst gebaut sind.
:::

### Häufige Workspace-Befehle

```bash
bun run build         # Build all packages in dependency order
bun run test          # Run vitest across all packages
bun run typecheck     # Type-check all packages (needs built packages)
bun run typecheck:fresh  # Build the packages, then type-check (clean checkouts)
bun run fmt           # Format with oxfmt
bun run lint          # Lint with oxlint
bun run demo          # Start the React demo dev server (port 4173)
bun run demo:vue      # Start the Vue demo dev server (port 4175)
bun run demo:angular  # Start the Angular demo dev server (port 4174)
bun run demo:vanilla  # Start the Vanilla JS demo dev server (port 4176)
bun run demo:svelte   # Start the Svelte demo dev server (port 4177)
```

## Nächste Schritte

- [Schnellstart](/de/guide/quick-start) - Präsentationen erstellen, analysieren, konvertieren und rendern.
- [Architektur](/de/guide/architecture) - wie die Schichten zusammenpassen.
- [Einschränkungen](/de/guide/limitations) - wichtige Hinweise vor dem Produktionseinsatz.
