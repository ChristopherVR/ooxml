---
title: Installation
description: Installez les packages pptx-viewer depuis npm, configurez les dépendances pour React, Vue 3, Angular, Svelte ou JavaScript natif, et exécutez le monorepo localement pour le développement.
---

# Installation

Les packages `pptx-viewer` sont publiés indépendamment sur npm. Installez uniquement ce dont vous avez besoin : le [moteur principal](/core/) indépendant du framework ou l'un des packages de visionneuse par framework.

::: tip Version de Node
Node.js **18 ou supérieur** est requis pour la compilation TypeScript et pour exécuter les packages en dehors du navigateur.
:::

## Choisir votre framework

| Framework                    | Package               | Notes                                                     |
| ---------------------------- | --------------------- | --------------------------------------------------------- |
| React                        | `pptx-react-viewer`   | Visionneuse, éditeur, présentateur, export, collaboration |
| Vue 3                        | `pptx-vue-viewer`     | Même moteur et mêmes fonctionnalités que le binding React |
| Angular                      | `pptx-angular-viewer` | Même moteur et mêmes fonctionnalités que le binding React |
| Svelte 5                     | `pptx-svelte-viewer`  | Même moteur et mêmes fonctionnalités que le binding React |
| Sans framework               | `pptx-vanilla-viewer` | Même moteur, DOM natif : aucune dépendance de framework   |
| Headless (Node / navigateur) | `pptx-viewer-core`    | Aucune interface, aucune dépendance de framework          |
| Outils IA / MCP              | `pptx-viewer-mcp`     | 73 outils MCP + CLI + codec Y.Doc                         |

## Installation depuis npm

### Visionneuse React

Le composant visionneuse/éditeur React, publié sous **`pptx-react-viewer`**. Le moteur principal est **intégré**, vous n'avez pas à l'installer séparément.

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

::: tip Autres dépendances peer
La visionneuse attend aussi `framer-motion`, `lucide-react`, `react-icons`, `jspdf`, `jszip`, `fast-xml-parser` et `i18next`/`react-i18next` : installez ceux dont votre usage a besoin.
:::

### Visionneuse Vue 3

Le composant visionneuse Vue 3, publié sous **`pptx-vue-viewer`**. Le moteur principal est intégré.

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

### Visionneuse Angular

Le composant visionneuse Angular, publié sous **`pptx-angular-viewer`**. Le moteur principal est intégré.

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

### Visionneuse Svelte 5

Le composant visionneuse Svelte 5, publié sous **`pptx-svelte-viewer`**. Le moteur principal est intégré : voir la [documentation Svelte](/svelte/) pour la référence complète des composants.

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

### Visionneuse JavaScript natif (sans framework)

La visionneuse sans framework, publiée sous **`pptx-vanilla-viewer`**. Le moteur principal est intégré ; il n'y a aucune dépendance peer de framework. L'affichage, la présentation et l'édition (derrière l'option `editable`) fonctionnent tous depuis une seule fonction de fabrique : voir la [documentation Vanilla JS](/vanilla/) pour la référence complète de l'API.

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

### Moteur principal

Le moteur indépendant du framework pour analyser, éditer, sérialiser et convertir des fichiers PPTX. Utilisez-le pour l'automatisation headless, les scripts de build ou les pipelines Node.js sans dépendance d'interface. Les packages d'interface ci-dessus intègrent le moteur principal ; vous n'avez donc pas à l'installer séparément si vous utilisez déjà l'un d'eux.

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

### Serveur MCP et outils

73 fonctions d'outils de manipulation PPTX, un serveur MCP pour agents IA et le codec de collaboration Y.Doc, le tout construit sur le moteur principal.

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

## Dépendances peer optionnelles

Certaines fonctionnalités du package React ne s'activent que lorsque leurs dépendances peer optionnelles sont présentes.

| Fonctionnalité                  | Dépendances peer optionnelles | Notes                                                              |
| ------------------------------- | ----------------------------- | ------------------------------------------------------------------ |
| **Modèles 3D** (GLB/GLTF)       | `three`                       | Sans elles, les éléments 3D se rabattent sur leur image d'affiche. |
| **Collaboration en temps réel** | `yjs`, `y-websocket`          | Yjs CRDT avec suivi de présence.                                   |

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

## Développement local (clonage du monorepo)

La visionneuse PowerPoint vit dans le monorepo `ooxml` sous `viewers/pptx`. Le monorepo utilise **Bun** comme gestionnaire de packages et exécuteur de workspaces.

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

::: warning L'ordre de construction est important
Les packages de visionneuse consomment `ooxml-core` et `ooxml-ui` construits ; construisez-les donc d'abord à la racine du dépôt (comme ci-dessus). Dans `viewers/pptx`, les packages se construisent dans l'ordre des dépendances :

```
core -> shared -> locales -> tools -> react / vue / angular / vanilla / svelte -> cli
```

`bun run build` dans `viewers/pptx` s'en charge pour vous. Si vous construisez un seul package manuellement (`cd packages/<pkg> && bun run build`), assurez-vous que ses dépendances sont construites d'abord.
:::

### Commandes courantes du workspace

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

## Étapes suivantes

- [Démarrage rapide](/fr/guide/quick-start) - créer, analyser, convertir et afficher des présentations.
- [Architecture](/fr/guide/architecture) - comment les couches s'articulent.
- [Limitations](/fr/guide/limitations) - mises en garde importantes avant la mise en production.
