---
title: Instalación
description: Instale los paquetes pptx-viewer desde npm, configure las dependencias entre pares para React, Vue 3, Angular, Svelte o vanilla JS, y ejecute el monorepo localmente para desarrollar.
---

# Instalación

Los paquetes `pptx-viewer` se publican de forma independiente en npm. Instale solo lo que necesite: el [motor core](/core/) independiente del framework o uno de los paquetes de visualizador por framework.

::: tip Versión de Node
Se requiere Node.js **18 o superior** para la compilación de TypeScript y para ejecutar los paquetes fuera del navegador.
:::

## Elija su framework

| Framework                   | Paquete               | Notas                                                        |
| --------------------------- | --------------------- | ------------------------------------------------------------ |
| React                       | `pptx-react-viewer`   | Visualizador, editor, presentador, exportación, colaboración |
| Vue 3                       | `pptx-vue-viewer`     | Mismo motor y conjunto de funciones que el enlace de React   |
| Angular                     | `pptx-angular-viewer` | Mismo motor y conjunto de funciones que el enlace de React   |
| Svelte 5                    | `pptx-svelte-viewer`  | Mismo motor y conjunto de funciones que el enlace de React   |
| Sin framework               | `pptx-vanilla-viewer` | Mismo motor, DOM puro: sin ninguna dependencia de framework  |
| Headless (Node / navegador) | `pptx-viewer-core`    | Sin interfaz, sin dependencia de framework                   |
| Herramientas IA / MCP       | `pptx-viewer-mcp`     | 73 herramientas MCP + CLI + códec Y.Doc                      |

## Instalación desde npm

### Visualizador React

El componente visualizador/editor de React, publicado como **`pptx-react-viewer`**. El motor core está **incluido**, por lo que no necesita instalarlo por separado.

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

::: tip Otras dependencias entre pares
El visualizador también espera `framer-motion`, `lucide-react`, `react-icons`, `jspdf`, `jszip`, `fast-xml-parser` e `i18next`/`react-i18next`: instale las que necesite su uso.
:::

### Visualizador Vue 3

El componente visualizador de Vue 3, publicado como **`pptx-vue-viewer`**. El motor core está incluido.

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

### Visualizador Angular

El componente visualizador de Angular, publicado como **`pptx-angular-viewer`**. El motor core está incluido.

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

### Visualizador Svelte 5

El componente visualizador de Svelte 5, publicado como **`pptx-svelte-viewer`**. El motor core está
incluido; consulte la [documentación de Svelte](/svelte/) para la referencia completa del componente.

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

### Visualizador Vanilla JS (sin framework)

El visualizador sin framework, publicado como **`pptx-vanilla-viewer`**. El motor core está incluido;
no hay dependencias entre pares de framework. Ver, presentar y editar (tras la opción
`editable`) funcionan desde una única función fábrica; consulte la
[documentación de Vanilla JS](/vanilla/) para la referencia completa de la API.

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

### Motor core

El motor independiente del framework para analizar, editar, serializar y convertir archivos PPTX. Úselo cuando necesite automatización sin interfaz, scripts de compilación o canalizaciones de Node.js sin dependencia de UI. Los paquetes de interfaz anteriores incluyen el motor core, por lo que no necesita instalarlo por separado si ya usa uno de ellos.

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

### Servidor MCP y herramientas

73 funciones de herramientas de manipulación de PPTX, un servidor MCP para agentes de IA y el códec de colaboración Y.Doc, todo construido sobre el motor core.

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

## Dependencias opcionales entre pares

Algunas funciones del paquete de React solo se activan cuando están presentes sus dependencias opcionales.

| Función                         | Dependencias opcionales | Notas                                                       |
| ------------------------------- | ----------------------- | ----------------------------------------------------------- |
| **Modelos 3D** (GLB/GLTF)       | `three`                 | Sin ellas, los elementos 3D recurren a su imagen de póster. |
| **Colaboración en tiempo real** | `yjs`, `y-websocket`    | Yjs CRDT con seguimiento de presencia.                      |

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

## Desarrollo local (clonar el monorepo)

El visualizador de PowerPoint vive en el monorepo `ooxml`, bajo `viewers/pptx`. El monorepo usa **Bun** como gestor de paquetes y ejecutor de workspaces.

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

::: warning El orden de compilación importa
Los paquetes del visualizador consumen `ooxml-core` y `ooxml-ui` ya compilados, así que compílelos primero en la raíz del repositorio (como arriba). Dentro de `viewers/pptx` los paquetes se compilan en orden de dependencias:

```
core -> shared -> locales -> tools -> react / vue / angular / vanilla / svelte -> cli
```

`bun run build` en `viewers/pptx` se encarga de esto. Al compilar un paquete a mano (`cd packages/<pkg> && bun run build`), asegúrese de que sus dependencias estén compiladas primero.
:::

### Comandos comunes del workspace

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

## Siguientes pasos

- [Inicio rápido](/es/guide/quick-start) - crear, analizar, convertir y renderizar presentaciones.
- [Arquitectura](/es/guide/architecture) - cómo encajan las capas.
- [Limitaciones](/es/guide/limitations) - advertencias importantes antes de pasar a producción.
