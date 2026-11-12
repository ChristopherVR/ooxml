# Bloques de construcción

`docx-react-viewer` exporta `useViewerBuildingBlocks`, `DocumentCanvas` y los componentes opcionales
`Toolbar`, `PageNavigation` y `StatusBar`. Usan el motor existente de WordEditor para edición, diseño,
historial y serialización, y siguen la forma de `useViewerBuildingBlocks` de PowerPoint: el hook
devuelve propiedades y el anfitrión renderiza lo que quiera.

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, navigationProps, statusProps, loading, error } =
		useViewerBuildingBlocks({
			documentModel,
			canEdit,
			handle,
			onDirtyChange: (dirty) => console.log(dirty),
			// Host-owned dialogs: omit one and its action reports as not enabled.
			onOpenHyperlinkDialog: (request) => openMyLinkDialog(request),
			onOpenFindReplace: (request) => openMyFindBar(request),
			onOpenImagePicker: (request) => openMyFilePicker(request),
		});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{statusProps && <span>{statusProps.wordText}</span>}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Muestre el lienzo incluso mientras carga. Las propiedades son null hasta que se monta. Dé una altura
al contenedor del lienzo.

## Lo que devuelve el hook {#what-the-hook-returns}

| Campo                      | Contenido                                                                         |
| -------------------------- | --------------------------------------------------------------------------------- |
| `canvasProps`              | Se distribuye sobre `DocumentCanvas`. Monta y libera el editor.                   |
| `toolbarProps`             | Estado de selección y formato, más los callbacks de edición (abajo).              |
| `navigationProps`          | Página, número de páginas, modo de vista y callbacks de navegación entre páginas. |
| `statusProps`              | Recuento de palabras y zoom, con sus callbacks.                                   |
| `loading`, `error`, `mode` | Estado de carga, el último error de carga o de documento, `'view'` o `'edit'`.    |

### toolbarProps {#toolbarprops}

Estado: `selection` (`{ from, to }`), `bold`, `italic`, `underline` (`null` cuando la selección mezcla
valores), `fontFamily`, `fontSize` (`null` cuando es mixto), `alignment`
(`'left' | 'center' | 'right' | 'justify'` o `null`), `list`
(`'bullet' | 'number' | 'other' | null`), `canEdit`, `canUndo`, `canRedo` y `labels` traducidas.

Callbacks, cada uno devuelve `false` cuando la acción no está permitida (solo lectura o todavía cargando):
`onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange(family)`, `onFontSizeChange(points)`,
`onAlignmentChange(alignment)`, `onListChange('bullet' | 'number' | 'remove')`, `onUndo`, `onRedo`.
El estado `list` y `onListChange` usan los mismos valores `'bullet' | 'number'`.

### navigationProps {#navigationprops}

`page`, `pageCount`, `pageText` ("Página 2 de 5"), `pagesAvailable`, `canGoPrevious`, `canGoNext`,
`viewMode` (`'draft' | 'print'`) y los callbacks `onGoToPage(n)`, `onPreviousPage`,
`onNextPage` y `onViewModeChange(mode)`. Las páginas solo existen en el diseño de impresión: en
borrador, `pagesAvailable` es false y los callbacks de página devuelven `false`. Cambie con
`onViewModeChange('print')`.

### statusProps {#statusprops}

`words`, `wordText` ("42 palabras"), `zoom` (porcentaje), `minZoom`, `maxZoom` y los callbacks
`onZoomChange(percent)`, `onZoomIn` y `onZoomOut`. El desplazamiento también actualiza `navigationProps`.

## Diálogos del anfitrión {#host-dialogs}

Una superficie no monta ningún diálogo propio. Estos flujos de trabajo pertenecen al anfitrión y se
llaman como los callbacks `onOpen...` de PowerPoint. Páselos como opciones; `toolbarProps` incluye
entonces un botón `open...` correspondiente (`openHyperlinkDialog`, `openFindReplace`, `openImagePicker`),
un indicador `can...` y `hostWorkflows` (qué callbacks se han proporcionado).

| Opción                  | Objeto de solicitud                                                              | Habilitado mientras                          |
| ----------------------- | -------------------------------------------------------------------------------- | -------------------------------------------- |
| `onOpenHyperlinkDialog` | `{ href, text, apply(href, text?), remove() }` para el vínculo de la selección   | editable                                     |
| `onOpenFindReplace`     | `{ search, findNext, findPrevious, replaceOne, replaceAll, status, canReplace }` | cargado (en solo lectura se mantiene buscar) |
| `onOpenImagePicker`     | `{ insert(file) }`: el anfitrión elige un archivo y se lo devuelve               | editable                                     |

Si se omite un callback, `canInsertLink`, `canFindReplace` o `canInsertPicture` es false y el
botón devuelve `false`. Las solicitudes vuelven a comprobar el permiso al usarse, de modo que un diálogo
dejado abierto al cambiar a solo lectura no puede editar.

## Componentes de la biblioteca {#library-components}

`Toolbar` dibuja deshacer y rehacer, familia y tamaño de fuente, negrita, cursiva, subrayado, cuatro
alineaciones, listas con viñetas y numeradas, y un botón por cada flujo de trabajo del anfitrión que
haya proporcionado. `PageNavigation` dibuja la página anterior, "Página n de m", la página siguiente y el
selector Borrador o Diseño de impresión. `StatusBar` dibuja el recuento de palabras y un selector de
zoom. Todos son opcionales y se controlan por propiedades; cada control tiene un nodo estable, de modo que
el foco del teclado sobrevive a las actualizaciones de formato. Usan los estilos compartidos de Office.
Los controles personalizados nunca los necesitan.

```tsx
{
	toolbarProps && <Toolbar {...toolbarProps} />;
}
{
	navigationProps && <PageNavigation {...navigationProps} />;
}
{
	statusProps && <StatusBar {...statusProps} />;
}
```

## Superficies de solo lectura {#read-only-surfaces}

La superficie es **de solo lectura hasta que el anfitrión pase `canEdit: true`**, igual que los bloques de
construcción de PowerPoint y Visio (el `WordEditor` completo sigue siendo editable por defecto). El permiso
se aplica antes de montar el lienzo, de modo que `mode` empieza como `'view'` y no hay un destello
editable. `readOnly` sigue prevaleciendo sobre `canEdit: true`. Mientras está en solo lectura, cada callback
de edición devuelve `false` y el documento no puede cambiar. El lector todavía puede dar foco a la
superficie, seleccionar texto con el ratón o el teclado (flechas, Mayús+flechas, Inicio, Fin, Ctrl+A),
copiar, y leer la selección desde `toolbarProps.selection`. Los atajos de teclado que pertenecen a la
aplicación (buscar, guardar, imprimir) no se instalan en una superficie; conéctelos a sus propios controles.

**Migrar un anfitrión.** Antes, los bloques de construcción eran editables por defecto. Un anfitrión que
dependía de eso debe pasar ahora `canEdit: true` (en el objeto de opciones o, para un adaptador nativo,
en las opciones reactivas); sin ello, la superficie se abre en solo lectura y todos los callbacks de
edición devuelven `false`.

## Ciclo de vida {#lifecycle}

Use `bytes` para cargar un archivo, o los métodos del manejador `load`, `save`, `download`, `markClean` y
`dirty`. Guardar no marca el documento como limpio hasta que el anfitrión confirma la persistencia.
El lienzo no monta barra de título, cinta, backstage, barra de estado ni diálogos de la aplicación.

Cada paquete de framework también exporta `mountEditorSurface(host, options)` para adaptadores de ciclo
de vida nativo. Su binding expone `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` y
`subscribe(listener)`, además del manejador del editor, `update` y `destroy`. Las opciones de superficie
excluyen las opciones exclusivas de la cinta (`showToolbar`, `showThumbnails`, `hiddenActions`, `ribbonAddIns`,
`onRibbonAddIn`), de modo que un anfitrión en TypeScript no puede pasar una que se ignoraría.

La API de composición requiere el runtime `ooxml-ui` correspondiente, publicado con estas adiciones.
Las importaciones existentes del visor completo siguen siendo compatibles con el runtime publicado
anterior. Un runtime más antiguo rechaza explícitamente una solicitud de composición en lugar de montar
la interfaz de la aplicación por accidente.

## Composición nativa con frameworks {#native-framework-composition}

Los adaptadores de framework viven en su propio paquete de framework. Comparten el auxiliar de duración
del motor y los descriptores de controles (`bindings/src/controls.ts`) que dibujan los componentes
opcionales. Cada adaptador expone `toolbarProps`, `navigationProps`, `statusProps`, `loading`, `error`,
`mode` y `canvasProps`, además del `handle` del editor (React recibe un `handle` como ref de opción en
lugar de devolverlo), y los componentes `Toolbar`, `PageNavigation` y `StatusBar` (Angular:
`ToolbarComponent`, `PageNavigationComponent`, `StatusBarComponent`; Svelte: `docx-svelte-viewer/Toolbar`,
`/PageNavigation`, `/StatusBar`). Su estado usa las primitivas reactivas normales de cada framework, y sus
componentes de lienzo montan la superficie de documento de Word existente.

### Vue {#vue}

```vue
<script setup lang="ts">
import { computed } from 'vue';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-vue-viewer';
const props = defineProps<{ documentModel: import('docx-core').DocumentModel; canEdit: boolean }>();
const { canvasProps, toolbarProps, error, handle } = useViewerBuildingBlocks(
	computed(() => ({ documentModel: props.documentModel, canEdit: props.canEdit })),
);
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<DocumentCanvas v-bind="canvasProps" class="document-surface" />
	<p v-if="error" role="alert">{{ error.message }}</p>
</template>
```

El composable observa una ref, una ref computada o un getter, y libera el motor junto con su ámbito.
`toolbarProps`, `loading`, `error` y `mode` son refs computadas de solo lectura. `handle` expone la API
imperativa existente del editor.

### Angular {#angular}

Proporcione un servicio `ViewerBuildingBlocks` por componente anfitrión. Las señales leídas en la plantilla
participan en la detección de cambios de Angular, incluidas las aplicaciones sin zone.js. `DestroyRef`
libera el motor cuando se destruye su proveedor.

```ts
import { Component, signal } from '@angular/core';
import { createDocument } from 'docx-core';
import {
	createViewerBuildingBlocks,
	ViewerBuildingBlocks,
	DocumentCanvasComponent,
	ToolbarComponent,
} from 'docx-angular-viewer';

@Component({
	selector: 'embedded-word',
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent],
	providers: [ViewerBuildingBlocks],
	template: `
		@if (blocks.toolbarProps(); as toolbar) {
			<word-toolbar [toolbarProps]="toolbar" />
		}
		<word-document-canvas [canvasProps]="blocks.canvasProps" />
	`,
})
export class EmbeddedWord {
	readonly canEdit = signal(true);
	readonly model = createDocument();
	readonly blocks = createViewerBuildingBlocks(() => ({
		documentModel: this.model,
		canEdit: this.canEdit(),
	}));
}
```

El servicio expone las señales de solo lectura `toolbarProps`, `loading`, `error` y `mode`, un `handle`
estable, y emisores de eventos tipados `documentChange`, `documentError`, `dirtyChange` y `pageChange`.
Suscríbase a esos emisores o proporcione los callbacks de opción existentes. Para anfitriones
imperativos, inyecte el servicio y llame a `update(options)` en lugar de usar el auxiliar.

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
  import { writable } from 'svelte/store';
  import { createDocument } from 'docx-core';
  import { createViewerBuildingBlocks } from 'docx-svelte-viewer/runtime';
  import DocumentCanvas from 'docx-svelte-viewer/DocumentCanvas';
  import Toolbar from 'docx-svelte-viewer/Toolbar';
  const options = writable({ documentModel: createDocument(), canEdit: true });
  const blocks = createViewerBuildingBlocks(options);
  const { toolbarProps, error } = blocks;
</script>
{#if $toolbarProps}<Toolbar toolbarProps={$toolbarProps} />{/if}
<DocumentCanvas {...blocks.canvasProps} />
{#if $error}<p role="alert">{$error.message}</p>{/if}
```

Llame a la fábrica durante la inicialización del componente. Acepta un store de opciones de solo lectura o un
objeto de opciones inicial; `blocks.update(next)` también admite actualizaciones de `$effect` basadas en runes.
Las instantáneas son stores de solo lectura. `documentCanvas` se exporta como acción para un marcado
personalizado del anfitrión. La destrucción del componente se da de baja de las opciones y destruye el motor.

### Solid {#solid}

```tsx
import { createSignal, Show } from 'solid-js';
import { createDocument } from 'docx-core';
import { createViewerBuildingBlocks, DocumentCanvas, Toolbar } from 'docx-solid-viewer';
function EmbeddedWord() {
	const [canEdit, setCanEdit] = createSignal(true);
	const model = createDocument();
	const blocks = createViewerBuildingBlocks(() => ({ documentModel: model, canEdit: canEdit() }));
	return (
		<>
			<Show when={blocks.toolbarProps()}>{(toolbar) => <Toolbar toolbarProps={toolbar()} />}</Show>
			<DocumentCanvas {...blocks.canvasProps} />
		</>
	);
}
```

La fábrica sigue su accesor de opciones en un efecto y devuelve accesores de señal para
`toolbarProps`, `loading`, `error` y `mode`. Su propietario controla la limpieza mediante `onCleanup`.
Los controles opcionales de la barra de herramientas mantienen sus nodos del DOM mientras cambia el estado
de formato, de modo que se conserva el foco del teclado. Todos los adaptadores aceptan `bytes` de forma
reactiva, borran las cargas fallidas cuando llega un nuevo modelo o una nueva carga, y mantienen su
manejador estable al reemplazar el lienzo.

### JavaScript nativo {#vanilla}

`docx-vanilla-viewer` exporta la sesión compartida como `createViewerBuildingBlocks(options)`
(`canvasProps`, `handle`, `state`, `subscribe`, `update`, `destroy`) y auxiliares de montaje que
devuelven una función de desmontaje: `mountDocumentCanvas(host, blocks)`, `mountToolbar(host, blocks)`,
`mountPageNavigation(host, blocks)` y `mountStatusBar(host, blocks)`.

```ts
import {
	createViewerBuildingBlocks,
	mountDocumentCanvas,
	mountToolbar,
	mountStatusBar,
} from 'docx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ documentModel, canEdit: true });
const stops = [
	mountDocumentCanvas(document.querySelector('#page')!, blocks),
	mountToolbar(document.querySelector('#tools')!, blocks),
	mountStatusBar(document.querySelector('#status')!, blocks),
];
blocks.subscribe((state) => console.log(state.toolbarProps?.bold)); // your own controls
// later
stops.forEach((stop) => stop());
blocks.destroy();
```

## Estilo compartido {#shared-styling}

Los componentes nativos opcionales de la barra de herramientas usan los estilos compartidos de `ooxml-ui`,
incluidos el espaciado, los fondos, los bordes, los estados pulsado y deshabilitado, y los anillos de foco
del teclado. Coloque la barra de herramientas y el lienzo debajo de un mismo ancestro `data-office-theme`, y
personalice allí las variables `--office-*`. Los cambios se aplican de inmediato en todos los frameworks.
Las sustituciones existentes `--dve-*` en un ancestro siguen siendo compatibles. Los controles creados por el
anfitrión conservan su propio estilo, salvo que el anfitrión adopte el auxiliar compartido
`prepareOfficeUiBlock(element, 'toolbar', 'docx')` de `ooxml-ui/theme`.
