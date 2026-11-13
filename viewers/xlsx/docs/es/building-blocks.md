# Bloques de construcción

Cada paquete de framework (`@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer`, `xlsx-angular-viewer`,
`xlsx-svelte-viewer`, `xlsx-solid-viewer` y `xlsx-vanilla-viewer`) exporta un hook o una fábrica
de bloques de construcción (`useViewerBuildingBlocks`, `createViewerBuildingBlocks`), un `SpreadsheetCanvas`
y los controles opcionales `Toolbar`, `FormulaBar` y `SheetTabs`. Todos usan la sesión de libro, la cuadrícula, las fórmulas, la validación, la protección y el historial que ya existen; no hay un segundo almacén de estado.
La forma sigue al `useViewerBuildingBlocks` del visor de PowerPoint.

```tsx
import {
	SpreadsheetCanvas,
	FormulaBar,
	SheetTabs,
	useViewerBuildingBlocks,
} from '@christophervr/xlsx-react-viewer';

function EmbeddedSheet({ workbook, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		workbook,
		canEdit,
		handle,
	});
	return (
		<>
			{toolbarProps && <FormulaBar {...toolbarProps} />}
			<SpreadsheetCanvas {...canvasProps} />
			{toolbarProps && <SheetTabs {...toolbarProps} />}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

El hook también devuelve `mode` (`'view'` o `'edit'`) y los valores reactivos `dirty`, `selection` (texto A1),
`collaboration` (quién está en la sala y el estado de la conexión) y `editError` (vea
[Fallos de validación](#validation-failures)). Renderice el lienzo mientras carga; `toolbarProps`
es null hasta que se monta. Dé una altura
a su contenedor. El lienzo monta solo la cuadrícula. Archivo, configuración, uso compartido, barra de título, cinta, barra de estado,
barra de fórmulas y pestañas de hoja no aparecen a menos que el host proporcione controles aparte.

`toolbarProps` contiene la selección y la celda activa, el estado de las hojas, el texto de la fórmula, el formato y los
estados `enabled`. Conecte `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange`,
`onFontSizeChange`, `onNumberFormatChange`, `onAlignmentChange`, `onUndo` y `onRedo`
a sus controles personalizados. Los campos de formato son null para valores mixtos. La inspección del formato
está limitada a 4096 celdas: las selecciones más grandes devuelven null y `formattingComplete=false`.

`enabled` tiene un indicador por cada callback: `bold`, `italic`, `underline`, `fontFamily`, `fontSize`,
`alignment`, `numberFormat`, `undo`, `redo` y los flujos de trabajo del host `openShareDialog`, `settings`,
`formatCells`, `findReplace` e `insertFunction`. `canUndo` y `canRedo` repiten los indicadores de deshacer y
rehacer. `options` enumera las opciones que puede ofrecer un control personalizado de fuente, tamaño, alineación o formato de número
(`fontFamilies`, `fontSizes`, `alignments`, `numberFormats`), y `labels` contiene los
nombres localizados de los controles.

`onSelect`, `onSheetChange`, `onFormulaChange` y `commitEdit` devuelven false cuando la validación de la celda bloquea la acción. Los callbacks de formato devuelven una `Promise<boolean>`. Confirman una edición de celda activa antes de actuar, y conservan su selección cuando los controles personalizados toman el foco. El texto de la fórmula se confirma con Enter en la `FormulaBar` opcional; una entrada rechazada activa aria-invalid.
El handle existente conserva el guardado, la serialización y el seguimiento de cambios. Un guardado se rechaza si la entrada
no puede confirmarse. Los permisos y la protección de la hoja siguen siendo la autoridad final.

Para cada framework, `mountEditorSurface(host, options)` expone la misma superficie, el handle
existente, `getToolbarProps()`, `subscribe(listener)`, `update` y `destroy`. Los hosts gestionan los
flujos de archivo y los diálogos opcionales; no se requiere ninguna barra de herramientas oculta de la biblioteca.

`formula`, `formulaCaret`, `editing` y `editMode` leen el puente de edición en vivo de la cuadrícula.
`onFormulaInput(text, caret?)` actualiza ese mismo borrador; `onFormulaChange(text)` lo confirma.
Un borrador rechazado puede corregirse con `onFormulaInput` sin confirmar antes el valor no válido.
`onCancelEdit()` lo cancela. No hay un almacén de estado de fórmula aparte.

La API de composición requiere el runtime de `ooxml-ui` de la misma versión del workspace; los paquetes de framework dependen de él y se publican con él.
Los imports de visor completo existentes siguen siendo compatibles con el runtime publicado anterior. Un runtime más antiguo rechaza una
solicitud de composición de forma explícita en lugar de montar la interfaz de la aplicación por accidente.

## Callbacks del host para los diálogos {#host-callbacks-for-dialogs}

La aplicación abre los diálogos de uso compartido, configuración, Formato de celdas, Buscar y reemplazar e Insertar
función. Una superficie nunca los monta. Pase un callback en su lugar, como hacen los bloques de construcción de
PowerPoint (`onOpenSettings`, `onOpenShareDialog`), y abra su propia interfaz:

```ts
useViewerBuildingBlocks({
	workbook,
	onOpenShareDialog: ({ activeCell }) => openMyShareDialog(activeCell),
	onOpenFormatCells: ({ selection }) => openMyFormatPanel(selection),
});
```

Los callbacks disponibles son `onOpenShareDialog`, `onOpenSettings`, `onOpenFormatCells`,
`onOpenFindReplace` y `onOpenInsertFunction`. Cada uno recibe `{ sheet, sheetName, selection,
activeCell }`. `toolbarProps` expone los `onOpenShareDialog`, `onOpenSettings`,
`onOpenFormatCells`, `onOpenFindReplace` y `onOpenInsertFunction` correspondientes (confirman primero una
edición en curso y devuelven una `Promise<boolean>`), y `enabled.openShareDialog` y similares son true solo cuando
se proporcionó el callback (los de formato e inserción también necesitan permiso de edición). `handle.share()`
llama a `onOpenShareDialog`; sin él no hace nada, y `enabled.openShareDialog` es false. Un
flujo sin callback es inalcanzable, nunca se monta en silencio. Los hosts que quieran los
diálogos incluidos renderizan en su lugar el `SpreadsheetEditor` completo.

## Menús contextuales y operaciones de flujo {#context-menus-and-workflow-operations}

Los tipos, las firmas y las condiciones de rechazo están en la [referencia de la API del elemento](/es/api#host-menus-and-workflows).

`onOpenFormatCells`, `onOpenFindReplace` y `onOpenInsertFunction` reciben, además del contexto
anterior, las operaciones para completar el flujo. Las mismas operaciones están en
`toolbarProps.workflows`, de modo que un panel siempre abierto no necesita callback. Cada
operación vuelve a leer el modo de solo lectura, la protección de la hoja y la edición de celda
pendiente al ejecutarse, pasa por la sesión de edición (deshacer, rehacer, estado de cambios y
`onWorkbookChange` se comportan como en la interfaz incluida) e informa de un rechazo mediante su
valor de retorno.

- `workflows.formatCells`: `getState()` devuelve `{ font, alignment, border, protection }` (cada uno
  `{ value, mixed }`: las propiedades compartidas y los nombres que difieren), `numFmt` y `fill`
  (null si hay valores mixtos) y `complete` (false si la selección es demasiado grande).
  `canApply()` y `apply(patch)` aplican un parche de estilo a la selección como un paso de deshacer.
- `workflows.find`: `findAll(query)`, `findNext(query, { backwards })`, `select(match)`,
  `replace(query, text)`, `replaceAll(query, text)` y `canReplace()`. Buscar funciona en un libro de
  solo lectura; reemplazar se rechaza allí y se omiten las celdas bloqueadas por la protección. Igual que Reemplazar en el propio Excel, `replace` y `replaceAll` no aplican la validación de datos: el reemplazo se escribe aunque la regla de la celda lo rechazaría al escribirlo.
- `workflows.insertFunction`: `categories()`, `search({ text, category })` (con `name`, `category`,
  `syntax` y una `description` localizada), `canInsert()` e `insert(name, args?)`, que inserta
  `NAME(` en el cursor del borrador de fórmula activo (iniciando uno si no hay) y devuelve
  `{ text, caret }` para que su campo lo siga. Un borrador de texto simple se rechaza.

Para dibujar su propio menú contextual, pase `onContextMenu(request)`. Se ejecuta con el clic
derecho en una celda, un encabezado de fila o columna y (si hay barra de pestañas) una pestaña de
hoja. Devuelva `true` para suprimir el menú nativo y muestre `request.items` (`id`, `label`,
`command`, `enabled`, `checked`, `shortcut`, `separatorBefore`) en `request.x`, `request.y`; ejecute
una elección con `request.execute(id)`. Cualquier otro valor muestra el menú nativo. Sin clic
derecho, `toolbarProps.contextMenu('cell' | 'row' | 'column' | 'tab', sheet?)` devuelve el mismo
modelo, así un host con sus propias pestañas obtiene el menú de pestaña. Una superficie registra
pocos comandos, por lo que la mayoría de los elementos aparecen con `enabled: false`; Format Cells
se habilita si `onOpenFormatCells` está instalado y el libro es editable. `enabled.contextMenu`,
`enabled.find`, `enabled.replace`, `enabled.applyFormatCells` y `enabled.insertFunctionCall`
indican lo que funciona ahora. Para conservar el menú nativo y modificarlo, pase en su lugar `onContextMenuItems(request)` (solo se
ejecuta si `onContextMenu` no devolvió `true`) y devuelva `{ hide, add }`. `hide` lista los id de
elementos nativos de `request.items`; `add` lista elementos del host `{ id, label, enabled,
requiresEdit, checked, separatorBefore, position, onSelect(request) }`, situados encima
(`position: 'top'`, por defecto) o debajo (`'bottom'`) de los nativos. `enabled` puede ser una
función. Los elementos nativos conservan su estado deshabilitado y se comprueban de nuevo al
ejecutarse; un elemento del host también (su `enabled` y `requiresEdit` frente a un libro de solo
lectura), así que un menú que quedó abierto no puede ejecutar una elección obsoleta. Complementa
`hiddenActions` y `contextMenu.extraCommands` de `customization`, que se aplican a la aplicación
incluida.

## Fallos de validación {#validation-failures}

Una edición de celda en curso debe confirmarse antes de dar formato, navegar o guardar. Cuando la validación de datos la rechaza, los callbacks devuelven `false`, la edición sigue abierta y el fallo se comunica de tres formas:

- `onValidationError({ message, title?, ref, style? })` en las opciones (el elemento envía un evento
  `validation-error` con el mismo detalle);
- la cadena reactiva `editError` (también `toolbarProps.editError`), que vuelve a ser null cuando el borrador
  se cambia, se corrige o se cancela;
- la `FormulaBar` opcional pone `aria-invalid` y muestra el mensaje en un `span` con `role="alert"`.

### La alerta de validación de datos {#the-data-validation-alert}

Los bloques de construcción nunca dibujan la alerta Reintentar / Cancelar de la cuadrícula. Los tres estilos de alerta de Excel
los muestra el host mediante `onOpenValidationAlert(request)` (también aceptado en `surfaceActions`
y por cada adaptador de framework). La solicitud contiene `{ ref, sheet, sheetName, value, style,
title, message }`, donde `value` es el texto que escribió el usuario, y los métodos que corresponden a los
botones de Excel:

| `style`       | Botones de Excel      | Métodos                                     |
| ------------- | --------------------- | ------------------------------------------- |
| `stop`        | Reintentar / Cancelar | `retry()`, `cancel()`                       |
| `warning`     | Sí / No / Cancelar    | `accept()` (Sí), `retry()` (No), `cancel()` |
| `information` | Aceptar / Cancelar    | `accept()` (Aceptar), `cancel()`            |

`retry()` mantiene el editor abierto con el borrador, `cancel()` restaura el valor anterior y `accept()`
confirma el valor de todos modos, exactamente como los botones incluidos. Resuelva una solicitud una vez; las llamadas
posteriores se ignoran. Hasta que se resuelva, el editor de celdas ignora nuevas confirmaciones, así que resuélvala siempre
(un callback que lanza una excepción cuenta como `retry()`). `onValidationError` y `editError` siguen informando del
fallo.

Sin el callback no se monta nada: todos los estilos actúan como Reintentar, la edición sigue abierta y el
fallo se informa mediante `onValidationError` / `editError`.
`toolbarProps.enabled.validationAlert` es true solo cuando el host proporciona el callback. El
`SpreadsheetEditor` completo y `<xlsx-editor>` mantienen su alerta incluida.

Los controles personalizados también deben mostrar `editError` y permitir al usuario corregir el borrador con
`onFormulaInput` o descartarlo con `onCancelEdit`.

## Toolbar opcional {#optional-toolbar}

La `Toolbar` de la biblioteca muestra cada control que admiten las propiedades, así que puede sustituir a una
construida a mano: deshacer, rehacer, familia de fuente, tamaño de fuente, negrita, cursiva, subrayado, alineación
a la izquierda, al centro y a la derecha, y formato de número. Los botones exponen `aria-pressed` (`mixed` para una selección mixta),
los selectores muestran un valor en blanco para el formato mixto, y todos los controles respetan `enabled`. Los botones no
toman el foco, así que la cuadrícula conserva su selección mientras da formato. La Toolbar no muestra los
flujos de diálogo de arriba; conecte `onOpenShareDialog` y los demás callbacks `onOpen*` a sus propios
controles. `toolbarModel(toolbarProps)` (exportado por `xlsx-vanilla-viewer` y el runtime de Svelte) devuelve
la misma lista de controles para renderizadores personalizados.

## Colaboración {#collaboration}

La coedición en tiempo real funciona en una superficie: pase `collaboration` (un id de sala, URL del servidor, proveedor o
sesión) en las opciones, y el hook devuelve el estado reactivo `collaboration` (`active`,
`roomId`, `status`, `synced`, `people`). `onUndo` y `onRedo` deshacen solo las ediciones de esta ventana mientras
se comparte. Salir o desmontar abandona la sala. La lista de presencia de la barra de título pertenece a la
aplicación, así que el host renderiza `collaboration.people` por su cuenta.

## JavaScript puro {#vanilla}

`xlsx-vanilla-viewer` exporta `createViewerBuildingBlocks(options)`, construido sobre la misma sesión que
los demás frameworks:

```ts
import { createViewerBuildingBlocks } from 'xlsx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ workbook, canEdit: true });
const stops = [
	blocks.mountToolbar(toolbarHost),
	blocks.mountFormulaBar(formulaHost),
	blocks.mountSheetTabs(tabsHost),
	blocks.mountCanvas(canvasHost),
];
blocks.subscribe((state) => render(state.dirty, state.editError, state.toolbarProps));
blocks.update({ workbook, canEdit: false });
blocks.getState().toolbarProps?.onBold();
blocks.destroy();
```

La superficie es **de solo lectura hasta que el host pasa `canEdit: true`**, igual que los bloques de construcción de PowerPoint y Visio (el `SpreadsheetEditor` completo sigue siendo editable por defecto). El permiso se aplica antes de montar el lienzo, así que `mode` empieza como `'view'` y no hay un destello editable.
`readOnly` sigue teniendo prioridad sobre `canEdit: true`.

Cada `mount*` devuelve su función de liberación. `getState()` devuelve la misma instantánea que los hooks
(`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration`, `editError`) y
`handle` es el handle imperativo habitual. Los controles parchean sus elementos en su sitio, así que el foco y
los borradores sobreviven a las actualizaciones.

## Composición nativa de frameworks {#native-framework-composition}

El código de cada framework se distribuye desde su propio paquete. Todos los adaptadores son dueños de un
motor de cuadrícula existente y exponen el mismo `EditorHandle`; sus instantáneas reactivas nativas exponen
`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration` y `editError`. Los controles nativos opcionales reciben una
prop/input `toolbarProps`. Los controles personalizados llaman a las acciones de esa misma instantánea.

### Vue {#vue}

Llame al composable durante `setup`. Pase un getter para seguir los cambios de permiso, documento, bytes y
callbacks. Las refs computadas se desenvuelven en las plantillas; al desmontar se libera el motor.

```vue
<script setup lang="ts">
import {
	useViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-vue-viewer';
const props = defineProps<{ workbook: import('xlsx-vue-viewer').Workbook; canEdit: boolean }>();
const { canvasProps, toolbarProps, loading, error, handle } = useViewerBuildingBlocks(() => ({
	workbook: props.workbook,
	canEdit: props.canEdit,
}));
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<FormulaBar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<SpreadsheetCanvas v-bind="canvasProps" style="height: 600px" />
	<SheetTabs v-if="toolbarProps" :toolbar-props="toolbarProps" />
</template>
```

### Angular {#angular}

Proporcione `ViewerBuildingBlocks` en cada host y luego inyecte esa instancia en el host o en sus
controles personalizados. `options` es una signal escribible; las instantáneas son signals computadas. El
`DestroyRef` del inyector se encarga de la limpieza. El servicio también expone los emisores de eventos `workbookChange`,
`workbookError`, `selectionChange` y `dirtyChange` para las suscripciones del host, junto a los
callbacks de las opciones. Use `takeUntilDestroyed()` al suscribirse en los hosts.

```ts
@Component({
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [SpreadsheetCanvasComponent, ToolbarComponent, FormulaBarComponent, SheetTabsComponent],
	template: `
		@if (blocks.toolbarProps(); as p) {
			<spreadsheet-toolbar [toolbarProps]="p" />
			<spreadsheet-formula-bar [toolbarProps]="p" />
			<button (click)="p.onBold()" [disabled]="!p.canEdit">Custom bold</button>
			<spreadsheet-sheet-tabs [toolbarProps]="p" />
		}
		<spreadsheet-canvas [mount]="blocks.canvasProps.mount" style="display:block;height:600px" />
	`,
})
export class WorkbookHost {
	readonly workbook = input.required<Workbook>();
	readonly canEdit = input(true);
	readonly blocks = inject(ViewerBuildingBlocks);
	constructor() {
		effect(() => this.blocks.options.set({ workbook: this.workbook(), canEdit: this.canEdit() }));
	}
}
```

Importe los componentes, el servicio y `Workbook` desde `xlsx-angular-viewer`, y `Component`, `input`,
`effect` e `inject` de `@angular/core`. `SpreadsheetBuildingBlocks`
es el nombre específico del producto para el mismo servicio. Su `handle` conserva load, save y la selección.

### Solid {#solid}

Cree el componente dentro de un owner de Solid. Los getters de entrada y los accesores de salida participan en el
grafo reactivo de Solid; la limpieza sigue a ese owner. Los componentes aceptan getters reactivos como props.

```tsx
import { Show } from 'solid-js';
import {
	createViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-solid-viewer';
function WorkbookHost(props: { workbook: Workbook; canEdit: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		workbook: props.workbook,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<Show when={blocks.toolbarProps()}>
				{(p) => (
					<>
						<Toolbar toolbarProps={p()} />
						<FormulaBar toolbarProps={p()} />
						<button onClick={() => p().onBold()}>Custom bold</button>
						<SheetTabs toolbarProps={p()} />
					</>
				)}
			</Show>
			<SpreadsheetCanvas {...blocks.canvasProps} />
		</>
	);
}
```

### Svelte 5 {#svelte-5}

Cree el store durante la inicialización del componente. `$blocks` se suscribe de forma nativa; use un
`$effect` para reenviar las opciones reactivas. La limpieza del motor sigue a la destrucción del componente.
Los fuentes de los componentes y sus declaraciones se publican como subrutas explícitas.

```svelte
<script lang="ts">
  import { createViewerBuildingBlocks, type Workbook } from 'xlsx-svelte-viewer/runtime';
  import SpreadsheetCanvas from 'xlsx-svelte-viewer/SpreadsheetCanvas.svelte';
  import Toolbar from 'xlsx-svelte-viewer/Toolbar.svelte';
  import FormulaBar from 'xlsx-svelte-viewer/FormulaBar.svelte';
  import SheetTabs from 'xlsx-svelte-viewer/SheetTabs.svelte';
  let { workbook, canEdit = false }: {workbook: Workbook; canEdit?: boolean} = $props();
  const blocks = createViewerBuildingBlocks();
  $effect(() => blocks.update({workbook, canEdit}));
</script>
{#if $blocks.toolbarProps}
  <Toolbar toolbarProps={$blocks.toolbarProps} />
  <FormulaBar toolbarProps={$blocks.toolbarProps} />
  <button onclick={() => $blocks.toolbarProps?.onBold()}>Custom bold</button>
  <SheetTabs toolbarProps={$blocks.toolbarProps} />
{/if}
<SpreadsheetCanvas {...blocks.canvasProps} />
```

`FormulaBar` comparte el borrador en vivo de la cuadrícula. Enter confirma y pone `aria-invalid` en una
entrada rechazada. Reintentar permite corregir; Escape cancela. El formato en conflicto, la selección, la navegación
entre hojas y el guardado siguen bloqueados mientras la validación rechaza un borrador. Readonly
bloquea las modificaciones y conserva la selección, la navegación entre hojas y la serialización.

## Apariencia compartida {#shared-appearance}

Coloque el lienzo y los controles nativos opcionales bajo un ancestro `data-office-theme="light|dark|auto"`. Las cinco
integraciones nativas de framework instalan los mismos estilos de control compartidos y
leen los tokens `--office-*` heredados. Personalice `--office-surface`, `--office-background`,
`--office-foreground`, `--office-border`, `--office-accent`, `--office-ring`, `--office-font`
y `--office-radius` en ese ámbito. Los cambios en tiempo de ejecución conservan los borradores y el foco.

Las opciones existentes `theme` y `themeColors` del lienzo siguen aplicándose localmente; no
reescriben los ancestros del host ni sus controles hermanos. Las sustituciones heredadas `--xve-*` en un
ancestro común siguen llegando tanto al lienzo como a los controles opcionales. Vea [Temas](/es/theming) para el
ámbito y ejemplos de compatibilidad.
