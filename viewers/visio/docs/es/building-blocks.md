# Bloques de construcción

`visio-react-viewer` exporta `useViewerBuildingBlocks`, `DiagramCanvas` y, de forma opcional, `Toolbar`, `PageNavigation`, `Inspector` y `ShapesPane`. Estos reutilizan ViewerController, el renderizador SVG existente, el inspector de selección y la ventana de plantillas. Todos los demás frameworks tienen las mismas piezas (véase más abajo), incluido JavaScript vanilla.

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error, markClean } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
		onOpenSettings: () => openMySettingsDialog(),
	});
	return (
		<>
			<PageNavigation {...toolbarProps} />
			<DiagramCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

`content` acepta el mismo origen de bytes o Blob que el visor. `handle` recibe el ViewerController existente; use `load`, `applyEdits`, `exportVsdx` y su `state.edit.dirty` reactivo para la persistencia en el anfitrión. `onDocumentChange`, `onDirtyChange` y `onDocumentError` son opcionales. Dé al contenedor del lienzo una altura y manténgalo montado mientras se carga.

El objeto plano `toolbarProps` es todo el contrato que necesita una interfaz personalizada; cada campo procede del controlador existente, de modo que nada se almacena dos veces:

| Campo                                                                                                                               | Significado                                                                                                                                                                                                                                                                                                             |
| ----------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `mode`, `canEdit`                                                                                                                   | `mode` es `'edit'` solo cuando el anfitrión concedió `canEdit` **y** el dibujo abierto es un `.vsdx` con origen (un `.vsd` heredado o un dibujo solo de modelo indica `'view'`). `canEdit` es verdadero solo cuando el anfitrión permite la edición, el dibujo tiene origen y no hay ninguna carga ni edición en curso. |
| `pageIndex`, `pages`, `zoom`, `selection`                                                                                           | Página actual, lista de páginas, factor de zoom y formas seleccionadas.                                                                                                                                                                                                                                                 |
| `canUndo`, `canRedo`, `canDuplicate`, `canGroup`, `canUngroup`, `canDelete`, `canCut`, `canCopy`, `canPaste`                        | Indicadores de habilitación. Las acciones de edición son falsas mientras el documento es de solo lectura; `canCopy` requiere un dibujo con origen y una selección en la página actual, pero no el permiso de edición.                                                                                                   |
| `canZoomIn`, `canZoomOut`                                                                                                           | Límites de zoom (del 10 % al 800 %).                                                                                                                                                                                                                                                                                    |
| `search`                                                                                                                            | `{ query, resultCount, activeIndex, truncated }` procedente de la búsqueda del controlador. `activeIndex` es -1 hasta que el anfitrión navega.                                                                                                                                                                          |
| `tool`, `tools`, `canDraw`                                                                                                          | Herramienta activa del lienzo (`'pointer'` o una de `rectangle`, `ellipse`, `line`, `freeform`, `arc`, `pencil`, `text`, `connector`), la lista con etiquetas localizadas e indicadores de activo, y si se permite dibujar.                                                                                             |
| `labels`                                                                                                                            | Etiquetas tipadas y localizadas (`DiagramToolbarLabels`), una entrada por cada control, en `en`, `de`, `es`, `fr` y `zh-CN`.                                                                                                                                                                                            |
| `onPageChange`, `onZoomChange`, `onZoomIn`, `onZoomOut`, `onFit`                                                                    | Acciones de visualización, siempre disponibles.                                                                                                                                                                                                                                                                         |
| `onUndo`, `onRedo`, `onDuplicate`, `onGroup`, `onUngroup`, `onDelete`, `onCut`, `onCopy`, `onPaste`, `onToolChange`, `onTextChange` | Acciones de edición. Mantienen las comprobaciones de permiso, de carga y de formato de origen del controlador. Las acciones del portapapeles usan el lienzo montado y el portapapeles del navegador, por lo que necesitan un gesto del usuario.                                                                         |
| `onSearchChange`, `onNextSearchResult`, `onPreviousSearchResult`                                                                    | Búsqueda, siempre disponible.                                                                                                                                                                                                                                                                                           |

`Inspector` consume `inspectorProps`, y `ShapesPane` consume `shapesProps`. El propio lienzo no monta barra de título, cinta, controles de página, panel de formas o plantillas, inspector, interfaz de archivos, recuperación, cuadros de configuración ni de uso compartido. Los límites existentes de edición y protección de VSDX siguen aplicándose.

## Permiso y modo {#permission-and-mode}

`canEdit` sigue a los bloques de construcción de PowerPoint: el dibujo es **de solo lectura hasta que el anfitrión pase `canEdit: true`**, y el permiso se aplica al crear el controlador, antes de montar el lienzo, de modo que nada sea editable brevemente (el visor completo conserva su valor predeterminado editable). Revocarlo cancela las ediciones pendientes, quita la herramienta de dibujo y elimina los controles de cambio de tamaño, rotación y puntos finales. La navegación de páginas, la selección, la búsqueda, el desplazamiento y el paneo, el zoom, el ajuste, copiar y la serialización siguen disponibles. El soporte de edición del formato de origen es independiente del permiso del anfitrión, por eso `mode` también lo comprueba.

## Qué puede hacer el lienzo {#what-the-canvas-can-do}

El lienzo independiente usa los mismos gestos del visor completo (una conexión compartida, no una copia): mover, texto en línea (doble clic, F2 o escribir sobre una forma seleccionada), controles de cambio de tamaño y rotación, puntos finales de línea, las herramientas de dibujo, la herramienta de texto y la herramienta de conector elegidas con `onToolChange`. La herramienta de texto arrastra un cuadro y abre un borrador de texto sin formato (Ctrl+Intro añade el cuadro de texto, Escape cancela); la herramienta de conector arrastra de forma a forma o de punto de conexión a punto de conexión y une los extremos exactamente como el visor completo, con una ruta en ángulo recto. Suprimir, Ctrl+Z, Ctrl+Y y Escape (salir de la herramienta de dibujo) funcionan en el lienzo enfocado, y se respetan los eventos de copiar, cortar y pegar del navegador. Todas estas acciones están sujetas a `canEdit`, exactamente como en el visor completo. Los rechazos de edición se anuncian en una región activa de tipo cortés dentro del lienzo.

## Interfaz opcional de la biblioteca {#optional-library-ui}

`Toolbar` muestra todos los controles que admite `toolbarProps`: historial, zoom y ajuste, duplicar, agrupar, desagrupar y eliminar, cortar, copiar y pegar, las herramientas de dibujo, la búsqueda y los botones de diálogo del anfitrión que aparecen más abajo, con su estado de habilitación. Es opcional: pase `include` (`'file' | 'history' | 'zoom' | 'selection' | 'clipboard' | 'tools' | 'search'`) para mostrar solo algunos grupos, u omítalo y conecte `toolbarProps` a sus propios controles. `PageNavigation` es el selector de páginas.

`ShapesPane` es la ventana de Formas (plantillas): Buscar formas, la plantilla de documento del dibujo y las plantillas integradas. Al hacer clic en un maestro, o al soltarlo sobre el lienzo montado, se sigue el mismo flujo de creación que en el visor completo, por lo que obedece el mismo permiso. Monte el lienzo sobre el mismo controlador; sin lienzo, los maestros se añaden en el centro de la página.

## Diálogos gestionados por el anfitrión {#host-owned-dialogs}

Como en los callbacks `onOpen...` de PowerPoint, los diálogos pertenecen al anfitrión y la biblioteca nunca monta uno en modo superficie. Pase cualquiera de `onOpenFile`, `onOpenPageSetup`, `onOpenSettings` y `onOpenShareDialog` al hook o a la factoría. Se devuelven en `toolbarProps` con `canOpenFile`, `canOpenPageSetup`, `canOpenSettings` y `canOpenShareDialog`; un callback que omita no aparece y su indicador es `false` (no habilitado), y el `Toolbar` de la biblioteca muestra un botón solo para los que haya suministrado.

## Persistencia y seguimiento de cambios pendientes {#persistence-and-dirty-tracking}

`controller.exportVsdx()` devuelve los bytes y nunca cambia `dirty`. Después de que el anfitrión los haya guardado, llama a `markClean()` (en el resultado del hook, en el controlador o en el handle). Eso borra `state.edit.dirty`, dispara `onDirtyChange(false)` y un evento `document-change` de tipo `save`. Las ediciones y los pasos de deshacer posteriores se miden a partir de ese punto guardado.

## Guardar mientras se edita el texto de una forma {#saving-while-editing-text}

Si el usuario pulsa un Guardar gestionado por el anfitrión mientras escribe dentro de una forma, el borrador aún no está en el documento. Llame primero a `await commitTextEdit()` (en el resultado del hook, la sesión, los bloques vanilla, el servicio de Angular o el controlador/handle) y después a `exportVsdx()`. Solo se resuelve cuando termina la edición del borrador, también si un blur o Esc ya la había iniciado, y devuelve un `TextEditOutcome`:

- `committed`: el borrador se aplicó como una edición que se puede deshacer; `text` es lo aplicado.
- `unchanged` (`reason`: `no-draft` o `same-text`): no hay nada que aplicar.
- `refused` (`reason`: `read-only` o `rejected`): el editor lo rechazó; `message` explica el motivo y `text` conserva el borrador para recuperarlo.
- `cancelled` (`reason`: `stale-document`, `superseded` o `destroyed`): el documento del borrador ya no existe o la edición fue reemplazada; `text` es el borrador.

Cada resultado incluye `sourceGeneration` y `documentGeneration`, y `pageId` y `shapeId` cuando había un borrador. Un borrador pertenece al documento en el que se escribió: cargar otro archivo o desmontar el lienzo lo descarta, por lo que nunca se aplica a otro documento. Si una carga descartó un borrador sin guardar, la siguiente llamada a `commitTextEdit()` lo devuelve una sola vez como `cancelled` / `stale-document` con su `text`, `pageId`, `shapeId` y la `sourceGeneration` anterior, para que un Guardar que coincida con un cambio de documento pueda recuperar lo escrito. La validación y deshacer/rehacer siguen la vía de edición normal.

## Paquetes {#packages}

Cada paquete de framework reexporta `ooxml-ui/visio`, de modo que `mountDiagramCanvas(host, controller)`, `mountDiagramInspector(host, controller)`, `mountDiagramShapes(host, controller, { locale? })`, `diagramToolbarProps(controller, fit, locale?, hostCallbacks?)` y `fitDiagramCanvas` están disponibles desde cualquiera de ellos. Suscríbase a través del controlador existente y llame a los disposers devueltos al desmontar. No se introduce un segundo almacén de estado.

Los paquetes del visor y `ooxml-ui` se publican juntos, y cada paquete del visor depende de la versión de `ooxml-ui` con la que se publicó, por lo que la API de composición siempre está presente. Un anfitrión que mezcle versiones a mano recibe un error explícito de un `ooxml-ui` más antiguo en lugar de montar la interfaz de la aplicación por accidente.

## Colaboración {#collaboration}

Los bloques de construcción no admiten la opción `collaboration`: compartir un dibujo en una sala es una propiedad del `<visio-viewer>` completo (consulte [colaboración](/es/collaboration)). Una interfaz personalizada puede ofrecer su propia entrada de uso compartido mediante `onOpenShareDialog`.

## Composición nativa en cada framework {#native-framework-composition}

La composición específica de cada framework reside en su paquete. Todos los adaptadores poseen un `ViewerController`, exponen su estado mediante la reactividad del framework, cancelan las cargas reemplazadas y liberan las suscripciones y los manejadores del lienzo junto con su propietario. Un lienzo que se haya quitado de forma independiente puede volver a montarse sobre ese mismo controlador sin recargar el origen.

### Vue {#vue}

```vue
<script setup lang="ts">
import {
	DiagramCanvas,
	Toolbar,
	PageNavigation,
	Inspector,
	useViewerBuildingBlocks,
} from 'visio-vue-viewer';
const props = defineProps<{ content?: Uint8Array; canEdit?: boolean }>();
const blocks = useViewerBuildingBlocks(() => ({ content: props.content, canEdit: props.canEdit }));
const { toolbarProps, loading, error, mode } = blocks;
defineExpose({ controller: blocks.controller });
</script>
<DiagramCanvas v-bind="blocks.canvasProps" />
<Toolbar :toolbar-props="toolbarProps" />
<PageNavigation :toolbar-props="toolbarProps" />
<Inspector v-bind="blocks.inspectorProps" />
```

`toolbarProps`, `state`, `loading`, `error` y `mode` son refs; las plantillas de Vue los desenvuelven. Pase un getter de opciones o una ref reactiva para seguir los cambios de origen, de permiso y de configuración regional.

### Solid {#solid}

```tsx
import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-solid-viewer';
function EmbeddedDiagram(props: { content?: Uint8Array; canEdit?: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		content: props.content,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<DiagramCanvas {...blocks.canvasProps} />
			<Toolbar {...blocks.toolbarProps()} />
		</>
	);
}
```

`toolbarProps`, `state`, `loading`, `error` y `mode` son accesores. El controlador está disponible como `blocks.controller`; la liberación sigue al propietario de Solid. `PageNavigation` acepta el mismo descriptor plano de la barra de herramientas, `Inspector` acepta `inspectorProps` y `ShapesPane` acepta `shapesProps`.

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
 import { writable } from 'svelte/store';
 import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-svelte-viewer';
 const options = writable({ canEdit: true });
 const blocks = createViewerBuildingBlocks(options);
 const toolbarProps = blocks.toolbarProps;
 // options.set({ canEdit: false }) updates permissions through the store subscription.
</script>
<DiagramCanvas {...blocks.canvasProps} />
<Toolbar toolbarProps={$toolbarProps} />
```

El argumento de opciones acepta un store de lectura o un objeto inicial simple; los anfitriones con objetos simples pueden llamar a `blocks.update(options)` desde un efecto de runes. `toolbarProps`, `state`, `loading`, `error` y `mode` son stores de lectura. `PageNavigation` recibe `toolbarProps`, `Inspector` recibe `inspectorProps` y `ShapesPane` recibe `shapesProps`. El controlador está disponible como `blocks.controller`.

### JavaScript vanilla {#vanilla-javascript}

```ts
import { createViewerBuildingBlocks } from 'visio-vanilla-viewer';
const blocks = createViewerBuildingBlocks({ canEdit: false, onOpenSettings: openSettings });
blocks.mountCanvas(document.querySelector('#canvas')!);
blocks.mountToolbar(document.querySelector('#toolbar')!, ['history', 'zoom']);
blocks.mountPageNavigation(document.querySelector('select')!);
blocks.mountInspector(document.querySelector('#inspector')!);
blocks.mountShapes(document.querySelector('#shapes')!);
blocks.update({ content: bytes, canEdit: true });
blocks.subscribe(() => render(blocks.toolbarProps, blocks.mode));
// later: blocks.markClean(); blocks.destroy();
```

La factoría vanilla ejecuta la misma sesión que los demás adaptadores. `toolbarProps`, `state`, `loading`, `error` y `mode` son getters; `subscribe(listener)` informa de cada cambio de estado o de opciones, y cada helper `mount...` devuelve su disposer; `destroy()` libera lo que siga montado. La barra de herramientas montada actualiza sus elementos en su sitio, de modo que el foco se conserva al refrescar.

### Angular {#angular}

```ts
import { Component, Input, inject, type OnChanges } from '@angular/core';
import {
	ViewerBuildingBlocks,
	DiagramCanvasComponent,
	ToolbarComponent,
} from 'visio-angular-viewer';
@Component({
	selector: 'embedded-diagram',
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [DiagramCanvasComponent, ToolbarComponent],
	template: '<visio-diagram-canvas /><visio-diagram-toolbar />',
})
export class EmbeddedDiagram implements OnChanges {
	@Input() content?: Uint8Array;
	@Input() canEdit = false;
	readonly blocks = inject(ViewerBuildingBlocks);
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit });
	}
}
```

El servicio con ámbito expone las señales de solo lectura `state`, `toolbarProps`, `loading`, `error` y `mode`, el controlador existente y los emisores de eventos `documentChange`, `dirtyChange` y `documentError`. `DiagramCanvasComponent`, `ToolbarComponent`, `PageNavigationComponent`, `InspectorComponent` y `ShapesPaneComponent` inyectan ese propietario. Coloque el proveedor en cada host de editor, en lugar de a nivel de toda la aplicación, para aislar varios documentos. Como alternativa, envuelva los controles en `<visio-building-blocks [content]="bytes" [canEdit]="allowed" (dirtyChange)="saveNeeded = $event">` usando `ViewerBuildingBlocksComponent`, que proporciona entradas, salidas y un proveedor con ámbito.

Todos los controles opcionales renderizan plantillas nativas del framework. Los anfitriones pueden omitirlos y llamar a las acciones planas de la barra de herramientas desde su propia interfaz. Los callbacks de la barra de herramientas conservan las comprobaciones de permiso, de carga y de formato de origen del controlador. Las pruebas de regresión en navegador cubren las seis composiciones nativas, incluidas la edición, deshacer y rehacer, la revocación de permisos, la serialización y recarga, la recuperación de errores y la liberación.

Para un anfitrión que ya usa señales de Angular, `createViewerBuildingBlocks(() => ({
content: bytes(), canEdit: allowed() }))` inyecta el servicio `ViewerBuildingBlocks` con ámbito y conecta el getter de opciones con un efecto de Angular. Llámelo en el contexto de inyección del componente. Los emisores de eventos del servicio están activos incluso antes de la primera llamada explícita a `update`.

Al cambiar la prop `mount` de un lienzo independiente se libera su propietario anterior antes de adjuntar el siguiente al mismo host. Cambiar el controlador de un inspector sustituye de forma similar su suscripción. Esto permite alternar entre sesiones de edición independientes sin manejadores de entrada ni suscripciones de estado huérfanos.

## Tema Office compartido {#shared-office-theme}

Cada barra de herramientas opcional nativa, selector de páginas e inspector sigue las mismas variables CSS `--office-*` que el lienzo de dibujo. Los bloques instalan su hoja de estilos compartida automáticamente; no existe ningún objeto de tema específico de un framework ni un segundo almacén de estado del documento. Defina temas con ámbito en su propio elemento host:

```html
<section
	data-office-theme="dark"
	style="--office-accent: #a78bfa; --office-font: Inter, sans-serif"
>
	<!-- DiagramCanvas, Toolbar, PageNavigation and Inspector -->
</section>
```

Use `data-office-theme="light"`, `"dark"` o `"auto"` en la raíz del documento o en una sección contenedora. `auto` sigue la preferencia del sistema, incluidos los cambios en tiempo de ejecución. Los hosts claros y oscuros con ámbito pueden coexistir en una misma página. Sin el atributo con ámbito, los bloques heredan el tema de la raíz y las variables del host a través del límite shadow del lienzo. Cambiar el tema no recarga el lienzo, no restablece la selección ni modifica los colores del documento ni sus bytes serializados.

Para una paleta específica del anfitrión, sobrescriba `--office-background`, `--office-surface`, `--office-foreground`, `--office-muted-foreground`, `--office-border`, `--office-accent`, `--office-ring` y `--office-font` en la sección contenedora. Las sobrescrituras heredadas de Visio `--vv-surface`, `--vv-ink`, `--vv-border`, `--vv-accent`, `--vv-focus` y `--vv-font` siguen admitidas, con prioridad sobre sus equivalentes canónicos. El lienzo mantiene además `--vv-background`, `--vv-secondary` y las sobrescrituras existentes de la barra de estado. La página del diagrama conserva su propia apariencia de documento mientras los controles que la rodean siguen el tema de la interfaz.
