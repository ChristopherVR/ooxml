# Adaptadores de framework

Todos los adaptadores montan `<docx-editor>` mediante la misma función `mountEditor`. Una asignación de propiedad es una sustitución externa del documento, no un evento de edición. `document-change` transporta directamente un `DocumentModel`; `document-error` transporta un `Error`. Ambos eventos burbujean a través del límite del shadow DOM.

React:

```tsx
import { WordEditor } from 'docx-react-viewer';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` expone `element`, `load(bytes)`, `save()`, `download()`, `markClean()`, `openCloudFile()` y `dirty`.

Vue:

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular:

```ts
import { WordEditorComponent } from 'docx-angular-viewer';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5:

```svelte
<script>
  import WordEditor from 'docx-svelte-viewer';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

SolidJS:

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from 'docx-solid-viewer';

function DocumentEditor() {
	const [model, setModel] = createSignal(initialDocument);
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

Solid utiliza las mismas opciones `documentModel`, `readOnly`, `locale`, `onDocumentChange` y `onDocumentError` que el adaptador compartido. Use `editorRef` para obtener los métodos `element`, `load` y `save` del editor montado. El propietario de Solid limpia el editor y sus escuchadores cuando el componente se desmonta.

Elemento personalizado:

```ts
import { registerDocxEditor } from 'docx-vanilla-viewer';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

El adaptador de montaje vanilla se exporta desde `docx-vanilla-viewer` y devuelve las operaciones update, destroy, load, save y element. Evita reinicios por realimentación cuando un componente padre devuelve el modelo emitido. Proporcione un objeto nuevo para los cambios de modelo externos; modificar un modelo in situ no es un mecanismo de reactividad admitido.

Los paquetes de framework usan sus propias APIs nativas de ciclo de vida y de referencias. Ninguna plantilla de framework contiene una barra de herramientas, un renderizador de páginas ni una implementación de comandos de documento. Las importaciones de SSR no registran elementos; el montaje es una operación del cliente. El contrato actual del navegador valida las versiones instaladas, no todas las versiones históricas de las dependencias peer.

Defina la propiedad opcional `locale` en cualquier editor de framework para localizar la cinta de opciones y los controles de búsqueda (`en`, `fr`, `de`, `es` o `zh-CN`; por ejemplo, `locale="de"`). Esto afecta solo al texto de la interfaz; no traduce el contenido del documento.

## Marco de la ventana y comandos de archivo {#window-chrome-and-file-commands}

El editor dibuja el marco estilo Word dentro de su shadow root: una barra de título (Guardar, Deshacer y Rehacer de acceso rápido, el nombre del archivo y el estado de guardado, y una búsqueda de comandos "Dígame qué desea hacer"), el selector de modo Edición/Lectura, **Comentarios** y **Compartir** en el extremo derecho de la fila de pestañas de la cinta (como en Word 365), una pestaña **Archivo** con su vista backstage (Información con notas de compatibilidad, Nuevo, Abrir, Guardar, Guardar una copia como DOCX, Imprimir), y una barra de estado (página, palabras, notas de compatibilidad, diseño Web/Impresión y zoom).

Defina `element.fileName` para dar nombre al documento abierto. Los comandos de archivo se anuncian primero como un evento `file-command` cancelable cuyo `detail.command` es `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` o `share`. Llame a `preventDefault()` para gestionarlo en su aplicación (por ejemplo, para guardar en su propio almacenamiento); en caso contrario, el editor usa su comportamiento predeterminado solo para navegador: un selector de archivos, una descarga o el diálogo de impresión. `share` (el botón Compartir) abre el diálogo Compartir que usan todos los editores de Office: inicie una sesión (nombre de la sesión, su nombre y un servidor opcional) o únase a una desde un enlace de invitación o un nombre de sesión, después copie el enlace de invitación, vea quién está en la sesión y deje de compartir, lo que establece la propiedad `collaboration`; cancélelo para compartir a su manera. Al elegir **Lectura** se emite `readonly-change`.

```ts
editor.fileName = 'Quarterly report.docx';
editor.addEventListener('file-command', (event) => {
	const { command } = (event as CustomEvent<{ command: string }>).detail;
	if (command === 'save') {
		event.preventDefault();
		void editor.save().then(uploadToMyStorage);
	}
});
```

## Guardado, estado de cambios sin guardar y opciones de la interfaz {#saving-dirty-state-and-ui-options}

`save()` se resuelve en un `Blob` del documento serializado (`saveBytes()` devuelve el `Uint8Array` sin procesar), `download(fileName?)` guarda e inicia una descarga del navegador, y `dirty` / `dirty-change` informan de las ediciones sin guardar. `dirty` pasa a ser verdadero en `document-change` y a falso tras Archivo > Guardar, `download()`, cargar un documento o `markClean()`; `save()` por sí solo no lo borra, porque solo el anfitrión sabe si el Blob se ha persistido. Cada adaptador expone `save`, `download`, `markClean` y `dirty` en su manejador, así como `onDirtyChange` / `onPageChange` (o el equivalente de evento del framework).

| Propiedad (atributo)                 | Predeterminado | Efecto                                                                                                                                                                                                                                                                                                                                                            |
| ------------------------------------ | -------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `showThumbnails` (`show-thumbnails`) | `false`        | Panel izquierdo con miniaturas de página. Las miniaturas requieren el diseño de Impresión; en la vista Borrador, el panel lo indica en lugar de listar páginas.                                                                                                                                                                                                   |
| `showToolbar` (`show-toolbar`)       | `true`         | `show-toolbar="false"` oculta la cinta; la barra de título y la barra de estado permanecen.                                                                                                                                                                                                                                                                       |
| `hiddenActions`                      | `[]`           | Controles de la cinta que se ocultan por su id estable en kebab-case (`RIBBON_ACTION_IDS`, por ejemplo `'bold'`, `'insert-table'`, `'track-changes'`, `'paste'`, `'word-count'`); los grupos y pestañas vacíos también se ocultan. Un botón de color y su flecha desplegable se ocultan juntos. Los ids desconocidos se ignoran. Los ids no cambian con `locale`. |
| `ribbonAddIns`                       | ninguno        | Pestañas que el anfitrión añade después de las integradas; consulte [Pestañas de complementos de la cinta](#ribbon-add-in-tabs).                                                                                                                                                                                                                                  |

**Obsoleto:** antes de que existieran los ids, `hiddenActions` recibía la etiqueta en inglés del control (`'Bold'`, `'Insert table'`). Esas cadenas `LegacyRibbonLabel` todavía se aceptan durante una versión más: se asignan a ids y el elemento emite un `document-warning` por cada asignación. Leer `hiddenActions` siempre devuelve ids. Los eventos `ribbon-action` nunca transportaron etiquetas (su detalle se identifica por el tipo de comando), así que no cambian.

`page-change` (`{ page, pageCount }`) se emite solo desde el diseño de Impresión. La paginación procede del propio motor de diseño de este editor y es una aproximación, no la paginación de Word; la barra de estado y los textos emergentes de las miniaturas lo indican así.

## Pestañas de complementos de la cinta {#ribbon-add-in-tabs}

Un anfitrión puede añadir sus propias pestañas de la cinta después de las de Word, como lo hace un complemento de Office.
`ribbonAddIns` acepta datos simples y es una propiedad de todos los adaptadores; un comando elegido ejecuta su callback `run` y se notifica mediante `onRibbonAddIn` (en Vue `office-ribbon-add-in`, en Angular `ribbonAddIn`, en Svelte `onribbonaddin`) con `{ tab, command }`.

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [
					{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() },
					{ id: 'options', label: 'Options', size: 'small' },
					{ id: 'send', label: 'Send', items: [{ id: 'send-mail', label: 'By mail' }] },
				],
			},
		],
	},
];

<WordEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

`size: 'small'` hace que los comandos llenen columnas de tres, y `items` crea un menú desplegable.
Pasar de nuevo un descriptor igual, como hace un framework en cada renderizado, conserva los paneles y solo sustituye los callbacks; omitir la propiedad elimina las pestañas. Una pestaña no puede usar el id de una pestaña integrada, y sus comandos nunca llegan al gestor de comandos del propio editor. El mismo descriptor (`RibbonAddInTab`, de `ooxml-core/ribbon`, reexportado por cada paquete de framework) funciona en los editores de Word, Excel, PowerPoint y Visio. El editor no incluye ninguna pestaña de complemento.

## Propiedades y eventos de los adaptadores {#binding-props-and-events}

Cada adaptador de framework reenvía las mismas propiedades del elemento y vuelve a emitir los mismos eventos, de modo que los nombres siguientes significan lo mismo en todos los adaptadores. Los nombres de evento son los del elemento (`file-command`); cada framework los escribe según su propio idioma.

| Propiedad del elemento | React / Solid / Vanilla | Propiedad en Vue | Input de Angular | Propiedad de Svelte |
| ---------------------- | ----------------------- | ---------------- | ---------------- | ------------------- |
| `fileName`             | `fileName`              | `fileName`       | `fileName`       | `fileName`          |
| `reviewAuthor`         | `reviewAuthor`          | `reviewAuthor`   | `reviewAuthor`   | `reviewAuthor`      |
| `themeColors`          | `themeColors`           | `themeColors`    | `themeColors`    | `themeColors`       |

`fileName`, `reviewAuthor` y `themeColors` se reenvían cuando la propiedad se define o cambia, de modo que un cambio de nombre hecho con Archivo > Abrir no lo revierte un nuevo renderizado que pase el nombre anterior.

| Evento del elemento | Opción React / Solid / Vanilla | Evento en Vue      | Output de Angular | Callback de Svelte  |
| ------------------- | ------------------------------ | ------------------ | ----------------- | ------------------- |
| `file-command`      | `onFileCommand(event)`         | `file-command`     | `fileCommand`     | `onfilecommand`     |
| `document-warning`  | `onDocumentWarning(message)`   | `document-warning` | `documentWarning` | `ondocumentwarning` |
| `readonly-change`   | `onReadOnlyChange(readOnly)`   | `readonly-change`  | `readOnlyChange`  | `onreadonlychange`  |
| `ribbon-action`     | `onRibbonAction(detail)`       | `ribbon-action`    | `ribbonAction`    | `onribbonaction`    |
| `ribbon-customize`  | `onRibbonCustomize(ids)`       | `ribbon-customize` | `ribbonCustomize` | `onribboncustomize` |

`file-command` es el único evento cancelable, así que su gestor recibe el propio evento y no su detalle. Llame a `event.preventDefault()` para asumir un comando de archivo (Guardar, Abrir, Nuevo y similares); `event.detail.command` indica cuál es. Los gestores de Vue, Angular y Svelte reciben el mismo objeto `CustomEvent`. Los demás eventos pasan su detalle directamente.
