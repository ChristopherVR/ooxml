# API del elemento y eventos

Todos los adaptadores de framework envuelven el mismo elemento personalizado `<docx-editor>`. Los adaptadores reenvían estas propiedades y vuelven a emitir estos eventos según las convenciones de cada framework (consulte los [adaptadores de framework](/bindings)).

## Atributos y propiedades

| Propiedad                      | Atributo          | Descripción                                                                                                                                                                                                                                |
| ------------------------------ | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `documentModel`                |                   | El `DocumentModel` mostrado. Asignarlo es una sustitución externa, no un evento de edición.                                                                                                                                                |
| `readOnly`                     | `read-only`       | Bloquea las ediciones locales. Los lotes de colaboración remotos se siguen aplicando.                                                                                                                                                      |
| `locale`                       | `locale`          | Idioma de la interfaz: `en`, `fr`, `de`, `es` o `zh-CN`. El contenido del documento nunca se traduce.                                                                                                                                      |
| `fileName`                     | `file-name`       | Nombre del documento abierto (por defecto `Document1.docx`).                                                                                                                                                                               |
| `reviewAuthor`                 | `review-author`   | Autor registrado en los cambios con seguimiento y en los comentarios (por defecto `Author`).                                                                                                                                               |
| `theme`                        | `theme`           | `light`, `dark` o `auto`. Consulte [temas](/theming).                                                                                                                                                                                      |
| `themeColors`                  |                   | Sustituciones parciales de tokens, aplicadas como propiedades personalizadas `--dve-*`.                                                                                                                                                    |
| `showThumbnails`               | `show-thumbnails` | Panel izquierdo de miniaturas de página (solo en Diseño de impresión).                                                                                                                                                                     |
| `readAloud`                    |                   | Opciones de Leer en voz alta. `stopOnEdit` (por defecto `false`) detiene también la reproducción al editar el documento; reemplazar, cargar o desmontar siempre la detiene.                                                                |
| `showToolbar`                  | `show-toolbar`    | `false` oculta la cinta de opciones; la barra de título y la barra de estado permanecen.                                                                                                                                                   |
| `hiddenActions`                |                   | Controles de la cinta que se ocultan por identificador estable; consulte la tabla en [adaptadores](/bindings).                                                                                                                             |
| `ribbonAddIns`                 |                   | Pestañas que el host añade tras las integradas; consulte [adaptadores](/bindings#ribbon-add-in-tabs).                                                                                                                                      |
| `collaboration`                |                   | Colaboración declarativa (`DocxCollaborationOptions` o `null`); consulte [colaboración](/es/collaboration#la-propiedad-collaboration).                                                                                                     |
| `customization`                |                   | Capacidades, páginas del host en la pestaña Archivo y comandos de menú del host (consulte [Personalización](#customization)).                                                                                                              |
| `surfaceOnly`                  |                   | Monta solo el documento y su superficie de edición. Elíjalo antes de que el elemento se conecte.                                                                                                                                           |
| `surfaceActions`               |                   | Diálogos del host para el modo superficie: `onOpenHyperlinkDialog`, `onOpenFindReplace` y `onOpenImagePicker`.                                                                                                                             |
| `googleFonts`                  |                   | Fuentes web para las fuentes del documento que no están instaladas ni incrustadas. Sin definir, sigue Archivo > Opciones > Centro de confianza (desactivado hasta activarlo); `false` lo desactiva; `{ baseUrl }` define el endpoint CSS2. |
| `storage`                      |                   | Almacenamiento propio del host para Guardar, Guardar como y Abrir; consulte «Almacenamiento del host» más abajo.                                                                                                                           |
| `fileProviders`                |                   | Ubicaciones en la nube para Archivo > Abrir; consulte «Ubicaciones en la nube en Archivo > Abrir» más abajo.                                                                                                                               |
| `loading` (lectura)            |                   | Verdadero mientras se carga un documento.                                                                                                                                                                                                  |
| `collaborationState` (lectura) |                   | La sala de colaboración: si está activa, su id de sala, estado, estado de sincronización y personas.                                                                                                                                       |
| `dirty` (lectura)              |                   | Verdadero tras una edición; falso tras guardar, cargar o llamar a `markClean()`.                                                                                                                                                           |

## Personalización {#customization}

`customization` acepta un objeto `WordCustomization`, el contrato compartido de `ooxml-core/embedding` que también usan los demás editores. Cada adaptador lo reenvía como la prop o entrada `customization`. Se aplica cuando cambia la identidad del objeto, así que manténgalo estable; `null` restaura los valores predeterminados.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [
			{ id: 'send', label: 'Send to chat', onSelect: ({ selectedText }) => send(selectedText) },
		],
	},
};
```

- `capabilities`: `false` elimina todas las vías a la capacidad (`edit`, `print`, `export`, `save`, `open`, `share`, `copy`, `comment`, `review`, `find`, `settings`): cinta, páginas de la pestaña Archivo, menús y atajos. `hiddenActions` no cambia; ambos se combinan.
- `backstage.extraPages`: páginas del host `{ id, label, group?, order?, render }` en la pestaña Archivo, tras las páginas integradas de su grupo. `render` es una función de montaje, una etiqueta de elemento personalizado o `{ element, properties }` y recibe un `WordRegionContext`.
- `contextMenu.extraCommands`: comandos del host `{ id, label, icon?, group?, when?, disabled?, onSelect }` encima (`top`, predeterminado) o debajo (`bottom`) de las entradas integradas. `onSelect` se ejecuta tras cerrarse el menú con un `WordMenuContext`. Las etiquetas se muestran tal cual (localícelas usted mismo).

`edit: false` es un bloqueo de solo lectura propio, junto al interruptor de solo lectura del host y un rol de colaboración `viewer`: ninguno anula a otro. Las capacidades son una política de interfaz, no protección del documento. El reemplazo de regiones (cuadro de opciones, paneles) aún no se ofrece para Word. Todos los productos juntos: [guía de integración](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Métodos

| Método                                                          | Descripción                                                                                                                                             |
| --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes)`                                                   | Carga bytes de DOCX o DOC heredado. Se rechaza y emite `document-error` si falla.                                                                       |
| `save()`                                                        | Se resuelve con un `Blob` del documento serializado.                                                                                                    |
| `saveBytes()`                                                   | Se resuelve con el `Uint8Array` sin procesar.                                                                                                           |
| `download(fileName?)`                                           | Guarda e inicia una descarga del navegador.                                                                                                             |
| `markClean()`                                                   | Borra el indicador de cambios sin guardar cuando el host ha persistido el documento.                                                                    |
| `setLoadedDocument(session)`                                    | Monta un `LoadedDocument` ya analizado por `loadDocument`.                                                                                              |
| `openCloudFile(file)`                                           | Abre un archivo de un `CloudFileProvider`, como hace Archivo > Abrir > OneDrive.                                                                        |
| `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` | Estado y callbacks para los controles de barra de herramientas, navegación y estado que construye el host.                                              |
| `subscribeEditorState(listener)`                                | Llama a `listener` ante cambios de selección y formato, incluidos los que no modifican el documento. Devuelve una función para cancelar la suscripción. |
| `startCollaboration(config)`                                    | Se une a una sesión de colaboración; consulte [colaboración](/es/collaboration).                                                                        |
| `startYjsCollaboration(session, options)`                       | Se une a una sala Yjs sincronizada con su paquete de origen correspondiente.                                                                            |
| `reconnectCollaboration()`, `resyncCollaboration()`             | Recuperan la conexión Yjs o solicitan un intercambio de estado con el proveedor.                                                                        |
| `getPendingCollaboration()`                                     | El lote de pasos pendiente, o `null`.                                                                                                                   |
| `receiveCollaboration(batch)`                                   | Aplica un lote aceptado por la autoridad.                                                                                                               |
| `stopCollaboration(discardPending?)`                            | Abandona la sesión; las ediciones pendientes deben confirmarse o descartarse explícitamente.                                                            |
| `publishPresence`, `receivePresence`, `leavePresence`           | Canal de presencia transitoria; consulte [colaboración](/es/collaboration#presencia).                                                                   |

## Eventos

Todos los eventos burbujean y son compuestos, por lo que atraviesan el límite del shadow DOM.

| Evento                 | Detalle                              | Significado                                                                                          |
| ---------------------- | ------------------------------------ | ---------------------------------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`                      | El documento cambió mediante acciones de edición o de revisión.                                      |
| `document-error`       | `Error`                              | Falló la carga, el guardado, la impresión o una edición.                                             |
| `document-warning`     | `string`                             | Aviso no fatal.                                                                                      |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }` | El usuario alternó Edición/Visualización en la interfaz.                                             |
| `file-command`         | `{ command, fileName? }`             | Cancelable. `command` es `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` o `share`. |
| `dirty-change`         | `boolean`                            | Cambió el estado de cambios sin guardar.                                                             |
| `page-change`          | `{ page, pageCount }`                | Cambió la página o el número de páginas en Diseño de impresión (paginación aproximada).              |
| `ribbon-action`        | `RibbonAction`                       | Se activó un control de la cinta.                                                                    |
| `ribbon-customize`     | `readonly string[]`                  | El usuario cambió los comandos visibles de la cinta; persista los identificadores si lo desea.       |
| `office-ribbon-add-in` | `{ tab, command }`                   | Se eligió un comando de una pestaña del host (`ribbonAddIns`).                                       |
| `presence-send`        | `PresenceMessage`                    | Cambió la presencia local; transpórtela a los demás clientes.                                        |
| `collaboration-send`   | `StepBatch`                          | Hay pasos locales pendientes; transpórtelos a la autoridad.                                          |
| `collaboration-change` | `DocxCollaborationState`             | Cambió el estado de la colaboración: `{ active, roomId, status, synced, people }`.                   |

El paquete exporta el mapa de eventos tipado (`DocxEditorEventMap`) y la lista en tiempo de ejecución `DOCX_EDITOR_EVENTS`.

Los adaptadores de framework reenvían las props `fileName`, `reviewAuthor` y `themeColors`, y los eventos `file-command`, `document-warning`, `readonly-change`, `ribbon-action` y `ribbon-customize`, con los nombres de su framework (por ejemplo `onFileCommand`, `fileCommand` y `onfilecommand`). La tabla está en [adaptadores de framework](/bindings#binding-props-and-events).

## Partes shadow

La barra de título (nombre del archivo, estado de guardado, acceso rápido y «Dígame») es la parte shadow `title-bar`.
Un host que dibuja su propia barra de documento la oculta con `docx-editor::part(title-bar) { display: none; }`;
la cinta y la barra de estado siguen visibles.

## Ubicaciones en la nube en Archivo > Abrir

`fileProviders` (una propiedad del elemento y una prop de cada adaptador) acepta
`CloudFileProvider`s de `ooxml-core/embedding`, por ejemplo `createOneDriveFileProvider`. Archivo >
Abrir muestra entonces "Este dispositivo" y una ubicación por proveedor. Un archivo abierto desde
una ubicación se guarda mediante su propio almacenamiento y es de solo lectura cuando el proveedor
indica que no se puede editar; consulte [integración con OneDrive](/onedrive#open-from-onedrive-in-file-open).

## Almacenamiento del host

`storage` (una propiedad del elemento y una prop de cada adaptador) acepta un `DocumentStorage` de
`ooxml-core/embedding`. Cuando está definido:

- Archivo > Guardar, Guardar como y Ctrl+S entregan los bytes guardados a
  `storage.save(bytes, meta)` en lugar de descargarlos. `meta` es `{ fileName, mimeType, reason }`,
  con `reason` igual a `'save'` o `'saveAs'`. Cuando se resuelve, el documento se marca como
  guardado; si se resuelve con `{ fileName }`, la barra de título adopta ese nombre. Un rechazo se
  notifica como `document-error` y el documento sigue con cambios sin guardar.
- Archivo > Abrir llama a `storage.open()` cuando el almacenamiento lo ofrece y abre el
  `{ bytes, fileName }` con el que se resuelve (`undefined` significa que el usuario canceló). Sin
  `open` se usa el selector de archivos.
- Solo se ejecuta un guardado a la vez: Guardar mientras hay un guardado pendiente se ignora, no se
  pone en cola. Una edición hecha mientras se ejecuta el guardado del host mantiene el documento
  con cambios sin guardar, ya que los bytes guardados no la contienen.
- `file-command` sigue emitiéndose primero. Un host que llama a `preventDefault()` gestiona el
  comando por sí mismo y no se llama a `storage`. Guardar como anuncia `saveAs` y después `save`.
- Un host que eliminó Guardar (`customization.capabilities.save: false`) no recibe ni el evento
  `save` ni `storage.save`; Guardar como (la capacidad `export`) entonces solo descarga.

Exportar, Exportar como texto e Imprimir siguen usando el navegador. Word no tiene
Autorrecuperación, así que los métodos de recuperación de `DocumentStorage` no se usan.

```ts
editor.storage = {
	async save(bytes, { fileName, mimeType }) {
		await fetch(`/files/${encodeURIComponent(fileName)}`, {
			method: 'PUT',
			headers: { 'content-type': mimeType },
			body: bytes,
		});
	},
};
```

## Lectura en voz alta {#read-aloud}

El hook de bloques de construcción (`useViewerBuildingBlocks`) añade cinco campos de lectura en voz alta a
`toolbarProps`. No son propiedades ni eventos del elemento. Usan la implementación del propio editor (la
síntesis de voz del navegador; no se envía nada). La [guía de bloques de construcción](/es/building-blocks#read-aloud)
describe su comportamiento.

| Campo             | Tipo                     | Comportamiento                                                                                                                                             |
| ----------------- | ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `canReadAloud`    | `boolean`                | Verdadero cuando hay un documento en pantalla y el navegador puede hablar. Los documentos de solo lectura cuentan.                                         |
| `readingAloud`    | `boolean`                | Verdadero solo en el editor que posee la voz. Si otro editor está leyendo, es falso.                                                                       |
| `readAloudIssue`  | `ReadAloudIssue \| null` | El último problema: `'unsupported'` mientras el navegador no tenga síntesis de voz, `'empty'` o `'error'`. `null` cuando no hay ninguno.                   |
| `onReadAloud`     | `() => ReadAloudStart`   | Lee la selección o desde el cursor hasta el final si no hay selección. Reinicia si ya está leyendo. Devuelve `'started'` o el motivo por el que no empezó. |
| `onStopReadAloud` | `() => boolean`          | Detiene la reproducción de este editor. Devuelve si estaba reproduciendo.                                                                                  |

Los códigos son uniones de cadenas simples. La biblioteca solo aporta los códigos; el host redacta el mensaje.

```ts
type ReadAloudIssue = 'unsupported' | 'empty' | 'error';
type ReadAloudStart = 'started' | ReadAloudIssue;
```

`'error'` significa que la voz falló. Un enunciado cancelado o interrumpido no se notifica como error.

Un botón de React del host que alterna la lectura y avisa del caso vacío:

```tsx
import { DocumentCanvas, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel }) {
	const { canvasProps, toolbarProps } = useViewerBuildingBlocks({ documentModel });
	return (
		<>
			{toolbarProps && toolbarProps.readAloudIssue !== 'unsupported' && (
				<>
					<button
						disabled={!toolbarProps.canReadAloud}
						aria-pressed={toolbarProps.readingAloud}
						onClick={() =>
							toolbarProps.readingAloud
								? toolbarProps.onStopReadAloud()
								: toolbarProps.onReadAloud()
						}
					>
						{toolbarProps.readingAloud ? 'Stop reading' : 'Read aloud'}
					</button>
					{toolbarProps.readAloudIssue === 'empty' && <span role="status">Nothing to read.</span>}
				</>
			)}
			<DocumentCanvas {...canvasProps} />
		</>
	);
}
```

La reproducción se detiene cuando el documento se reemplaza o se vuelve a cargar y cuando el editor se desmonta; una edición normal solo la detiene con `readAloud: { stopOnEdit: true }`.
La `Toolbar` opcional dibuja su propio interruptor «Read aloud» a partir de los mismos campos.
