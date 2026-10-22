# API del elemento y eventos

Todos los adaptadores de framework envuelven el mismo elemento personalizado `<docx-editor>`. Los adaptadores reenvían estas propiedades y vuelven a emitir estos eventos según las convenciones de cada framework (consulte los [adaptadores de framework](/bindings)).

## Atributos y propiedades

| Propiedad         | Atributo          | Descripción                                                                                                                   |
| ----------------- | ----------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `documentModel`   |                   | El `DocumentModel` mostrado. Asignarlo es una sustitución externa, no un evento de edición.                                   |
| `readOnly`        | `read-only`       | Bloquea las ediciones locales. Los lotes de colaboración remotos se siguen aplicando.                                         |
| `locale`          | `locale`          | Idioma de la interfaz: `en`, `fr`, `de`, `es` o `zh-CN`. El contenido del documento nunca se traduce.                         |
| `fileName`        | `file-name`       | Nombre del documento abierto (por defecto `Document1.docx`).                                                                  |
| `reviewAuthor`    | `review-author`   | Autor registrado en los cambios con seguimiento y en los comentarios (por defecto `Author`).                                  |
| `theme`           | `theme`           | `light`, `dark` o `auto`. Consulte [temas](/theming).                                                                         |
| `themeColors`     |                   | Sustituciones parciales de tokens, aplicadas como propiedades personalizadas `--dve-*`.                                       |
| `showThumbnails`  | `show-thumbnails` | Panel izquierdo de miniaturas de página (solo en Diseño de impresión).                                                        |
| `showToolbar`     | `show-toolbar`    | `false` oculta la cinta de opciones; la barra de título y la barra de estado permanecen.                                      |
| `hiddenActions`   |                   | Controles de la cinta que se ocultan por identificador estable; consulte la tabla en [adaptadores](/bindings).                |
| `ribbonAddIns`    |                   | Pestañas que el host añade tras las integradas; consulte [adaptadores](/bindings#ribbon-add-in-tabs).                         |
| `customization`   |                   | Capacidades, páginas del host en la pestaña Archivo y comandos de menú del host (consulte [Personalización](#customization)). |
| `dirty` (lectura) |                   | Verdadero tras una edición; falso tras guardar, cargar o llamar a `markClean()`.                                              |

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

| Método                                                | Descripción                                                                                  |
| ----------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| `load(bytes)`                                         | Carga bytes de DOCX o DOC heredado. Se rechaza y emite `document-error` si falla.            |
| `save()`                                              | Se resuelve con un `Blob` del documento serializado.                                         |
| `saveBytes()`                                         | Se resuelve con el `Uint8Array` sin procesar.                                                |
| `download(fileName?)`                                 | Guarda e inicia una descarga del navegador.                                                  |
| `markClean()`                                         | Borra el indicador de cambios sin guardar cuando el host ha persistido el documento.         |
| `startCollaboration(config)`                          | Se une a una sesión de colaboración; consulte [colaboración](/es/collaboration).             |
| `startYjsCollaboration(session, options)`             | Se une a una sala Yjs sincronizada con su paquete de origen correspondiente.                 |
| `reconnectCollaboration()`, `resyncCollaboration()`   | Recuperan la conexión Yjs o solicitan un intercambio de estado con el proveedor.             |
| `getPendingCollaboration()`                           | El lote de pasos pendiente, o `null`.                                                        |
| `receiveCollaboration(batch)`                         | Aplica un lote aceptado por la autoridad.                                                    |
| `stopCollaboration(discardPending?)`                  | Abandona la sesión; las ediciones pendientes deben confirmarse o descartarse explícitamente. |
| `publishPresence`, `receivePresence`, `leavePresence` | Canal de presencia transitoria; consulte [colaboración](/es/collaboration#presencia).        |

## Eventos

Todos los eventos burbujean y son compuestos, por lo que atraviesan el límite del shadow DOM.

| Evento                 | Detalle               | Significado                                                                                    |
| ---------------------- | --------------------- | ---------------------------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`       | El documento cambió mediante acciones de edición o de revisión.                                |
| `document-error`       | `Error`               | Falló la carga, el guardado, la impresión o una edición.                                       |
| `document-warning`     | `string`              | Aviso no fatal.                                                                                |
| `readonly-change`      | `boolean`             | El usuario alternó Edición/Visualización en la interfaz.                                       |
| `file-command`         | `{ command }`         | Cancelable. `command` es `new`, `open`, `save`, `export`, `print` o `share`.                   |
| `dirty-change`         | `boolean`             | Cambió el estado de cambios sin guardar.                                                       |
| `page-change`          | `{ page, pageCount }` | Cambió la página o el número de páginas en Diseño de impresión (paginación aproximada).        |
| `ribbon-action`        | `RibbonAction`        | Se activó un control de la cinta.                                                              |
| `ribbon-customize`     | `readonly string[]`   | El usuario cambió los comandos visibles de la cinta; persista los identificadores si lo desea. |
| `office-ribbon-add-in` | `{ tab, command }`    | Se eligió un comando de una pestaña del host (`ribbonAddIns`).                                 |
| `presence-send`        | `PresenceMessage`     | Cambió la presencia local; transpórtela a los demás clientes.                                  |
| `collaboration-send`   | `StepBatch`           | Hay pasos locales pendientes; transpórtelos a la autoridad.                                    |

El paquete exporta el mapa de eventos tipado (`DocxEditorEventMap`) y la lista en tiempo de ejecución `DOCX_EDITOR_EVENTS`.
