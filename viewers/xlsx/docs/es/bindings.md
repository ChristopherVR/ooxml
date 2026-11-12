# Bindings de framework

Todos los bindings montan `<xlsx-editor>` mediante la misma función `mountEditor` de `packages/bindings/src/index.ts`, de modo que cada framework tiene una semántica idéntica.

## Props {#props}

| Prop             | Tipo                          | Por defecto  | Notas                                                                                                                                                                            |
| ---------------- | ----------------------------- | ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook`       | `Workbook`                    | ninguno      | Se muestra cuando se pasa un **objeto nuevo**. El libro que emite el editor nunca se vuelve a asignar (sin bucle).                                                               |
| `bytes`          | `Uint8Array \| ArrayBuffer`   | ninguno      | Bytes del archivo; se vuelve a cargar cada vez que se pasa un array nuevo.                                                                                                       |
| `src`            | `string`                      | ninguno      | URL que se descarga y carga cuando cambia la cadena; el nombre del archivo por defecto es el último segmento de la URL.                                                          |
| `fileName`       | `string`                      | `Book1.xlsx` | Se reenvía solo cuando el padre lo cambia, así que Archivo > Abrir en el editor puede renombrar el libro.                                                                        |
| `readOnly`       | `boolean`                     | `false`      |                                                                                                                                                                                  |
| `locale`         | `string`                      | `en`         | `en`, `fr`, `de`, `es`, `zh-CN` o una etiqueta que se asigne a una de ellas (`de-DE`).                                                                                           |
| `theme`          | `'light' \| 'dark' \| 'auto'` | `auto`       | `auto` sigue el sistema operativo.                                                                                                                                               |
| `authorName`     | `string`                      | `Author`     | Se registra en los comentarios nuevos.                                                                                                                                           |
| `showToolbar`    | `boolean`                     | `true`       | La cinta de opciones.                                                                                                                                                            |
| `showFormulaBar` | `boolean`                     | `true`       | El cuadro de nombres y la barra de fórmulas.                                                                                                                                     |
| `hiddenActions`  | `string[]`                    | `[]`         | Controles de la cinta, elementos de la barra de título y entradas del menú contextual que se ocultan, por id estable.                                                            |
| `ribbonAddIns`   | `RibbonAddInTab[]`            | ninguno      | Pestañas que el host añade después de las integradas; vea [Pestañas de complementos de la cinta](#ribbon-add-in-tabs).                                                           |
| `themeColors`    | `Partial<Record<...>>`        | `{}`         | Sustituciones de tokens, vea [temas](/es/theming).                                                                                                                               |
| `collaboration`  | `XlsxCollaborationOptions`    | ninguno      | Comparte el libro abierto en una sala mientras está definido; un objeto nuevo vuelve a unirse, `null` sale. Vea [colaboración](/es/collaboration).                               |
| `storage`        | `DocumentStorage`             | ninguno      | Almacenamiento gestionado por el host: Guardar y Guardar como van a `storage.save()`, y Archivo > Abrir a `storage.open()` cuando lo tiene.                                      |
| `fileProviders`  | `CloudFileProvider[]`         | ninguno      | Ubicaciones en la nube que Archivo > Abrir ofrece junto a "Este dispositivo"; un libro abierto desde una de ellas se guarda de vuelta en ese archivo. Mantenga estable el array. |
| `customization`  | `ExcelCustomization`          | ninguno      | `capabilities`, `backstage.extraPages` y `contextMenu.extraCommands`; mantenga estable el objeto (`null` restaura los valores por defecto).                                      |

## Callbacks y eventos {#callbacks-and-events}

| Callback (React, Solid, vanilla) | Evento de Vue          | Salida de Angular     | Prop de Svelte          | Carga útil                                   |
| -------------------------------- | ---------------------- | --------------------- | ----------------------- | -------------------------------------------- |
| `onWorkbookChange`               | `workbook-change`      | `workbookChange`      | `onworkbookchange`      | `Workbook`                                   |
| `onWorkbookError`                | `workbook-error`       | `workbookError`       | `onworkbookerror`       | `Error`                                      |
| `onSelectionChange`              | `selection-change`     | `selectionChange`     | `onselectionchange`     | `{ sheet, ref, active }`                     |
| `onDirtyChange`                  | `dirty-change`         | `dirtyChange`         | `ondirtychange`         | `boolean`                                    |
| `onReadOnlyChange`               | `readonly-change`      | `readOnlyChange`      | `onreadonlychange`      | `boolean`                                    |
| `onFileCommand`                  | `file-command`         | `fileCommand`         | `onfilecommand`         | `CustomEvent<FileCommandDetail>`, cancelable |
| `onWorkbookWarning`              | `workbook-warning`     | `workbookWarning`     | `onworkbookwarning`     | `string[]`                                   |
| `onRibbonAction`                 | `ribbon-action`        | `ribbonAction`        | `onribbonaction`        | `{ id }`                                     |
| `onRibbonCustomize`              | `ribbon-customize`     | `ribbonCustomize`     | `onribboncustomize`     | `string[]` (acciones ocultas)                |
| `onSheetChange`                  | `sheet-change`         | `sheetChange`         | `onsheetchange`         | `{ index, name }`                            |
| `onRibbonAddIn`                  | `office-ribbon-add-in` | `ribbonAddIn`         | `onribbonaddin`         | `{ tab, command }`                           |
| `onCollaborationChange`          | `collaboration-change` | `collaborationChange` | `oncollaborationchange` | `XlsxCollaborationState`                     |
| `onReady`                        | `ready`                | `ready`               | `onready`               | el elemento `<xlsx-editor>`                  |

Los fallos de carga (archivos dañados, cifrados o no admitidos) llegan a `onWorkbookError`; una descarga fallida de `src` también. Una descarga de `src` que sigue en curso se descarta cuando llegan `bytes` nuevos o un `workbook` nuevo, o cuando `src` se borra o cambia.

`onFileCommand` recibe el evento cancelable `file-command` en sí, no su detalle: llame a `event.preventDefault()` para tomar el control de un comando de Archivo (Nuevo, Abrir, Guardar, Guardar como, Exportar, Exportar CSV, Imprimir, Compartir), y lea `event.detail.command` para ver cuál es. Vue, Angular y Svelte emiten o llaman de vuelta con el mismo objeto de evento. Cada otro callback recibe su carga útil directamente. Vea la [API del elemento](/es/api).

Cada prop se reenvía solo cuando cambia su valor (`hiddenActions` y `themeColors` se comparan por valor), de modo que un nuevo render del padre nunca deshace lo que el usuario cambió dentro del editor: Edición / Visualización, Archivo > Opciones (tema o idioma) o Personalizar cinta. Un padre controlado mantiene su estado sincronizado mediante `onReadOnlyChange` y `onRibbonCustomize` (en Angular, `[(readOnly)]`).

## Pestañas de complementos de la cinta {#ribbon-add-in-tabs}

Un host puede añadir sus propias pestañas de cinta después de las de Excel, como lo hace un complemento de Office.
`ribbonAddIns` acepta datos simples y es una prop de todos los bindings; un comando elegido
ejecuta su callback `run` y se notifica a través de `onRibbonAddIn` (Vue
`office-ribbon-add-in`, Angular `ribbonAddIn`, Svelte `onribbonaddin`) con
`{ tab, command }`.

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

<SpreadsheetEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

Los comandos con `size: 'small'` llenan columnas de tres y `items` crea un menú desplegable.
Pasar de nuevo un descriptor igual, como hace un framework en cada render, mantiene los paneles
y solo sustituye los callbacks; omitir la prop elimina las pestañas. Una pestaña no puede usar el
id de una pestaña integrada, y sus comandos nunca llegan al manejo de comandos propio del editor.
El mismo descriptor (`RibbonAddInTab`, de `ooxml-core/ribbon`, reexportado por cada paquete de
framework) funciona en los editores de Word, Excel, PowerPoint y Visio. Ningún editor trae
pestañas de complemento de serie.

## Handle {#handle}

Cada adaptador expone el mismo handle: React a través de `ref`, Vue a través de la ref de plantilla, Angular a través de la instancia del componente, Solid a través de `editorRef`, Svelte a través de las exportaciones del componente y JavaScript puro como valor devuelto por `mountEditor`.

| Miembro                         | Efecto                                                                                                                                                    |
| ------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `element`                       | El `<xlsx-editor>` montado.                                                                                                                               |
| `load(bytes, fileName?)`        | Abre bytes `.xlsx`, `.xlsm`, `.xls` o `.csv`; rechaza si falla.                                                                                           |
| `newWorkbook()`                 | Sustituye el libro por uno en blanco.                                                                                                                     |
| `save()`                        | Se resuelve con un Blob `.xlsx`. No borra `dirty`.                                                                                                        |
| `saveBytes(format?)`            | Bytes `'xlsx'` (por defecto) o `'csv'` (hoja activa).                                                                                                     |
| `download(fileName?)`           | Guarda, descarga en el navegador y marca el libro como limpio.                                                                                            |
| `markClean()`                   | Borra `dirty` después de que usted haya persistido un Blob guardado.                                                                                      |
| `select(ref)`, `getSelection()` | Selección como referencia A1 (`'B2:D8'`).                                                                                                                 |
| `setActiveSheet(index)`         | Cambia de hoja.                                                                                                                                           |
| `share()`                       | Abre el diálogo Compartir. Desmontar el editor abandona cualquier sala compartida.                                                                        |
| `openCloudFile(file)`           | Abre un archivo de un `CloudFileProvider` como lo hace Archivo > Abrir > OneDrive: su nombre, su estado de solo lectura y el almacenamiento para Guardar. |
| `dirty`                         | Ediciones sin guardar. El `isDirty()` de Svelte está obsoleto; lea `dirty`.                                                                               |

Llamar a un método del handle antes de que el editor esté montado lanza `Editor is not mounted`.

## Renderizado en el servidor {#server-rendering}

Todos los paquetes se importan sin DOM. El registro (`defineXlsxEditor`) y el montaje son operaciones del cliente.
