# API del elemento

`<xlsx-editor>` (clase `XlsxEditorElement`) es el único editor que monta cada paquete. Regístrelo con `defineXlsxEditor()` (lo exporta cada paquete; el registro es idempotente y no hace nada sin `customElements`), o deje que `mountEditor` y los componentes de framework lo hagan por usted.

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const editor = document.createElement('xlsx-editor');
editor.fileName = 'Budget.xlsx';
document.body.append(editor);
await editor.load(bytes, 'Budget.xlsx');
```

## Atributos y propiedades

| Atributo           | Propiedad        | Tipo                                                                 | Valor por defecto |
| ------------------ | ---------------- | -------------------------------------------------------------------- | ----------------- |
| `locale`           | `locale`         | `en \| fr \| de \| es \| zh-CN` (cualquier etiqueta se asigna a uno) | `en`              |
| `read-only`        | `readOnly`       | boolean                                                              | `false`           |
| `file-name`        | `fileName`       | string                                                               | `Book1.xlsx`      |
| `author-name`      | `authorName`     | string                                                               | `Author`          |
| `theme`            | `theme`          | `light \| dark \| auto`                                              | `auto`            |
| `show-toolbar`     | `showToolbar`    | boolean (`"false"` lo oculta)                                        | `true`            |
| `show-formula-bar` | `showFormulaBar` | boolean (`"false"` lo oculta)                                        | `true`            |

Solo propiedades:

| Propiedad            | Tipo                                    | Notas                                                                                             |
| -------------------- | --------------------------------------- | ------------------------------------------------------------------------------------------------- |
| `themeColors`        | `Partial<Record<...>>`                  | Sustituciones de tokens aplicadas como propiedades personalizadas `--xve-*` ([temas](/theming)).  |
| `hiddenActions`      | `string[]`                              | Controles de la cinta que se ocultan por identificador estable.                                   |
| `ribbonAddIns`       | `RibbonAddInTab[]`                      | Pestañas que el anfitrión añade tras las integradas ([enlaces](/bindings#ribbon-add-in-tabs)).    |
| `workbook`           | `Workbook \| null`                      | Obtiene el libro abierto o, al asignarla, lo sustituye.                                           |
| `dirty`              | `boolean` (solo lectura)                | Ediciones sin guardar.                                                                            |
| `activeSheet`        | `number`                                | Índice, a partir de cero, de la hoja visible.                                                     |
| `collaboration`      | `XlsxCollaborationOptions \| null`      | Comparte el libro abierto en una sala mientras está definida ([colaboración](/es/collaboration)). |
| `collaborationState` | `XlsxCollaborationState` (solo lectura) | Si se está compartiendo, el estado de la conexión y quién está en la sala.                        |

## Métodos

| Método                        | Devuelve              | Notas                                                                                                                                 |
| ----------------------------- | --------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`      | `Promise<void>`       | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv`; el formato se detecta a partir de los bytes. Emite `workbook-error` y se rechaza si falla. |
| `newWorkbook()`               | `void`                | Un libro en blanco llamado `Book1.xlsx`.                                                                                              |
| `save()`                      | `Promise<Blob>`       | `.xlsx`. No borra `dirty`.                                                                                                            |
| `saveBytes(format?)`          | `Promise<Uint8Array>` | `'xlsx'` (por defecto) o `'csv'` (la hoja activa, con valores con formato).                                                           |
| `download(name?)`             | `Promise<void>`       | Guarda, descarga y después llama a `markClean()`.                                                                                     |
| `markClean()`                 | `void`                |                                                                                                                                       |
| `select(ref)`                 | `void`                | Una celda o un rango en notación A1 de la hoja activa (`'C3'`, `'A1:D20'`).                                                           |
| `getSelection()`              | `string`              | La selección actual como referencia A1.                                                                                               |
| `setActiveSheet(index)`       | `void`                |                                                                                                                                       |
| `undo()`, `redo()`            | `void`                | Mientras se comparte, solo se deshacen las ediciones de esta ventana.                                                                 |
| `startCollaboration(options)` | `Promise<void>`       | Se une a una sala; se rechaza si ya se está compartiendo o no hay ningún libro abierto.                                               |
| `stopCollaboration()`         | `void`                | Abandona la sala y borra `collaboration`; el libro queda como está.                                                                   |
| `reconnectCollaboration()`    | `void`                | Reinicia la conexión conservando los cambios hechos sin conexión.                                                                     |
| `share()`                     | `void`                | Abre Archivo > Compartir.                                                                                                             |
| `focusGrid()`                 | `void`                | Mueve el foco del teclado a la cuadrícula para que al escribir se edite la celda activa.                                              |

## Eventos

Todos los eventos son `CustomEvent` que se propagan (bubble) y atraviesan el límite del shadow DOM (`composed`).

| Evento                 | `detail`                                                                                   | Notas                                                                                                |
| ---------------------- | ------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                             | Tras una edición, una carga o un libro nuevo.                                                        |
| `workbook-error`       | `{ error, message }`                                                                       | Falló una carga o un guardado.                                                                       |
| `workbook-warning`     | `{ warnings }`                                                                             | Contenido que el cargador leyó pero no pudo modelar.                                                 |
| `readonly-change`      | `{ readOnly }`                                                                             |                                                                                                      |
| `ribbon-action`        | `{ id }`                                                                                   | Se ejecutó un comando de la cinta.                                                                   |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' }` | **Cancelable**: llame a `preventDefault()` para gestionarlo usted mismo.                             |
| `selection-change`     | `{ sheet, ref, active }`                                                                   | `ref` es el rango seleccionado y `active` la celda activa.                                           |
| `sheet-change`         | `{ index, name }`                                                                          |                                                                                                      |
| `dirty-change`         | `{ dirty }`                                                                                |                                                                                                      |
| `office-ribbon-add-in` | `{ tab, command }`                                                                         | Se eligió un comando de una pestaña del anfitrión (`ribbonAddIns`).                                  |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                               | Se inició o detuvo el uso compartido, cambió la conexión o alguien entró o salió.                    |
| `validation-error`     | `{ message, title?, ref, style? }`                                                         | Una edición de celda en curso no superó la validación y sigue abierta hasta corregirla o cancelarla. |

```ts
editor.addEventListener('file-command', (event) => {
	if (event.detail.command === 'save') {
		event.preventDefault();
		void editor
			.save()
			.then(uploadToMyStorage)
			.then(() => editor.markClean());
	}
});
```

## Partes del shadow DOM

El elemento expone sus regiones como partes CSS del shadow DOM, para aplicar estilos con `xlsx-editor::part(grid)` y para las pruebas: `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs` y `status-bar`.
