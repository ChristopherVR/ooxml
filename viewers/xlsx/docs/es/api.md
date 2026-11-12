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

| Propiedad            | Tipo                                    | Notas                                                                                                                                                                                            |
| -------------------- | --------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `themeColors`        | `Partial<Record<...>>`                  | Sustituciones de tokens aplicadas como propiedades personalizadas `--xve-*` ([temas](/theming)).                                                                                                 |
| `hiddenActions`      | `string[]`                              | Controles de la cinta, elementos de la barra de título y entradas de menú contextual que se ocultan por identificador estable.                                                                   |
| `ribbonAddIns`       | `RibbonAddInTab[]`                      | Pestañas que el anfitrión añade tras las integradas ([enlaces](/bindings#ribbon-add-in-tabs)).                                                                                                   |
| `customization`      | `ExcelCustomization \| null`            | Capacidades, páginas del host en la pestaña Archivo y comandos de menú del host (consulte [Personalización](#customization)).                                                                    |
| `workbook`           | `Workbook \| null`                      | Obtiene el libro abierto o, al asignarla, lo sustituye.                                                                                                                                          |
| `dirty`              | `boolean` (solo lectura)                | Ediciones sin guardar.                                                                                                                                                                           |
| `activeSheet`        | `number`                                | Índice, a partir de cero, de la hoja visible.                                                                                                                                                    |
| `collaboration`      | `XlsxCollaborationOptions \| null`      | Comparte el libro abierto en una sala mientras está definida ([colaboración](/es/collaboration)).                                                                                                |
| `collaborationState` | `XlsxCollaborationState` (solo lectura) | Si se está compartiendo, el estado de la conexión y quién está en la sala.                                                                                                                       |
| `googleFonts`        | `false \| { baseUrl }`                  | Alternativa de fuentes web para las fuentes que el lector no tiene instaladas ni incrustadas. Si no se define, se aplica la opción del Centro de confianza del lector (desactivada por defecto). |
| `surfaceOnly`        | boolean                                 | Defínalo antes de montar para omitir la interfaz de la aplicación; combínelo con `surfaceActions`.                                                                                               |
| `surfaceActions`     | `SurfaceHostActions`                    | Callbacks del host que sustituyen a los cuadros de diálogo cuando `surfaceOnly` está definido. Un flujo sin callback se notifica como no habilitado en `getToolbarProps().enabled`.              |
| `loading`            | `boolean` (solo lectura)                | Indica si se está abriendo un libro.                                                                                                                                                             |

## Personalización {#customization}

`customization` acepta un objeto `ExcelCustomization`, el contrato compartido de `ooxml-core/embedding` que también usan los demás editores. Cada adaptador lo reenvía como la prop o entrada `customization`. Se aplica cuando cambia la identidad del objeto, así que manténgalo estable; `null` restaura los valores predeterminados.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send range', onSelect: ({ ref }) => send(ref) }],
	},
};
```

- `capabilities`: `false` elimina todas las vías a la capacidad (`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`): cinta, páginas de la pestaña Archivo, menús y atajos. `hiddenActions` no cambia; ambos se combinan.
- `backstage.extraPages`: páginas del host `{ id, label, group?, order?, render }` en la pestaña Archivo, tras las páginas integradas de su grupo. `render` es una función de montaje, una etiqueta de elemento personalizado o `{ element, properties }` y recibe un `ExcelRegionContext`.
- `contextMenu.extraCommands`: comandos del host `{ id, label, icon?, group?, when?, disabled?, onSelect }` encima (`top`, predeterminado) o debajo (`bottom`) de las entradas integradas. `onSelect` se ejecuta tras cerrarse el menú con un `ExcelMenuContext`. Las etiquetas se muestran tal cual (localícelas usted mismo).

`edit: false` es un bloqueo de solo lectura propio, junto al interruptor de solo lectura del host y un rol de colaboración `viewer`: ninguno anula a otro. Las capacidades son una política de interfaz, no protección del documento. El reemplazo de regiones (cuadro de opciones, paneles) aún no se ofrece para Excel. Todos los productos juntos: [guía de integración](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Métodos

| Método                           | Devuelve                  | Notas                                                                                                                                                                                                                         |
| -------------------------------- | ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`         | `Promise<void>`           | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv`; el formato se detecta a partir de los bytes. Emite `workbook-error` y se rechaza si falla.                                                                                         |
| `newWorkbook()`                  | `void`                    | Un libro en blanco llamado `Book1.xlsx`.                                                                                                                                                                                      |
| `save()`                         | `Promise<Blob>`           | `.xlsx`. No borra `dirty`.                                                                                                                                                                                                    |
| `saveBytes(format?)`             | `Promise<Uint8Array>`     | `'xlsx'` (por defecto) o `'csv'` (la hoja activa, con valores con formato).                                                                                                                                                   |
| `download(name?)`                | `Promise<void>`           | Guarda, descarga y después llama a `markClean()`.                                                                                                                                                                             |
| `markClean()`                    | `void`                    |                                                                                                                                                                                                                               |
| `select(ref)`                    | `void`                    | Una celda o un rango en notación A1 de la hoja activa (`'C3'`, `'A1:D20'`).                                                                                                                                                   |
| `getSelection()`                 | `string`                  | La selección actual como referencia A1.                                                                                                                                                                                       |
| `setActiveSheet(index)`          | `void`                    |                                                                                                                                                                                                                               |
| `undo()`, `redo()`               | `void`                    | Mientras se comparte, solo se deshacen las ediciones de esta ventana.                                                                                                                                                         |
| `startCollaboration(options)`    | `Promise<void>`           | Se une a una sala; se rechaza si ya se está compartiendo o no hay ningún libro abierto.                                                                                                                                       |
| `stopCollaboration()`            | `void`                    | Abandona la sala y borra `collaboration`; el libro queda como está.                                                                                                                                                           |
| `reconnectCollaboration()`       | `void`                    | Reinicia la conexión conservando los cambios hechos sin conexión.                                                                                                                                                             |
| `share()`                        | `void`                    | Abre el cuadro de diálogo Compartir.                                                                                                                                                                                          |
| `focusGrid()`                    | `void`                    | Mueve el foco del teclado a la cuadrícula para que al escribir se edite la celda activa.                                                                                                                                      |
| `openCloudFile(file)`            | `Promise<void>`           | Abre un archivo que un host obtuvo de un `CloudFileProvider`, como Archivo > Abrir > OneDrive: solo lectura cuando `file.canEdit` es false; Guardar y Guardar como se dirigen a `file.storage`. Se rechaza si la carga falla. |
| `commitEdit()`                   | `boolean`                 | Confirma una edición de celda en curso antes de que el host navegue o guarde. `false` indica que la validación lo bloqueó.                                                                                                    |
| `getToolbarProps()`              | `SpreadsheetToolbarProps` | El estado sin interfaz de la barra de herramientas y las banderas `enabled` para controles personalizados ([bloques de construcción](/building-blocks)).                                                                      |
| `subscribeEditorState(listener)` | `() => void`              | Llama a `listener` en cada cambio de estado del editor; devuelve la función para cancelar la suscripción.                                                                                                                     |

## Eventos

Todos los eventos son `CustomEvent` que se propagan (bubble) y atraviesan el límite del shadow DOM (`composed`).

| Evento                 | `detail`                                                                                              | Notas                                                                                                                                        |
| ---------------------- | ----------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                                        | Tras una edición, una carga o un libro nuevo.                                                                                                |
| `workbook-error`       | `{ error, message }`                                                                                  | Falló una carga o un guardado.                                                                                                               |
| `workbook-warning`     | `{ warnings }`                                                                                        | Contenido que el cargador leyó pero no pudo modelar.                                                                                         |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }`                                                                  | El estado efectivo de solo lectura. `locked` lo establece un rol de colaboración `viewer`; `hostReadOnly` es el interruptor propio del host. |
| `ribbon-action`        | `{ id }`                                                                                              | Se ejecutó un comando de la cinta.                                                                                                           |
| `ribbon-customize`     | `{ hiddenActions }`                                                                                   | El usuario cambió qué comandos de la cinta se muestran (Archivo > Opciones > Personalizar la cinta).                                         |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' \| 'share' }` | **Cancelable**: llame a `preventDefault()` para gestionarlo usted mismo.                                                                     |
| `selection-change`     | `{ sheet, ref, active }`                                                                              | `ref` es el rango seleccionado y `active` la celda activa.                                                                                   |
| `sheet-change`         | `{ index, name }`                                                                                     |                                                                                                                                              |
| `dirty-change`         | `{ dirty }`                                                                                           |                                                                                                                                              |
| `office-ribbon-add-in` | `{ tab, command }`                                                                                    | Se eligió un comando de una pestaña del anfitrión (`ribbonAddIns`).                                                                          |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                                          | Se inició o detuvo el uso compartido, cambió la conexión o alguien entró o salió.                                                            |
| `validation-error`     | `{ message, title?, ref, style? }`                                                                    | Una edición de celda en curso no superó la validación y sigue abierta hasta corregirla o cancelarla.                                         |

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

## Almacenamiento del host {#host-owned-storage}

`storage` (una propiedad del elemento y una prop de cada binding) recibe un `DocumentStorage` de
`ooxml-core/embedding`. Cuando está definida:

- Archivo > Guardar, Guardar como y Ctrl+S entregan el libro guardado a `storage.save(bytes, meta)` en
  lugar de descargarlo. `meta` es `{ fileName, mimeType, reason }`, con `reason` `'save'` o `'saveAs'`;
  los orígenes `.xls` y `.csv` se guardan como `.xlsx`, igual que en una descarga. Cuando la promesa se
  resuelve, el libro queda marcado como sin cambios; si se resuelve con `{ fileName }`, la barra de
  título adopta ese nombre. Un rechazo se notifica como `workbook-error` y el libro sigue con cambios.
- Archivo > Abrir llama a `storage.open()` cuando el almacenamiento lo tiene y abre el
  `{ bytes, fileName }` con el que se resuelve (`undefined` significa que el usuario canceló). Sin
  `open` se usa el selector de archivos.
- Solo se ejecuta un guardado a la vez: Guardar mientras hay un guardado pendiente se ignora, no se
  encola. Una edición hecha mientras se ejecuta el guardado del host mantiene el libro con cambios, ya
  que los bytes guardados no la contienen.
- `file-command` se sigue disparando primero; un host que llama a `preventDefault()` gestiona el comando
  por su cuenta y no se llama a `storage`.

Exportar, Guardar como CSV e Imprimir siguen descargando o imprimiendo en el navegador. Excel no tiene
Autorrecuperación, por lo que no se usan los métodos de recuperación de `DocumentStorage`.

```ts
editor.storage = {
	save: (bytes, { fileName }) => myDrive.put(fileName, bytes),
	open: () => myDrive.pick(), // resolves { bytes, fileName } or undefined
};
```

## Ubicaciones en la nube en Archivo > Abrir {#cloud-places-in-file-open}

`fileProviders` (una propiedad del elemento y una prop de cada binding) recibe un array de
`CloudFileProvider` (`ooxml-core/embedding`), por ejemplo `createOneDriveFileProvider` de
`ooxml-core/embedding/onedrive`. Vacío u omitido, Archivo > Abrir no cambia. Con proveedores, Archivo >
Abrir muestra "Este dispositivo" (Examinar) y una ubicación por proveedor, que muestra el explorador
compartido `<office-ui-cloud-files>` limitado a `.xlsx`. Sus textos siguen el `locale` del editor.

- Al abrir un archivo se pregunta si descartar los cambios sin guardar, como hace Abrir; después se
  cargan los bytes por la ruta de carga normal, se adopta el nombre del archivo y se cierra la vista
  Archivo.
- A partir de ahí el libro se guarda a través del `storage` propio del archivo, no del `storage` del
  host: Guardar y Ctrl+S escriben en el mismo elemento, Guardar como escribe una copia, y el libro solo
  queda marcado como sin cambios cuando el proveedor se ha resuelto. Un guardado rechazado se notifica
  como `workbook-error` y el libro sigue con cambios. Abrir, crear o cargar otro libro devuelve Guardar
  al `storage` del host.
- Un archivo que el proveedor indica como `canEdit: false` se abre en solo lectura (`readOnly` es true,
  se dispara un evento `readonly-change`). Prevalece el más estricto entre el `readOnly` del host y el
  del archivo, y el valor propio del host no se modifica.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken,
		signIn,
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

## Partes del shadow DOM

El elemento expone sus regiones como partes CSS del shadow DOM, para aplicar estilos con `xlsx-editor::part(grid)` y para las pruebas: `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs`, `status-bar`, `title-bar`, `backstage` y `page-layout` (vista Diseño de página).
