# Almacenamiento gestionado por el host

Para un proveedor de archivos existentes de Microsoft Graph con una demostración sin registro de aplicación, consulte [Integración con OneDrive](/es/guide/onedrive) y [pruebas de OneDrive](/es/guide/onedrive-testing).

Cada enlace acepta una prop `storage` (en Angular: entrada; en Vanilla: opción) que entrega la E/S de archivos del visualizador al host. Es el contrato `DocumentStorage` de `ooxml-core/embedding`, el mismo que aceptan los editores de Word, Excel y Visio, y `ooxml-ui/pptx` lo reexporta como tipo.

```ts
import type { DocumentStorage } from 'ooxml-ui/pptx';

const storage: DocumentStorage = {
	// Required: File > Save and Save As.
	async save(bytes, { fileName, mimeType, reason }) {
		await uploadToMyServer(fileName, bytes, mimeType);
		// Optional: report the name the document now has.
		return { fileName };
	},
	// Optional, all three together: AutoRecover copies.
	readRecovery: (id) => myRecovery.get(id),
	writeRecovery: (entry) => myRecovery.put(entry),
	clearRecovery: (id) => myRecovery.delete(id),
	// Optional: lets Version History list the kept copies.
	listRecovery: () => myRecovery.list(),
};
```

## Guardar y Guardar como {#save-and-save-as}

Sin `storage`, el visualizador descarga el archivo guardado. Con él, Guardar y cada formato de Guardar como (`.pptx`, `.ppsx`, `.pptm`, `.ppt`) llaman a `storage.save(bytes, meta)` en su lugar:

- `meta.fileName` es el nombre base del origen más la extensión del formato escrito (`savedPresentationFileName`), así que `report.ppt` se guarda como `report.pptx`.
- `meta.mimeType` sigue al formato (`.ppt` es el tipo OLE2, `.ppsx` y `.pptm` tienen los suyos propios).
- `meta.reason` es el comando que se ejecutó: `'save'` para Archivo > Guardar (la barra de título, el elemento Guardar del backstage, la barra de herramientas móvil y `downloadPptx()` en Vanilla/Svelte), y `'saveAs'` para cada entrada de Guardar como, incluida Guardar como > Presentación de PowerPoint (.pptx).

Un guardado fallido (la función `save()` del host rechazó la promesa, o la serialización falló) se informa de la misma forma en los cinco enlaces: el evento `error` del enlace (`onError`, `@error`, `(error)`, `onerror`) recibe `Save failed: <reason>`, el usuario ve el aviso "Save failed" en la pila de avisos del visualizador, y la copia de AutoRecover se conserva (solo se descarta tras un guardado correcto y si Opciones > Guardar lo pide). Los métodos de guardado resuelven la promesa en lugar de rechazarla.

Los bytes son los mismos que devuelve `getContent()`, así que una presentación protegida con contraseña se entrega cifrada.

## Copias de recuperación {#recovery-copies}

El autoguardado de recuperación (las props `autosave` y `autosaveIntervalMs`, indexadas por `filePath`) escribe normalmente en la base de datos IndexedDB propia de la biblioteca. Cuando `storage` proporciona `readRecovery`, `writeRecovery` y `clearRecovery`, cada ruta de recuperación las usa en su lugar: el escritor de autoguardado, el aviso de recuperación que aparece al cargar (y su botón Descartar), la limpieza tras un guardado correcto, el Historial de versiones, Archivo > Abrir > Reciente, Cuenta > Almacenamiento y privacidad y `restoreSessionDeck(storage)`. El ID de la entrada es la clave `filePath`; `savedAt` es una marca de tiempo en milisegundos. Un almacenamiento que solo proporcione algunos de los tres métodos se ignora para la recuperación, y IndexedDB sigue en uso.

Las listas (Recientes, uso de la cuenta, Historial de versiones) necesitan `listRecovery`; sin él, esas listas están vacías. La limpieza de Opciones > Guardar ("Eliminar archivos en caché", "borrar caché al cerrar" y la purga por "retención de caché") borra las copias del host solo si proporciona `listRecovery`. Sin él, el visualizador no toca las copias del host, y la retención es responsabilidad del host.

## Sin evento `file-command` {#no-file-command-event}

Los elementos de Word, Excel y Visio anuncian un evento `file-command` cancelable antes de Guardar, Abrir y Nuevo. Los enlaces de PowerPoint no tienen ese evento: para tomar el control del guardado, pase `storage`; para tomar el control de Archivo > Abrir, use `onOpenFile` (en Vue y Angular: la entrada `onOpenFile`).
