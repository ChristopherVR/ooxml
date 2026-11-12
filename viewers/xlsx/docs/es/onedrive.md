# Integración con OneDrive

Para una configuración paso a paso (registro en Microsoft Entra, inicio de sesión y un ejemplo completo), empiece por la [guía de inicio rápido de OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

## Archivo > Abrir > OneDrive {#file-open-onedrive}

Pase un proveedor como `fileProviders` y Archivo > Abrir muestra una ubicación de OneDrive junto a "Este dispositivo":

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your sign-in owns the token
		signIn,
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

La ubicación muestra los archivos recientes y Mis archivos, busca y abre un vínculo para compartir. Un libro elegido se carga como cualquier otro, se guarda de vuelta en el mismo elemento con Guardar (una copia con Guardar como), y se marca como limpio solo cuando la subida se resuelve; si la subida falla, sigue sin guardar. Un archivo que `canWrite` rechaza (otra unidad, un archivo compartido) se abre en solo lectura. `fileProviders` es una prop de todos los bindings de framework (`fileProviders` en React, Vue, Solid, Svelte, Angular y JavaScript puro). La aplicación de demostración (`demos/xlsx`) pasa por defecto el proveedor de OneDrive de la demo; configúrelo con `VITE_ONEDRIVE_CLIENT_ID` en `demos/xlsx/demo-vanilla/.env.local` o pegue un token de Graph Explorer en su diálogo de inicio de sesión. El cableado de `storage` de más bajo nivel que se describe a continuación es para hosts que construyen su propia interfaz de Abrir.

El mismo contrato de almacenamiento también admite proveedores personalizados. Vea
[configuración de almacenamiento y fusión independiente del proveedor](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)
para `createStorageSession`, las sustituciones de fusión por formato y la configuración de arranque de la suite.
Use estos puntos de extensión con todos los paquetes de framework de Excel.

Excel usa la conexión compartida `ooxml-core/embedding/onedrive`. Resuelva el vínculo para compartir en el host, verifique los permisos específicos del elemento y entregue al editor existente los bytes XLSX descargados:

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
await editor.load(opened.bytes, opened.file.name);
```

El elemento y los handles de framework exponen `saveBytes('xlsx')`, `load(bytes, fileName?)` y
`markClean()`. Un guardado personalizado del host llama a `session.save(await editor.saveBytes('xlsx'))`; marque como limpio
solo cuando la persistencia haya tenido éxito y no haya ediciones más recientes. Use la ruta de guardado normal del editor para confirmar una edición de celda activa. CSV es una exportación aparte y nunca debe reemplazar el elemento XLSX conectado.
Guardar nativo delega en `storage` salvo que el host impida el evento cancelable `file-command`.

El mismo contrato de `storage` y `readOnly` llega a `@christophervr/xlsx-react-viewer`,
`xlsx-vue-viewer`, `xlsx-angular-viewer`, `xlsx-svelte-viewer`, `xlsx-solid-viewer` y
`xlsx-vanilla-viewer`, así como al elemento simple `ooxml-ui/xlsx`. `@christophervr/xlsx-core`
es la fachada del modelo, no un servicio de autenticación.

Vea la [referencia del paquete de OneDrive de Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#excel-xlsx)
para la autenticación, el uso compartido y la recuperación que conserva tanto la versión local como la remota. La
[guía de pruebas](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) explica
los accesorios de prueba sin credenciales y la aceptación con una cuenta real; no se incluye ninguna interfaz ni servicio de simulación.

Registre `mergeWorkbookSnapshots` de `ooxml-core/xlsx/collab` como callback `merge` de la sesión
para combinar las ediciones independientes admitidas del libro durante Guardar. La suite lo activa
automáticamente; las ediciones de celdas que se solapan y los cambios no admitidos siguen necesitando recuperación. Esto no conecta
las sesiones de colaboración del libro con la coedición de Microsoft Excel, no actualiza datos externos ni cambia
las [limitaciones de funciones](/es/features).

## Cuenta > Servicios conectados {#account-connected-services}

Los mismos `fileProviders` también rellenan la página **Cuenta** de la vista Archivo con una lista de "Servicios conectados"
(el `<office-ui-connected-services>` compartido): una fila por proveedor con la cuenta con la que se ha iniciado sesión y los botones Conectar o Desconectar, que llaman al `signIn` y `signOut` propios del proveedor. Un proveedor sin
`signOut` muestra que el host gestiona su sesión. La lista sigue los `subscribe` de los proveedores, así que iniciar sesión desde Archivo > Abrir la actualiza, y sus textos proceden de los diccionarios de Excel en cada idioma. Sin `fileProviders` no hay página Cuenta.

## Abrir un archivo de la nube desde su propia interfaz {#open-a-cloud-file-from-your-own-ui}

Archivo > Abrir solo es accesible cuando hay un documento abierto. Para abrir un archivo desde una página de inicio o desde su propio selector, obtenga un `CloudOpenedFile` de cualquier `CloudFileProvider` (el explorador compartido `<office-ui-cloud-files>` lo emite en `office-cloud-file-open`) y páselo a `editor.openCloudFile(file)` en `<xlsx-editor>` (el handle de cada binding también lo expone). Sigue la misma
ruta que Archivo > Abrir > OneDrive: el nombre del archivo, solo lectura cuando `canEdit` es false, y Guardar dirigido al
almacenamiento propio del archivo hasta que se cargue otro documento.
