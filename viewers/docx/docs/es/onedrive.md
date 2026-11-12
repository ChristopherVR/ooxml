# Integración con OneDrive

Para una configuración paso a paso (registro en Microsoft Entra, inicio de sesión y un ejemplo completo),
empiece por la [guía rápida de OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

Para otros servicios o flujos personalizados de OneDrive, use la misma propiedad `storage` con
[`createStorageSession` y la guía de integración independiente del proveedor](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md).
Los anfitriones pueden sustituir la autenticación, el transporte, los gestores de fusión, la interfaz de
conflictos y la política de compartición sin cambiar ningún binding de Word para frameworks.

Todos los paquetes de Word usan la conexión compartida `ooxml-core/embedding/onedrive`. El anfitrión resuelve
el enlace de compartición, proporciona la autenticación de Microsoft y el permiso de escritura verificado,
y conecta la sesión resultante al contrato de almacenamiento existente del editor.

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
editor.fileName = opened.file.name;
await editor.load(opened.bytes);
```

El elemento `<docx-editor>` simple de `ooxml-ui/docx` expone `saveBytes()`. Cada manejador de framework expone
`save()`, que devuelve un `Blob`, además de `load(bytes)`, `markClean()` y `dirty`. Un botón Guardar
personalizado puede convertir el Blob en bytes y llamar a `session.save(bytes)`; marque como limpio solo
cuando la persistencia tenga éxito y no haya intervenido una edición más reciente. El guardado nativo pasa
automáticamente por `storage`, salvo que el anfitrión impida el evento `file-command` cancelable.

## Abrir desde OneDrive en Archivo > Abrir {#open-from-onedrive-in-file-open}

Defina `fileProviders` (una propiedad del elemento y una propiedad de todos los bindings) con una lista de
`CloudFileProvider`, y Archivo > Abrir mostrará ubicaciones: "Este dispositivo" (el botón Examinar) y una
ubicación por cada proveedor, cada una con el explorador compartido `<office-ui-cloud-files>` (Recientes,
Mis archivos, búsqueda, abrir desde un enlace de compartición) filtrado a `.docx`. Sus etiquetas proceden de
los diccionarios de Word en todos los idiomas de visualización.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDocxSnapshots } from 'ooxml-core/docx';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your Microsoft authentication
		signIn, // opens your sign-in UI and resolves with the account
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			docx: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

Un archivo elegido en una ubicación se carga como cualquier otro documento, toma su nombre de archivo y es de
solo lectura cuando el proveedor informa `canEdit: false` (prevalece el más restrictivo entre el `readOnly` del
anfitrión y el archivo). Archivo > Guardar y Guardar como van al almacenamiento propio de ese archivo en lugar
del `storage` del anfitrión hasta que se cargue otro documento, y la barra de título muestra guardado solo
cuando la subida se ha resuelto; una subida fallida se queda sin guardar y lanza `document-error`. Abrir desde
una ubicación no dispara `file-command` (el anfitrión eligió los proveedores); Guardar y Guardar como sí lo
hacen, y un anfitrión que los cancele los gestiona él mismo. Un anfitrión que elimine la capacidad `open`
oculta también la página y las ubicaciones.

La demo de Word pasa `createDemoOneDriveProvider` de `demos/onedrive` por defecto. Inicie sesión desde
Archivo > Abrir > OneDrive: defina `VITE_ONEDRIVE_CLIENT_ID` (y, opcionalmente, `VITE_ONEDRIVE_AUTHORITY`) en
`demos/docx/demo-vanilla/.env.local` para usar su propio registro de aplicación de Entra mediante MSAL, escriba
un ID de cliente en el diálogo de la demo, o pegue un token de Graph Explorer en su campo "Access token".
Registre el origen de la demo (con una barra final) como URI de redirección de aplicación de página única. La
demo solo edita archivos de la unidad propia del usuario que ha iniciado sesión; los demás archivos se abren
en solo lectura.

Esto se aplica a `docx-react-viewer`, `docx-vue-viewer`, `docx-angular-viewer`, `docx-svelte-viewer`,
`docx-solid-viewer` y `docx-vanilla-viewer`. Cada uno reenvía las mismas propiedades u opciones `storage` y
`readOnly`. `docx-core` es una fachada del modelo y no implementa el inicio de sesión de Microsoft.

La [referencia de OneDrive de Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#word-docx)
contiene la matriz de paquetes, la configuración de autenticación, los ámbitos de compartición y la
recuperación de conflictos. La [guía de pruebas](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)
separa los transportes locales de prueba de la aceptación en vivo con Microsoft. No se incluye ninguna
interfaz ni servicio de simulación.

Registre `mergeDocxSnapshots` de `ooxml-core/docx` como callback `merge` de la sesión para combinar ediciones
independientes de texto y de propiedades durante el guardado. La suite lo activa automáticamente. Las
estructuras ambiguas o las sustituciones incompatibles conservan ambas versiones para su recuperación. Esto no
conecta las salas de colaboración de Word con la coautoría de Microsoft Office. Los [límites de conservación](/es/preservation)
existentes siguen aplicándose.

## Cuenta > Servicios conectados {#account-connected-services}

Las mismas `fileProviders` también completan la página **Cuenta** vista Archivo con una lista de "Servicios
conectados" (el `<office-ui-connected-services>` compartido): una fila por proveedor, con la cuenta con la que
se ha iniciado sesión y los botones Conectar o Desconectar, que llaman a `signIn` y `signOut` propios del
proveedor. Un proveedor sin `signOut` muestra que el anfitrión gestiona su sesión. La lista sigue el `subscribe`
de los proveedores, de modo que iniciar sesión desde Archivo > Abrir la actualiza, y sus textos proceden de los
diccionarios de Word en todos los idiomas. Sin `fileProviders` no hay página Cuenta.

## Abrir un archivo de la nube desde su propia interfaz {#open-a-cloud-file-from-your-own-ui}

Archivo > Abrir solo es accesible una vez abierto un documento. Para abrir un archivo desde una página de inicio
o su propio selector, obtenga un `CloudOpenedFile` de cualquier `CloudFileProvider` (el explorador compartido
`<office-ui-cloud-files>` lo emite en `office-cloud-file-open`) y páselo a `editor.openCloudFile(file)` en
`<docx-editor>` (el manejador de cada binding también lo expone). Sigue la misma ruta que Archivo > Abrir >
OneDrive: el nombre del archivo, solo lectura cuando `canEdit` es false, y Guardar dirigido al almacenamiento
propio del archivo hasta que se cargue otro documento.
