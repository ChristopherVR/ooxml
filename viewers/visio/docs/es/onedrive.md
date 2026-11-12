# Integración con OneDrive

Para una configuración paso a paso (registro en Microsoft Entra, inicio de sesión y un ejemplo completo), empiece por el [inicio rápido de OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

Para otro servicio o una implementación propia de OneDrive, consulte [integraciones de almacenamiento independientes del proveedor](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md).
`createStorageSession` se conecta a la misma propiedad `storage` en todos los paquetes de Visio,
con comprobaciones de revisión gestionadas por el anfitrión, manejadores de fusión, interfaz de conflictos y políticas de uso compartido.

## Archivo > Abrir con ubicaciones en la nube {#file-open-with-cloud-places}

Pase `fileProviders` (una propiedad del elemento y una prop de cada enlace) para añadir ubicaciones en la nube a Archivo > Abrir. Cuando la lista no está vacía, Abrir muestra **Este dispositivo** (la baldosa habitual de Examinar) y una ubicación por cada proveedor, que muestra el explorador compartido `<office-ui-cloud-files>` para los dibujos `.vsdx`: iniciar sesión, Recientes, Mis archivos, búsqueda y apertura desde un vínculo de uso compartido. Los textos proceden del idioma de visualización de Visio en cada configuración regional.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDrawingSnapshots } from 'ooxml-core/visio/collab';

const oneDrive = createOneDriveFileProvider({
	getAccessToken, // the host's own sign-in
	signIn, // opens the host's sign-in UI and resolves with the account
	canWrite: async (file) => file.driveId === (await myDriveId()),
	mergers: {
		vsdx: ({ base, local, remote }) => mergeDrawingSnapshots(base.bytes, local, remote.bytes),
	},
});
// <VisioViewer fileProviders={[oneDrive]} /> (React); [fileProviders] in Angular; :file-providers in Vue.
```

Un dibujo elegido allí se carga por la ruta normal de `.vsdx` y toma el nombre del archivo; después:

- Es de solo lectura cuando lo es el archivo (`canEdit` en false). Un anfitrión que ya dejó el visor en solo lectura (`canEdit`, `readOnly`, un rol `viewer` de colaboración o `capabilities.edit`) sigue en solo lectura para un archivo editable: prevalece la más restrictiva de las dos.
- Archivo > Guardar y Guardar como usan el almacenamiento propio de ese archivo en lugar de `storage` del anfitrión, hasta que se cargue otro dibujo (Abrir, Nuevo, `load` o `document`). Guardar se resuelve solo después de que el proveedor haya almacenado los bytes: el dibujo se lee entonces como limpio, y queda sin guardar con un `document-error` cuando falla la subida.
- El evento `file-command` cancelable se dispara primero (`open`, `save`); cancelar `open` deja el dibujo actual en su sitio. La vista de backstage se cierra cuando el dibujo está abierto.
- Los diagnósticos de exportación se informan como en cualquier otro guardado. El `.vsd` heredado no es un destino de edición en la nube.

Las demostraciones pasan una ubicación de OneDrive por defecto, con inicio de sesión mediante el diálogo exclusivo de la demostración (un identificador de cliente, o un token de Graph Explorer pegado en **Token de acceso**). Ese diálogo es configuración de la demostración, no interfaz del producto; consulte `demos/onedrive`. Para omitir el diálogo, defina `VITE_ONEDRIVE_CLIENT_ID` (y opcionalmente `VITE_ONEDRIVE_AUTHORITY`) en `demos/visio/.env.local`, ejecute `bun run dev` en `viewers/visio` y registre la dirección que muestra (con una barra final) como URI de redirección de aplicación de una sola página.

## Cuenta > Servicios conectados {#account-connected-services}

Los mismos `fileProviders` también completan la página **Cuenta** vista de Archivo con una lista de «Servicios conectados» (el `<office-ui-connected-services>` compartido): una fila por proveedor con la cuenta con la sesión iniciada y los botones Conectar o Desconectar, que llaman a `signIn` y `signOut` propios del proveedor. Un proveedor sin `signOut` indica que el anfitrión gestiona su sesión. La lista sigue los `subscribe` de los proveedores, de modo que iniciar sesión desde Archivo > Abrir la actualiza, y sus textos proceden de los diccionarios de Visio en cada idioma. La página también conserva el perfil local y la información del producto. Como en Word y Excel, la página Cuenta aparece solo cuando hay `fileProviders`; sin ellos, el perfil sigue siendo editable en **Archivo > Opciones**.

## Uso directo de una sesión {#using-a-session-directly}

Visio usa la conexión compartida `ooxml-core/embedding/onedrive` para los archivos VSDX nativos. El anfitrión gestiona el inicio de sesión de Microsoft, la resolución de vínculos y la verificación de permisos de cada elemento. Entregue los bytes descargados al visor existente, con una restricción de solo lectura del anfitrión cuando el acceso sea desconocido:

```ts
// Set these through the element or binding's existing props/options.
const canEdit = opened.file.access === 'write';
const options = { storage: session.storage, canEdit, readOnly: !canEdit };
// After mounting with options:
await viewer.load(opened.bytes);
```

`exportVsdx()` devuelve un objeto que contiene los `bytes` nativos y los diagnósticos del escritor. Un botón Guardar propio puede llamar a `await session.save((await viewer.exportVsdx()).bytes)`. El Guardar nativo de Archivo usa el mismo contrato de almacenamiento, salvo que el anfitrión impida el evento `file-command` cancelable. Revise los diagnósticos en lugar de prometer una exportación sin pérdidas. Las vistas previas de `.vsd` binarios no son destinos de edición nativa en OneDrive.

Esto se aplica al elemento `ooxml-ui/visio` y a `visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`, `visio-svelte-viewer`, `visio-solid-viewer` y `visio-vanilla-viewer`. `visio-core` es la fachada del modelo. La conexión no añade autenticación de Microsoft a ella.

Consulte la [referencia de OneDrive de Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#visio-vsdx) para la conexión de paquetes, la autenticación, el uso compartido con control de permisos y la recuperación de conflictos que conserva ambas versiones. Consulte las [pruebas](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) para los transportes locales de prueba y la aceptación con el servicio real. No se incluye ninguna interfaz ni servicio de simulación.

Registre `mergeDrawingSnapshots` de `ooxml-core/visio/collab` como callback `merge` de la sesión para combinar, al guardar, las ediciones independientes de formas y propiedades que se admiten. La suite lo hace automáticamente. Los cambios solapados o ambiguos siguen requiriendo recuperación. Esto no es la coautoría de Microsoft ni un bloqueo de archivo. Siguen aplicándose los [límites de Visio](/parity) existentes y los diagnósticos de exportación.

## Abra un archivo de la nube desde su propia interfaz {#open-a-cloud-file-from-your-own-ui}

Archivo > Abrir solo es accesible cuando ya hay un documento abierto. Para abrir un archivo desde una página de inicio o desde su propio selector, obtenga un `CloudOpenedFile` de cualquier `CloudFileProvider` (el explorador `<office-ui-cloud-files>` compartido lo emite en `office-cloud-file-open`) y páselo a `viewer.openCloudFile(file)` en `<visio-viewer>` (el handle de cada enlace también lo expone). Sigue la misma ruta que Archivo > Abrir > OneDrive: el nombre del archivo, solo lectura cuando `canEdit` es false, y Guardar dirigido al almacenamiento propio del archivo hasta que se cargue otro documento.
