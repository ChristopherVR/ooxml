---
title: Integración con OneDrive
description: Conecte cada enlace del visualizador de PowerPoint al almacenamiento compartido de Office en OneDrive, con uso compartido y recuperación de conflictos.
---

# Integración con OneDrive

Cada paquete de framework de PowerPoint también acepta un proveedor de almacenamiento personalizado mediante su prop `storage` existente. Consulte [integración independiente del proveedor y configuración](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md) para `createStorageSession`, las anulaciones de autenticación y transporte, los manejadores de fusión por formato, los diálogos de conflicto del host y las opciones explícitas de uso compartido.

PowerPoint usa la misma conexión opcional `ooxml-core/embedding/onedrive` que Word, Excel y Visio. El host aporta la autenticación de Microsoft y los permisos verificados de cada elemento. No hay registro de aplicación de Microsoft incluido, ni servicio de simulación, ni demostración simulada de OneDrive.

Para una configuración paso a paso con capturas (registro de la aplicación, inicio de sesión, un cuadro de enlace para compartir y un ejemplo completo en React), empiece por el [inicio rápido de OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md). Cada demostración de framework ofrece OneDrive en Archivo > Abrir para probarlo con su propio OneDrive; la sección "Probarlo en las demostraciones de PowerPoint" del inicio rápido explica la configuración.

## Lugares de Archivo > Abrir: `fileProviders` {#file-open-places-fileproviders}

Pase `fileProviders` (una lista de `CloudFileProvider`, por ejemplo desde `createOneDriveFileProvider` en `ooxml-core/embedding/onedrive`) y la página Abrir muestra **Este dispositivo** más un lugar por cada proveedor. La prop se escribe igual en todos los enlaces: `fileProviders` (React, entrada de Angular, Svelte, opción de vanilla) y `:file-providers` (Vue).

```tsx
<PowerPointViewer content={bytes} canEdit fileProviders={[oneDriveProvider]} />
```

Para combinar las ediciones que otra persona guardó entretanto, dé al proveedor un fusionador de PowerPoint. Sin uno, una versión más nueva de OneDrive siempre abre las opciones de conflicto; `merge: false` pregunta de la misma forma para todos los formatos, y `maxMergeAttempts` (3 por defecto cuando aplica una fusión) limita los reintentos:

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergePptxSnapshots } from 'ooxml-core/pptx';

const oneDriveProvider = createOneDriveFileProvider({
	getAccessToken,
	canWrite,
	signIn,
	mergers: {
		pptx: ({ base, local, remote }) => mergePptxSnapshots(base.bytes, local, remote.bytes),
	},
});
```

Un archivo abierto desde un lugar se carga como cualquier apertura (se reinicia el historial de deshacer, se dispara `onLoad`). A partir de ahí, el visualizador usa el almacenamiento propio de ese archivo para Archivo > Guardar, Guardar como y las copias de recuperación, en lugar de la prop `storage`, y lo abre en solo lectura cuando el servicio no permite editarlo; un host con `canEdit={false}` o `readOnly` permanece en solo lectura. Esto dura hasta que se carga otro documento: una presentación nueva, una apertura desde este dispositivo o desde Recientes, o un nuevo `content` del host. Un guardado fallido muestra el aviso de guardado fallido, y la barra de estado mantiene **Cambios sin guardar**.

## Cuenta > Servicios conectados {#account-connected-services}

Los mismos `fileProviders` también rellenan una sección **Servicios conectados** en **Archivo > Cuenta**, debajo de las tarjetas de perfil, almacenamiento e inicio de sesión. Cada proveedor es una fila con su cuenta (nombre y correo) y **Conectar** o **Desconectar**, que llaman a los métodos `signIn` y `signOut` del propio proveedor; un proveedor sin `signOut` muestra que su sesión la gestiona su organización. La lista sigue el `subscribe` del proveedor, de modo que iniciar sesión desde Archivo > Abrir la actualiza. Sin `fileProviders` no aparece nada adicional, y la tarjeta de inicio de sesión existente `accountAuth` no cambia. Todos los enlaces muestran el elemento compartido `office-ui-connected-services` (como `pptx-ui-connected-services`) con etiquetas traducidas.

La [guía completa de integración con Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) documenta la autenticación, el almacenamiento de sesión, las copias seguras en la nube, los roles y alcances de uso compartido, la recuperación de conflictos y la referencia de cada tipo de paquete.

```ts
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';

const drive = createOneDriveClient({
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
});
const opened = await drive.open(sharingLink);
const session = createOneDriveSession(drive, opened);
// Pass session.storage and opened.file.access === 'write' to the viewer.
```

La sesión y los servicios de permisos de este ejemplo son integraciones del host, no objetos incluidos. Importe el adaptador como dependencia directa de `ooxml-core`. `ooxml-ui/pptx/onedrive` sigue existiendo solo como superficie de compatibilidad para la importación anterior específica de PowerPoint.

## Los cinco enlaces {#all-five-bindings}

| Enlace/paquete                 | Abrir bytes                           | Restringir edición            | Exportar el PPTX actual           |
| ------------------------------ | ------------------------------------- | ----------------------------- | --------------------------------- |
| React, `pptx-react-viewer`     | `content={opened.bytes}`              | `canEdit={canEdit}`           | `await ref.current.getContent()`  |
| Vue, `pptx-vue-viewer`         | `:content="opened.bytes"`             | `:can-edit="canEdit"`         | `await viewer.value.getContent()` |
| Angular, `pptx-angular-viewer` | `[content]="opened.bytes"`            | `[canEdit]="canEdit"`         | `await viewer.getContent()`       |
| Svelte, `pptx-svelte-viewer`   | `source={opened.bytes}`               | `editable={canEdit}`          | `await viewer.getContent()`       |
| Vanilla, `pptx-vanilla-viewer` | `await viewer.loadFile(opened.bytes)` | `viewer.setEditable(canEdit)` | `await viewer.save()`             |

`canEdit` (`editable` en Svelte) significa `opened.file.access === 'write'`. Un acceso desconocido sigue siendo de solo lectura. Cada enlace acepta `storage` como su prop, entrada u opción existente. Asigne `session.storage` para conectar el Guardar nativo de File a OneDrive. `onOpenFile` puede abrir el selector de enlaces para compartir del host; no pase la URL de la página de uso compartido de Microsoft a un cargador de fuentes de documentos genérico.

El visualizador marca la presentación como guardada solo después de que `storage.save` se resuelve, y solo cuando no se editó nada mientras la subida estaba en curso; un guardado rechazado la mantiene sin guardar. Para un botón Guardar gestionado por el host, exporte el PPTX actual y llame a `session.save(bytes)`. Mantenga el editor montado y conserve el historial de deshacer. Haga un seguimiento de las ediciones hechas durante la subida por separado de la instantánea subida. No dé por bueno que un método de guardado nativo se resuelva como prueba de éxito: algunos métodos informan de los fallos mediante `onError` y la pila de avisos. Consulte [almacenamiento gestionado por el host](/es/guide/host-owned-storage).

## Conflictos y uso compartido {#conflicts-and-sharing}

Un cambio que solo afecta a la revisión puede reintentarse una vez cuando los bytes descargados de la nube siguen siendo iguales a la base cargada. Un conflicto real de contenido conserva las versiones base, local y remota en `OneDriveConflictError`. Ofrezca **Guardar mis cambios como copia de OneDrive**, **Conservar la copia local y abrir la más reciente**, **Descargar mis cambios** y **Seguir editando**. No sobrescriba ni recargue en silencio. Registre `mergePptxSnapshots` como el callback `merge` de la sesión para combinar automáticamente las ediciones independientes de diapositivas compatibles. La suite lo hace por defecto. Los solapamientos no compatibles siguen necesitando recuperación; la sesión no bloquea el archivo hasta la confirmación final de la subida.

Guardar como a un PPTX nativo compatible crea una copia separada en la nube mediante el almacenamiento de la sesión. Las exportaciones a PPT/PPTM/PPSX u otro formato no deben reemplazar el PPTX original; manténgalas como un flujo de exportación separado y explícito. Una operación de Compartir crea un enlace solo después de que el usuario elija su rol de vista o edición y su alcance. Compartir un enlace de OneDrive no conecta una sala de colaboración de PowerPoint con la coautoría de Microsoft PowerPoint.

Consulte [pruebas de integración](/es/guide/onedrive-testing) y las [limitaciones de formato](/es/guide/limitations) existentes.
