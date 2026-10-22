# API del elemento y eventos

Todos los bindings envuelven el mismo elemento personalizado `<teams-app>`. Los bindings reenvían estas propiedades y vuelven a emitir estos eventos según el idioma de su framework (consulta las [guías por framework](/frameworks/react)).

## Atributos y propiedades

| Propiedad          | Atributo        | Descripción                                                                                                                 |
| ------------------ | --------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `workspaceId`      | `workspace-id`  | La sala compartida: de 1 a 128 caracteres entre letras, dígitos, `-` y `_`.                                                 |
| `userName`         | `user-name`     | Nombre para mostrar.                                                                                                        |
| `userId`           | `user-id`       | Id estable; por defecto usa uno recordado en este navegador.                                                                |
| `config`           | `server-config` | `{ mode: 'local' \| 'server', syncUrl?, signalingUrl?, iceServers, iceTransportPolicy?, token? }`. El atributo acepta JSON. |
| `uploadFile`       |                 | Adónde van los adjuntos; por defecto, el endpoint `/files` del servidor.                                                    |
| `openers`          |                 | Manejadores por tipo de archivo de Office (`docx`, `xlsx`, `pptx`, `vsdx`).                                                 |
| `embeds`           |                 | Callbacks por tipo de archivo que devuelven la URL de una página de visor compatible con el sandbox; útil para PowerPoint.  |
| `client` (lectura) |                 | El cliente del núcleo que hay detrás del elemento (el estado entra, las acciones salen).                                    |

Si omites `config`, el elemento ofrece un cuadro de diálogo de ajustes (recordado en este navegador) o funciona en modo local.

## Vistas previas de contenido

Al abrir un adjunto se muestra por defecto una vista previa del espacio de trabajo. Word y Excel usan los elementos de editor nativos, inicialmente en solo lectura; Visio usa su visor. Markdown tiene un renderizador de un subconjunto seguro; los archivos de texto muestran el texto literal; HTML y los sitios web usan marcos en sandbox. `openers` y el evento cancelable `teams-open-file` siguen teniendo prioridad.

PowerPoint requiere una página de incrustación proporcionada por el anfitrión:

```ts
const embeds = {
	pptx: ({ url }: { url: string | undefined }) =>
		`https://viewer.example.com/embed?source=${encodeURIComponent(url ?? '')}`,
};
```

Pasa `embeds` a cualquier binding de componente. La página debe entender la URL del archivo y admitir un sandbox con scripts y formularios pero sin permiso de mismo origen. Las descargas de archivos omiten las credenciales; los servidores de archivos de origen cruzado deben permitir CORS. Sin un visor de PowerPoint configurado, el panel indica esa limitación y ofrece la apertura externa.

El método `previewContent({ attachment, url })` del elemento abre localmente contenido proporcionado por el anfitrión. Para un sitio web usa `attachment: { name: 'Project site', kind:
'other', mime: 'text/html' }`. Esto no crea una pestaña de canal compartida. La vista de archivos también ofrece **Preview website**. Los sitios que rechazan ser enmarcados pueden abrirse externamente.

## Pestañas de canal compartidas y copias de Excel

Usa **Pin as tab** junto a un archivo subido, o **Add tab** para añadir un sitio web con nombre. Las definiciones se comparten; cada usuario selecciona su propia pestaña. El creador puede renombrarla o quitarla mediante comprobaciones de propiedad en el cliente. Quitar una pestaña en otro lugar mantiene disponible un panel local ya abierto hasta que se cierre.

Excel ofrece edición local y **Save copy to channel** cuando hay almacenamiento de archivos configurado. Cada guardado sube una copia con un nombre único y la publica en el canal de origen, conservando el original. Los guardados fallidos conservan las ediciones; la navegación pide confirmar el descarte del trabajo sin guardar y espera a que termine un guardado en curso. Esto no es coedición ni escritura con control de versiones. Los anfitriones que reemplacen a la fuerza el espacio de trabajo o el elemento deben gestionar su propia confirmación de trabajo sin guardar.

Acciones sin procesar del cliente:

```ts
const tab = teams.addTab('Budget', { type: 'file', attachment });
teams.addTab('Project site', { type: 'website', url: 'https://example.com' });
teams.renameTab(tab!.id, 'Budget review');
teams.removeTab(tab!.id);
await teams.saveFileCopy(channelId, file);
```

`TeamsState.tabs` contiene las definiciones del canal seleccionado; `canUploadFiles` indica que hay almacenamiento configurado. `OpenFileDetail.channelId` registra el canal capturado al abrir. `previewContent` también acepta este campo cuando un anfitrión quiere habilitar las copias guardadas en el canal. Las comprobaciones de propiedad requieren autorización en el servidor antes de poder tratarse como control de acceso.

## Creación, subida y búsqueda de archivos

Ambas vistas de archivos ofrecen subida, creación de un libro de Excel en blanco y búsqueda. Las subidas y los libros nuevos tienen como destino el canal seleccionado, que aparece nombrado junto a los controles. Necesitan almacenamiento configurado; los fallos pueden reintentarse en el canal original. Los archivos con el mismo nombre usan URL de almacenamiento separadas. Las subidas aceptan hasta 20 archivos de 32 MiB cada uno (el límite propio del servidor de referencia es de 25 MB por archivo salvo que aumente `maxFileBytes`); la interfaz informa de la actividad pero no ofrece progreso en bytes ni cancelación.

```ts
await teams.uploadFiles(channelId, files);
const workbook = await teams.createWorkbook(channelId, 'Budget'); // Budget.xlsx
```

La creación publica un adjunto respaldado por un archivo XLSX real. La búsqueda filtra los metadatos cargados por nombre de archivo, autor, canal y tipo reconocido. No busca dentro del contenido de los archivos. Estas acciones capturan el canal explícito antes de la creación o subida asíncrona, de modo que seleccionar otro canal no puede redirigir el resultado.

## Eventos

| Evento                | Significado                                                                                     |
| --------------------- | ----------------------------------------------------------------------------------------------- |
| `teams-ready`         | El elemento está listo.                                                                         |
| `teams-open-file`     | Un usuario abrió un archivo adjunto. Cancelable: `preventDefault()` para abrirlo por tu cuenta. |
| `teams-config-change` | El usuario guardó una nueva configuración del servidor en el cuadro de diálogo de ajustes.      |

Grafías en los bindings: `onReady`, `onOpenFile`, `onConfigChange` (React, Solid, Svelte, vanilla); `@ready`, `@open-file`, `@config-change` (Vue); `(ready)`, `(openFile)`, `(configChange)` (Angular).

## El cliente sin procesar

`createTeams(options)` (el hook sin procesar de cada binding) devuelve el cliente del núcleo:

```ts
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
teams.on('notice', (notice) => show(notice)); // errors a user should see
```

`getState()` es un único `TeamsState` inmutable: `status` de la conexión, `channels` con contadores de no leídos, los `messages` y `files` del canal seleccionado, `people` (presencia), `typing`, el estado del compositor, los resultados de búsqueda y `call`. Las acciones incluyen `select`, `createChannel`, `send`, `startReply`, `startEdit`, `deleteMessage`, `toggleReaction`, `notifyTyping`, `setAvailability`, `search`, `openCall`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`, `toggleScreenShare`, `toggleHand` y `destroy`. Consulta [Arquitectura](/architecture) para ver cómo encajan las capas.
