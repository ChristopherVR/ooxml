---
title: Revisión dentro de su aplicación
description: Una receta para la revisión colaborativa de presentaciones en una aplicación anfitriona que gestiona las cuentas, los permisos, la sesión, la persistencia y la clave de IA.
---

# Revisión dentro de su aplicación

La colaboración suele ser más útil dentro de un flujo de trabajo que sus usuarios ya tienen: una mesa de operaciones comerciales, una cola de cumplimiento normativo, un registro de CRM o un tablero de proyectos. Varias personas abren la misma presentación desde ese sistema, hacen pequeños cambios a mano o con el asistente de IA, y el sistema controla quién puede hacer qué y qué versión es la vigente.

Esta receta divide el trabajo según una sola línea:

| Su aplicación gestiona                                           | El visualizador gestiona                                                  |
| ---------------------------------------------------------------- | ------------------------------------------------------------------------- |
| El inicio de sesión, las cuentas y el nombre y color del usuario | El renderizado y la edición de la presentación                            |
| Quién puede abrir, editar o solo observar cada elemento          | El esquema de las diapositivas en el documento Yjs compartido             |
| El documento Yjs, la awareness, el proveedor y el relé           | La presencia (cursores, selección, "seguir a un colaborador")             |
| Cuándo la sesión está lista y cuándo deja de estarlo             | Retener los cambios hasta que usted indique que la sala está lista        |
| Dónde se guardan las versiones y quién las escribe               | Convertir la presentación compartida en bytes PPTX para el propietario    |
| El modelo de IA, su clave y quién puede usarlo                   | El panel del asistente, sus herramientas de presentación y las propuestas |

El ejemplo de aplicación en
[`demos/pptx/demo-host-review`](https://github.com/ChristopherVR/ooxml/tree/main/demos/pptx/demo-host-review)
implementa cada paso que se describe a continuación con el enlace vanilla. Todos los enlaces aceptan las mismas
opciones `collaboration` y `ai`, así que los pasos se leen igual en React, Vue, Angular y Svelte.

## 1. Asigne sus permisos a un rol {#_1-map-your-permissions-to-a-role}

Mantenga su propio modelo de permisos. Derive de él el `role` del visualizador al montar el editor:

| Su permiso (ejemplo)     | `role`           | Notas                                                        |
| ------------------------ | ---------------- | ------------------------------------------------------------ |
| Propietario del elemento | `'owner'`        | Edita y recibe instantáneas de `onWriteBack` para guardarlas |
| Revisor                  | `'collaborator'` | Edita en directo junto con los demás                         |
| Observador               | `'viewer'`       | Solo lectura; pase también `editable: false`                 |
| Sin acceso               | (no lo monte)    | Su API ni siquiera debería revelar que el elemento existe    |

`role: 'viewer'` se aplica **solo en el navegador**. Aplíquelo también en el servidor (paso 3).

## 2. Controle la sesión {#_2-own-the-session}

Cree usted mismo el documento Yjs, la awareness y el proveedor, y páselos como `collaboration.externalSession` (consulte [Colaboración gestionada por el anfitrión](/es/guide/host-owned-collaboration) para ver el contrato completo). El visualizador nunca se conecta, no se reconecta ni los destruye.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const grant = await api.requestGrant(reviewId); // your API: ticket, role, intent, user
const doc = new Y.Doc();
const awareness = new Awareness(doc);
const provider = new WebsocketProvider(`wss://${location.host}/collab`, reviewId, doc, {
	awareness,
	params: { ticket: grant.ticket },
});

const viewer = createPptxViewer(container, {
	source: sourceBytes, // the same source PPTX for every participant
	editable: grant.role !== 'viewer',
	collaboration: {
		roomId: reviewId,
		serverUrl: '', // unused with an external session
		userName: grant.user.name, // from your directory, not typed by the user
		userColor: grant.user.color,
		role: grant.role,
		sessionIntent: grant.intent,
		externalSession: { doc, awareness, getSnapshot, subscribe },
	},
});
```

`getSnapshot()` devuelve `{ status, synced }`. Informe `synced: true` solo cuando el proveedor se haya sincronizado **y** el usuario siga conservando el acceso con el que se montó el editor. Si lo establece en `false`, el editor pasa inmediatamente a solo lectura.

**Quién inicializa la sala.** Una sala vacía la llena el primer participante con `sessionIntent: 'create'`; todos los demás pasan `'join'` y adoptan lo que haya en ella. Deje que su servidor decida, no los clientes: el ejemplo concede `'create'` a un solo escritor cada vez, con una concesión temporal breve y solo mientras la sala siga vacía, de modo que dos personas que abran el elemento en el mismo momento nunca publiquen cada una su presentación inicial.

## 3. Admita y controle los sockets en el servidor {#_3-admit-and-police-sockets-on-the-server}

Un WebSocket no puede llevar la cookie de sesión de su sitio a un relé en otro origen, y un token de larga duración en una URL acaba en los registros. Emita en su lugar un **ticket firmado de corta duración** desde una llamada de API autenticada normal, y haga que el relé confíe solo en lo que dice el ticket:

- **Al actualizar la conexión**: verifique la firma, la caducidad y la sala del ticket. Rechace cualquier otra cosa.
- **En cada mensaje**: una conexión de solo lectura puede pedir el documento (paso 1 de sincronización) y compartir presencia, pero sus mensajes de sincronización paso 2 y de actualización se descartan. Esta es la mitad del servidor de `role: 'viewer'`.
- **Awareness**: permita que cada conexión hable solo en nombre de los identificadores de cliente de awareness que reclamó primero, para que nadie pueda borrar o suplantar el cursor de otra persona.
- **Renovación del ticket**: un ticket solo admite una conexión. Renuévelo antes de que caduque y actualice `provider.params.ticket`, para que una reconexión automática use uno válido.

No tiene que escribir este relé usted mismo. `openteams-server` (el servidor de referencia del visualizador OpenTeams, un servidor Yjs compatible con y-websocket) acepta un hook `authorizeSync` que devuelve `'write'`, `'read'` o un rechazo por socket, ya fija los identificadores de awareness a sus conexiones y expone `closeSync`, `syncPrincipals`, `isSyncRoomEmpty` y `handleUpgrade` para los pasos siguientes. El `server/relay.ts` del ejemplo son unas sesenta líneas alrededor de él.

## 4. Cambie el acceso mientras la gente trabaja {#_4-change-access-while-people-are-working}

Cuando un propietario cambia el acceso de alguien, cierre los sockets abiertos de esa persona con un código de cierre de aplicación (el ejemplo usa `4403`). El cliente reacciona:

1. informando `synced: false`, lo que congela el editor;
2. destruyendo el visualizador y después su propio documento y proveedor;
3. pidiendo una nueva concesión y montando una sesión nueva con el nuevo rol, o mostrando "access removed" si no hay ninguno.

Empiece con un Y.Doc **nuevo** tras una reducción de permisos. Las ediciones que el usuario hizo justo antes del cambio se descartaron en el relé, y reutilizar el documento antiguo seguiría mostrándolas en local.

## 5. Guarde versiones con un solo escritor {#_5-save-versions-with-one-writer}

Monte exactamente un participante, normalmente el propietario del elemento, con `role: 'owner'` y `onWriteBack(bytes)`. El visualizador agrupa los cambios (`writeBackDebounceMs`, 5 segundos por defecto) y le entrega una instantánea PPTX de la presentación compartida para guardarla como versión. Los demás participantes nunca escriben de vuelta, así que no hay una carrera de "gana el último guardado".

`onWriteBack` no es una confirmación de almacenamiento duradero: los reintentos, los fallos y la elección de un escritor cuando el propietario está desconectado siguen siendo decisiones suyas. Si la sala debe sobrevivir a un reinicio del relé, persista también las actualizaciones de Yjs en el servidor (por ejemplo, con su propio almacén o con un backend de proveedor como y-leveldb).

## 6. Añada cambios asistidos por IA {#_6-add-ai-assisted-changes}

Pase la opción `ai` solo a las cuentas que puedan usarla, y apúntela a **su** propia ruta, para que la clave del modelo nunca llegue al navegador:

```ts
ai: {
	connection: {
		kind: 'endpoint',
		api: `/api/ai/chat?review=${reviewId}`,
		credentials: 'same-origin',
	},
	writePolicy: 'stage', // every change is a proposal the reviewer accepts or rejects
},
```

La ruta comprueba la cuenta y el permiso de quien llama, y después ejecuta el modelo con las herramientas de presentación de solo esquema de `ooxml-ui/pptx/ai` (consulte [Asistente de IA](/es/guide/ai-assistant)). Las llamadas a herramientas se ejecutan en el navegador del revisor, sobre la presentación en vivo. Una propuesta aceptada es una edición normal de ese revisor: pasa por el historial del editor (un Ctrl+Z la revierte) y llega a la sala igual que llega lo que escribe. Su registro de auditoría puede anotar que se usó el asistente; el documento compartido no marca qué ediciones proceden de él.

## Ejecute el ejemplo {#run-the-example}

```sh
bun install                                   # at the repository root
cd demos/pptx/demo-host-review
bun run start                                 # http://localhost:4190
```

Abra la página en dos navegadores (o en una ventana normal y otra privada) e inicie sesión con personas distintas: Ana es la propietaria del elemento, Ben y Chen revisan y Dee observa. Como Ana, cambie el acceso de Dee en el panel lateral y observe cómo el editor de Dee pasa entre solo lectura, edición y sin acceso sin recargar la página. Para probar el asistente, establezca `REVIEW_AI_MODEL` (y `REVIEW_AI_API_KEY`, o `REVIEW_AI_BASE_URL` para un servidor local compatible con OpenAI) antes de iniciarlo.

## Límites {#limits}

- El ejemplo guarda todo en memoria: las cuentas, las versiones y las salas compartidas desaparecen al detenerse el proceso. Su inicio de sesión es un selector de cuentas.
- El contenido de la awareness (el nombre y el color de visualización) procede del cliente. El relé fija los identificadores de awareness a las conexiones, pero no comprueba el contenido frente a la cuenta del ticket.
- Los comentarios, un estado de aprobación de la revisión y la atribución de cada cambio no forman parte del modelo compartido en este ejemplo; valide esos flujos por separado antes de depender de ellos.
- Que los navegadores pares converjan no es lo mismo que que PowerPoint vuelva a abrir la versión guardada sin una reparación; compruebe presentaciones representativas en PowerPoint antes de prometerlo.
