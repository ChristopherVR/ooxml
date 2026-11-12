---
title: Colaboración gestionada por el host
description: Aporte un documento Yjs, la conciencia (awareness) y el estado de sincronización existentes sin dar al visualizador la propiedad del transporte.
---

# Colaboración gestionada por el host

Use `collaboration.externalSession` cuando su aplicación ya gestione la autenticación, un proveedor de Yjs, el almacenamiento sin conexión o una conexión compartida. React, Vue, Angular, Svelte y Vanilla aceptan el mismo contrato. El visualizador gestiona su esquema de diapositivas y la presencia; el host gestiona la conexión, la reconexión, la autorización y la vida útil de los recursos.

```ts
import type { CollaborationConfig, ExternalCollaborationSession } from 'pptx-react-viewer';

// These resources and the store belong to your application, not to a viewer mount.
const session: ExternalCollaborationSession = {
	doc: hostDocument,
	awareness: hostAwareness,
	getSnapshot: () => connectionStore.getSnapshot(),
	subscribe: (listener) => connectionStore.subscribe(listener),
};

const collaboration: CollaborationConfig = {
	roomId: 'presentation-123',
	serverUrl: '', // Required by the shared config, but unused for an external session.
	userName: 'Alice',
	sessionIntent: 'join',
	externalSession: session,
};
```

Importe los tipos desde el paquete de su framework. Pase `collaboration` a la prop o opción habitual del visualizador. El documento debe ser un `Y.Doc` real; la conciencia debe pertenecer a ese documento. Mantenga estable el objeto de sesión y notifique a los suscriptores cuando cambien `status` o `synced`. Una suscripción devuelve una función para cancelarla.

El host y el visualizador deben resolver el mismo runtime de Yjs. El visualizador deja Yjs fuera de su paquete y lo declara como par opcional. Instale `yjs` en la aplicación, tanto para la colaboración integrada como para la gestionada por el host:

```sh
npm install yjs
```

Ver y editar sin colaboración no requiere este par. El host también instala el proveedor y la implementación de conciencia que elija. Deduplique Yjs si su árbol de dependencias contiene varias versiones: los tipos de copias de runtime distintas no pueden mezclarse de forma segura.

## Uso del contrato de colaboración compartido {#using-the-shared-collaboration-contract}

Word, Excel y Visio aceptan una única forma de opción `collaboration` desde `ooxml-core/collab` (`CollaborationOptions`: `roomId`, `serverUrl`, `provider`, `session`, `user`, `intent`). PowerPoint acepta la misma forma en su prop `collaboration`, junto con su propio `CollaborationConfig`, de modo que un host que integre varios productos puede entregarles una sola sesión:

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({ roomId, provider: transportProvider({ transport }), user });
// <PowerPointViewer collaboration={{ session, user: { name: 'Ada' }, intent: 'join' }} />
```

- `session` es prestada: pasa a ser el `externalSession`, y el visualizador nunca la destruye. El host la crea, la conecta y la destruye. `roomId` toma por defecto la sala de la sesión.
- `user.name`, `color`, `avatar` y `role` se asignan a `userName`, `userColor`, `userAvatar` y `role`; `intent` se asigna a `sessionIntent`. Los campos heredados explícitos prevalecen cuando se establecen ambos.
- Cada enlace ejecuta el mismo `resolveCollaborationConfig` (desde `ooxml-ui/pptx`) donde lee la prop, de modo que se comportan de forma idéntica. Un `CollaborationConfig` heredado se devuelve sin cambios.

Sin una `session`, las mismas opciones también cubren los casos en que un host no quiere construir una sesión:

```ts
// Your own transport: the viewer opens and destroys the session around it.
collaboration = { roomId, provider: transportProvider({ transport }), user: { name: 'Ada' } };
// No server: tabs and windows of this browser share over a BroadcastChannel.
collaboration = { roomId, user: { name: 'Ada' } };
```

El transporte se elige en el orden compartido: `session` (prestada, nunca destruida) > `provider` > `serverUrl` > BroadcastChannel. Para `provider` y para la sala de BroadcastChannel, el visualizador llama a `openCollaborationSession` (el runtime de `ooxml-core/collab` solo se importa entonces, compartiendo la única copia de Yjs), presenta la sesión por el mismo camino de sesión prestada que sigue una `session` del host, y la destruye al detener, al desmontar o cuando cambia la configuración. La sesión es editable una vez sincronizada, o tras el periodo de gracia para un par solitario en una malla; `intent: 'join'` sigue esperando la presentación de la sala, y `role: 'viewer'` permanece de solo lectura. El estado informa `connecting` mientras se abre la sesión y `error` si el nombre de la sala no es válido o el navegador no tiene BroadcastChannel.

Una `CollaborationConfig` heredada conserva sus transportes: `serverUrl` selecciona y-websocket, y `transport: 'webrtc'` o un `serverUrl` vacío selecciona y-webrtc sin servidor (que también sincroniza las pestañas del mismo navegador). Solo la entrada del contrato compartido que no indica `serverUrl` ni `transport` usa la sala de BroadcastChannel. La presencia del visualizador y el latido de identidad de la sesión comparten el campo `presence` de la conciencia; el visualizador conserva su carga de diapositiva y cursor entre latidos, de modo que los cursores remotos siguen visibles.

## Informar del estado compartido {#reporting-the-shared-state}

Cada enlace informa del mismo `CollaborationState` que informan Word, Excel y Visio (`active`, `roomId`, `status`, `synced`, `people`, con la persona local marcada como `self`; el tipo se importa de `ooxml-core/collab` y también lo reexporta cada paquete). Se dispara cuando una sesión empieza o se detiene (el estado apagado, `status: 'off'`) y cuando cambian el estado, la marca de sincronizado o las personas, tanto en sesiones gestionadas por el host como en las gestionadas por el visualizador, y nunca repite un estado sin cambios:

| Enlace  | Notificación                                               |
| ------- | ---------------------------------------------------------- |
| React   | prop `onCollaborationChange`                               |
| Vue     | evento `collaboration-change`                              |
| Angular | salida `collaborationChange`                               |
| Svelte  | prop de callback `oncollaborationchange`                   |
| Vanilla | opción `onCollaborationChange` y `getCollaborationState()` |

Vanilla conserva `onCollaborationStatus` y `getCollaborationStatus()` para el estado de conexión sin procesar.

## La preparación es decisión del host {#readiness-is-the-host-s-decision}

`getSnapshot()` devuelve `{ status, synced }`. El estado es `connecting`, `connected`, `disconnected` o `error`. `synced` significa que el documento autoritativo está disponible y que el visualizador puede publicar ediciones. Es independiente del estado de la red:

- Mantenga `synced: false` mientras se carga o se hidrata el documento compartido. Ningún tiempo de espera habilita las escrituras en nombre del host.
- Ponga `synced: true` tras la sincronización inicial. Las diapositivas existentes de la sala se adoptan antes de que pueda publicarse cualquier presentación local de arranque.
- Manténgalo en true durante un periodo sin conexión si las ediciones pueden quedar en cola en el proveedor de Yjs del host.
- Póngalo en false para que el editor sea de solo lectura y suspenda la escritura duradera durante una nueva sincronización. La edición se reanuda cuando se adopta el estado autoritativo de la sala.

El texto aceptado en línea y las vistas previas de gestos se escriben de forma síncrona en el documento prestado, de modo que un cambio posterior de preparación no puede descartar una pulsación final en cola. El proveedor del host puede seguir agrupando el tráfico de red. Una edición de texto activa se cierra cuando se revoca la preparación; al reanudar se adopta la sala autoritativa en lugar de reproducir un borrador local obsoleto.

Use `sessionIntent: 'create'` para el participante que puede sembrar una sala vacía, y `'join'` para los participantes cuya presentación de arranque no debe sembrarla. Un File > Open explícito sigue siendo una sustitución intencionada. Coordine la creación de la sala en el host; esto no es autorización en el servidor ni un protocolo de elección.

## Propiedad y persistencia {#ownership-and-persistence}

El visualizador nunca crea un segundo proveedor, no desconecta el proveedor del host y no destruye el documento o la conciencia suministrados. Al desmontarse, elimina sus escuchas y restaura solo el campo de presencia que todavía le pertenece. Los demás campos de conciencia y las escrituras posteriores del host se conservan. Use un único escritor activo de presencia de diapositiva por instancia de conciencia.

El host puede persistir las actualizaciones de Yjs directamente. Como alternativa, designe a un participante con `role: 'owner'` y `onWriteBack(bytes)` para recibir instantáneas PPTX con antirrebote, usando el mismo PPTX de origen y las mismas opciones de guardado que el editor. Las instantáneas pendientes y en vuelo se cancelan si la sesión termina o el host revoca la preparación. La elección del escritor sigue siendo responsabilidad del host. Devuelva una promesa desde `onWriteBack` para confirmar el almacenamiento: el visualizador informa `saving`, `saved` o `failed` en `collaborationState.writeBack` (solo para propietarios), serializa las escrituras (una en vuelo; los cambios posteriores se agrupan en un único seguimiento) y reintenta un rechazo como `writeBackRetry` (`{ attempts, baseDelayMs?, maxDelayMs? }`) indica; un retorno simple se cuenta como confirmado de inmediato. Una escritura pendiente se vacía en `pagehide` y cuando la página pasa a estado oculto. Los navegadores no esperan el trabajo asíncrono al descargar una página, así que una escritura que necesite un viaje de red todavía puede perderse al cerrar. Un host cuya `write` deba sobrevivir a un cierre debe enviar los bytes con `navigator.sendBeacon(url, bytes)` o con `fetch(url, { method: 'PUT', body: bytes, keepalive: true })`, que el navegador completa después de que la página desaparezca (los cuerpos keepalive están limitados a unos 64 KiB, así que un documento grande puede seguir sin llegar).
Una sesión activa con `role: 'viewer'` hace que el editor sea de solo lectura y bloquea la publicación, tanto para la colaboración integrada como para la gestionada por el host. El servidor debe aplicar el control de acceso de forma independiente.

## Límites de integración {#integration-boundaries}

Esta opción cambia la propiedad de la sesión, no el modelo de datos de colaboración existente. El visualizador es dueño de `pptx:slides` y de sus tipos Yjs anidados. Dé a todos los participantes el mismo PPTX de origen para los recursos del paquete, los temas y los medios. Otros datos de la aplicación pueden ocupar claves de Yjs separadas. No reconcilie de forma independiente instantáneas completas de diapositivas desde el host.

Compruebe los límites de mensajes y de persistencia del host con presentaciones representativas, incluidas la actualización inicial del documento y `Y.encodeStateAsUpdate(doc)` tras editar. El tamaño de un archivo PPTX comprimido no es su tamaño de instantánea Yjs: el esquema compartido también transporta XML analizado, metadatos de formas y grupos y cargas de recursos. Mover las imágenes fuera del documento por sí solo puede no bastar. Esta API de sesión no añade fragmentación de transporte ni cambia los límites del host.

No añade integración con un backend de comentarios, una implementación nueva de deshacer solo local, ni sincronización de cada campo de nivel de presentación. Valide esos flujos por separado antes de exponerlos en un producto colaborativo. Reabrir la presentación en PowerPoint nativo también es una comprobación de fidelidad distinta de la convergencia entre dos pares en el navegador.

## Probar las demostraciones {#trying-the-demos}

Inicie el relé de demostración (`bun run collab`) y una demostración de framework, y luego abra:

```text
/?externalSession=1&room=example&sample=1&server=ws://localhost:1234
```

Un segundo participante usa la misma sala sin `sample=1`. La franja del host expone los controles de preparación y de montaje del editor, conservando el documento y el proveedor. La demostración desactiva BroadcastChannel para que la sincronización use el relé WebSocket.

Añada `&headless=1` para probar una interfaz personalizada en cualquiera de las cinco demostraciones. Estos ejemplos usan los bloques públicos del enlace sin montar un visualizador completo oculto. Su botón Guardar, gestionado por el host, llama a la ruta pública de serialización; el selector de zoom le permite comprobar los cursores remotos y los contornos de selección a distintas escalas.

Consulte [Hooks de React](/react/hooks), [Colaboración en Vue](/vue/collaboration), [Colaboración en Angular](/angular/collaboration), [Colaboración en Svelte](/svelte/collaboration) y [Interfaces personalizadas en Vanilla](/vanilla/api#custom-shell-collaboration) para cada composición.

Para una aplicación completa alrededor de esta opción (cuentas, relé con permisos comprobados, cambios de acceso en vivo, versiones del propietario y una ruta de IA protegida), consulte [Revisión dentro de su aplicación](/es/guide/review-workflow).
