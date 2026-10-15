# Protocolo de colaboración

El componente web admite dos modos de colaboración excluyentes: pasos de ProseMirror ordenados por una autoridad, y la fusión opcional con Yjs a través de una `CollabSession` propiedad de la aplicación. Ambos usan el mismo editor y los seis adaptadores de framework. El host se encarga del transporte, la autenticación, la persistencia y el momento de la exportación.

## Demo en vivo: dos ventanas, cualquier framework

Una sesión es simplemente un nombre. Los dos paneles de abajo son aplicaciones de demo independientes, compiladas con adaptadores de framework distintos (React y Vue), que se unen a la sesión `docs-collab`: el primer panel abre el documento de ejemplo y aloja la sesión; el segundo se une por nombre y recibe el documento, cada edición y el cursor de la otra persona. Escriba en cualquiera de los dos paneles.

<iframe
	src="/ooxml/docx/demo/?sample=1&room=docs-collab&name=Ada"
	title="Anfitrión de la sesión de docx-viewer (React)"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/docx/demo-vue/?room=docs-collab&name=Grace"
	title="Invitado de la sesión de docx-viewer (Vue)"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Combine dos frameworks cualesquiera

Todas las demos aceptan `?room=<session>`. La ventana que además pide `sample=1` aloja la sesión; cualquier otra ventana con el mismo `room` se une a ella, con el framework con el que se haya compilado. La [demo en vivo](/#live-demo) de la página de inicio tiene un selector de framework para cada ventana, y en dos pestañas puede abrir, por ejemplo, `/demo-angular/?sample=1&room=my-room` y `/demo-svelte/?room=my-room`.

### Cómo lo hace la demo y por qué solo funciona en un mismo navegador

El protocolo del editor es independiente del transporte, así que la demo aporta el transporte: un `BroadcastChannel` con el nombre de la sesión. La ventana anfitriona ejecuta la autoridad de referencia en memoria (`createCollaborationAuthority`), ordena los lotes de pasos de todos y difunde los aceptados; un invitado parte del documento inicial del anfitrión y reproduce los lotes aceptados, por lo que el framework de cada ventana no importa. Un `BroadcastChannel` conecta ventanas, pestañas y marcos de un mismo perfil de navegador y nada más, la sesión termina cuando se cierra su ventana anfitriona y no se persiste nada. Para personas en máquinas distintas, su aplicación debe proporcionar el transporte, la autoridad, la identidad y el almacenamiento (vea más abajo).

También hay una [demo de una sola página](/demo/collaboration.html){target="_self"} con dos editores en la misma página y un botón **Pause delivery** para probar ediciones concurrentes; añada `?guest=vue` para ejecutar su segundo editor en otro framework.

## Qué se admite y qué no

| Admitido                                                                                     | No admitido                                                                            |
| -------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| Lotes de pasos de ProseMirror ordenados y validados a través de una autoridad                | Un servidor alojado, la red, la autenticación o la persistencia (corresponden al host) |
| Fusión de texto y formato con Yjs, cursores relativos, hilos de comentarios y deshacer local | Edición concurrente de estilos, definiciones de numeración y notas fuera del cuerpo    |
| Una autoridad de referencia en memoria para pruebas y la demo local                          | Comandos estructurales de tablas durante la colaboración                               |

## Modo Yjs

Cargue el mismo paquete DOCX de origen en cada editor, incluidos sus medios, estilos,
comentarios, notas y partes opacas. El host proporciona un `documentId` estable para
ese paquete y designa exactamente a un creador de la sala. Este identificador es un
contrato del host, no un token de autenticación ni una suma de comprobación automática
del paquete.

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({
	roomId,
	provider: transportProvider({ transport }), // Application-owned transport
	user: { name: 'Ada', role: 'collaborator' },
});
if (!session.synced) {
	await new Promise<void>((resolve) => {
		const off = session.on('synced', (synced) => {
			if (synced) {
				off();
				resolve();
			}
		});
	});
}
editor.startYjsCollaboration(session, {
	documentId: sourcePackageId,
	initializeIfEmpty: designatedCreator,
});
editor.publishPresence({ name: 'Ada', color: '#2563eb' });
```

Espere a la sincronización inicial real con el proveedor, en lugar de al periodo de
gracia de la puerta de escritura, antes de unirse. El creador inicializa una sala
vacía; los demás participantes adoptan el cuerpo existente y los atributos de página y
sección en lugar de sobrescribirlos con una instantánea local. Un ID de documento o un
formato de sala que no coincidan rechazan la unión. Una sesión de autoridad y una sesión
Yjs no pueden controlar un mismo editor al mismo tiempo.

El esquema de Word compartido actual inicializa salas `word-yjs-v3`. Cada propiedad
directa modelada de una imagen, referencia de nota, marcador de campo o salto tiene un
atributo independiente sobre su base de propiedades importada. Los cambios concurrentes
en propiedades distintas se fusionan, incluida la eliminación de propiedades; el
deshacer local conserva los cambios del otro autor. Los conflictos sobre una misma
propiedad siguen el orden de Yjs. El historial de formato sigue teniendo una instantánea
y un autor previos por fragmento; la atribución completa de revisiones concurrentes de
varios autores sigue sin terminar.

Esto cambia la compatibilidad de las salas: el visor actualizado rechaza las salas v1 y
v2, y los clientes antiguos rechazan las salas v3. Los consumidores del núcleo que usan
esquemas heredados todavía pueden usar salas v1/v2 equivalentes, solo con creación de
comentarios de texto. Para actualizar una sala persistida, exporte su DOCX actual con el
cliente antiguo correspondiente, cargue ese paquete en los clientes actualizados y deje
que el creador designado inicialice una sala v3 nueva y vacía. Mantenga a todos los
participantes en el esquema correspondiente. El adaptador no reescribe una sala en
directo ni migra automáticamente la persistencia del proveedor.

El modo Yjs usa el enlace estable de Yjs 13, `y-prosemirror` 1.3.7. Los atributos del
documento raíz tienen una asignación independiente porque el enlace no los sincroniza.
El texto, el formato y las marcas de revisión con seguimiento viajan por el fragmento
del cuerpo. Las partes de imagen PNG, JPEG, GIF, BMP y SVG recién insertadas usan el
enrutamiento de recursos compartido del núcleo con cargas binarias sin procesar; los
nombres son únicos por cliente y las partes son inmutables. Los bytes de las imágenes
siguen disponibles para deshacer/rehacer y para las exportaciones tras detener la
colaboración. Los recursos del paquete existente permanecen en el origen cargado que
corresponde a cada editor. El host sigue controlando la sincronización de la
exportación y la persistencia de la sala; esto no coordina guardados simultáneos ni
demuestra una exportación sin pérdidas.

El ajuste de registro de Control de cambios también viaja mediante las transacciones del
documento. Al alternarlo se actualiza a los demás participantes, participa en el
deshacer/rehacer local y se exporta como `w:trackRevisions`. Al unirse a una sala
compatible que carece de este ajuste se conserva el valor del paquete cargado hasta que
un participante lo cambie. Las marcas de revisión conservan el nombre del autor de la
edición y la marca de tiempo UTC de la transacción que la registró. Los demás
participantes conservan esos valores en lugar de reatribuir las ediciones remotas; los
nombres y relojes aportados por el cliente no autentican la autoría ni la hora.

Ctrl/Cmd+Z y el deshacer de la cinta usan el historial local de Yjs. La edición
ordinaria también mantiene los registros y anclas de los hilos en transacciones que se
pueden deshacer; las instantáneas locales de hilos se excluyen de los atributos del
documento Yjs. Tras detener Yjs, sus hilos actuales pasan a ser la instantánea del
historial ordinario. El historial sobrevive al desacoplado y al remontaje del editor;
las vistas desconectadas adoptan los cambios al remontarse. La conciencia (awareness)
usa posiciones relativas para asignar el cursor y la selección, con la interfaz de
cursores ya localizada. `publishPresence` actualiza el perfil de awareness y devuelve
`null` en modo Yjs; `leavePresence` borra el cursor. El proveedor transporta la
awareness directamente, por lo que `presence-send` y `receivePresence` siguen siendo
API del modo de autoridad.

`reconnectCollaboration()` reinicia el proveedor de Yjs conservando el documento
compartido. `resyncCollaboration()` solicita un intercambio de estado y devuelve
`false` cuando el proveedor no puede hacerlo. El host también puede usar
`session.connect()`, `disconnect()`, `peers()`, las suscripciones de estado,
sincronización y error, y `destroy()`. `stopCollaboration()` desacopla el enlace de Word
y conserva su contenido actual y los nuevos recursos de imagen, también mientras está
desmontado; no destruye la sesión que pertenece al host. Destruya esa sesión cuando el
host abandone la sala.

Los editores de solo lectura y los roles de visor reciben las actualizaciones remotas
pero bloquean las escrituras locales en el documento. Los roles de cliente son
orientativos: aplique también la autorización en el servidor o en el proveedor. La
edición fuera del cuerpo y los comandos estructurales de tablas siguen desactivados. Los
atributos de página y de sección se conservan y sincronizan, pero la colaboración en
varias historias (stories) y un arranque completo del paquete siguen sin terminar.

Las salas Yjs nuevas también comparten anclas de comentarios, respuestas, resolución y
eliminación a través del panel de revisión existente. El creador inicializa los
comentarios cargados. Cada comentario tiene atributos de texto independientes, rangos
relativos al elemento y un registro; las respuestas concurrentes conservan sus propios
ID. La resolución usa el orden de conflictos de los mapas de Yjs. Eliminar un comentario
raíz oculta sus respuestas, incluidas las escritas concurrentemente sin conexión.
Deshacer restaura el comentario raíz junto con su ancla y muestra las respuestas que
sobrevivieron; el deshacer de cada autor conserva las operaciones de los demás. La
exportación toma una instantánea de los últimos hilos compartidos incluso con el editor
desacoplado.

Las salas V3 llevan una capacidad de comentarios `inline-relative-v1` en su mapa de
identidades de Word. Las imágenes, saltos, referencias de nota y marcadores de campo
conservan anclas independientes ante adiciones y eliminaciones concurrentes, ediciones
del texto contiguo y deshacer local. Las anclas importadas inicializan estos rangos
cuando se crea una sala nueva. Comentar el resultado o el marcador de un campo complejo
equilibrado abarca el campo completo, igual que el comportamiento nativo de Word
registrado. Las salas heredadas equivalentes conservan su capacidad `independent-v1`
solo de texto.
Las salas compatibles que no la tienen conservan sus comentarios cargados y desactivan
la edición de comentarios. Cree una sala nueva a partir de una instantánea guardada
canónica para actualizar; no cambie la capacidad en una sala en directo. El modo de
pasos con autoridad no sincroniza los metadatos de los comentarios, por lo que su
edición de comentarios permanece desactivada. Estos comportamientos del cliente no
proporcionan autorización en el servidor, persistencia duradera ni las menciones,
notificaciones y asignación de tareas de los comentarios modernos de Word.

Pruebe la [demo Yjs de dos pares](/demo/collaboration.html?mode=yjs){target="_self"}.
Sus acciones Pause delivery, Reconnect providers y Resync providers ejercitan la
recuperación en memoria. No ofrece red de producción ni almacenamiento duradero.

## Protocolo

El componente web expone `startCollaboration`, `getPendingCollaboration`, `receiveCollaboration` y `stopCollaboration`. Los hosts pueden usar `createCollaborationAuthority(model, options)` para crear la autoridad de referencia en memoria con el mismo esquema y la misma conversión de modelo que el editor.

## Flujo del cliente y del host

Todos los participantes que editan el mismo documento usan el mismo `sessionId`; cada encarnación del editor usa un `clientId` único. Reutilizar un ID de cliente puede hacer que los demás confundan los pasos de otro con confirmaciones. Todos los participantes deben usar el mismo esquema y el mismo documento inicial. Para unirse a una autoridad existente, inicialice el cliente con su versión actual y su documento canónico:

```ts
const client = new CollaborationClient({ sessionId, clientId, version: currentVersion });
const state = EditorState.create({ doc: canonicalDoc, plugins: [history(), client.plugin] });
```

Para el elemento personalizado, llame a `editor.startCollaboration({ sessionId, clientId, version })` después de cargar la instantánea del modelo correspondiente. Mantenga montado cada receptor mientras participe: `receiveCollaboration` se aplica a través de la vista del editor en directo. Escuche `collaboration-send`; envíe cada `StepBatch` a la autoridad y entregue el lote aceptado a cada editor con `receiveCollaboration(batch)`. Los rechazos como `stale`, `out-of-order` o `wrong-session` deben activar una recuperación a nivel del host en lugar de descartarse en silencio.

Tras las transacciones locales, llame a `client.createPendingBatch(state)`. Devuelve `null` cuando no hay pasos sin confirmar y, en caso contrario, un sobre inmutable. Siga reintentando exactamente ese sobre hasta que la autoridad lo acepte; no inicie otra solicitud mientras esté en curso. El método devuelve el mismo sobre hasta que se reciba su lote aceptado correspondiente.

El host envía el sobre a su autoridad y retransmite el `batch` aceptado a todos los participantes, incluido el remitente. La `version` de un lote aceptado es la versión de la autoridad inmediatamente anterior a los pasos del lote. Los participantes pasan el sobre a `client.receive(state, batch)`. Con `applied`, despache la transacción devuelta por la vía habitual de actualización de estado del editor. Así las actualizaciones ordinarias del modelo, los eventos de estado y el comportamiento de guardado también se ejecutan con los cambios remotos. Aplicar lotes remotos no consulta el ajuste de solo lectura de la vista; ese ajuste solo bloquea las ediciones locales.

El protocolo acepta como máximo 500 pasos y 1.000.000 de caracteres JSON serializados por lote. La `CollaborationAuthority` de referencia es una implementación en memoria de ordenación y rebase. Acepta solicitudes de la versión actual y asigna las solicitudes obsoletas a través de los pasos aceptados intermedios. Su historial predeterminado conserva 2000 pasos y almacena en caché 5000 confirmaciones de solicitudes para los reintentos. Una solicitud obsoleta anterior al historial retenido se rechaza; una versión futura se rechaza por estar fuera de orden. El cliente recuerda los últimos 1000 ID de lotes recibidos para ignorar duplicados. Los hosts deben ofrecer una puesta al día o recargar desde su instantánea canónica tras un hueco, una versión caducada o un lote sobredimensionado rechazado.

## Presencia

`PresenceClient` ofrece un canal transitorio independiente para los nombres, colores, cursores y selecciones de los participantes. No usa ni modifica el flujo de pasos del documento. Añada su `.plugin` junto al complemento de colaboración. `publish(state, { name, color })` devuelve un mensaje para el transporte del host, o `null` mientras los pasos locales del documento sigan pendientes de confirmación. `leave(state)` devuelve un mensaje de retirada. Tras recibir un mensaje, despache la transacción de `receive(state, message)` cuando su estado sea `applied`; ignore `duplicate` y recupere el estado ante `stale`, `out-of-order` o `wrong-session` según la política del host.

Los mensajes de presencia llevan el ID de sesión compartido, el ID de cliente del remitente, la versión de colaboración confirmada y una secuencia monótonamente creciente por remitente. Retransmita los mensajes en orden y vuelva a publicar tras la confirmación de una edición; las posiciones se definen respecto a la versión confirmada del mensaje. El complemento asigna los cursores y selecciones existentes a través de las transacciones posteriores del documento, incluidas las ediciones recibidas mientras el editor está en solo lectura. Acepta nombres de hasta 80 caracteres imprimibles, colores de `PRESENCE_PALETTE` y como máximo 100 participantes visibles. Los hosts deben limitar la frecuencia y hacer caducar a los participantes desconectados en su capa de transporte. La presencia no se persiste, no se autentica y no se incluye en la exportación a DOCX.

## Identidades estables

Para el elemento personalizado compartido, la presencia se activa con `startCollaboration`:

```ts
import { PRESENCE_PALETTE } from 'docx-vanilla-viewer';

editor.addEventListener('presence-send', (event) => {
	transport.sendPresence(event.detail); // Application-owned transport
});
editor.publishPresence({ name: 'Ada', color: PRESENCE_PALETTE[0] });
// Incoming messages: editor.receivePresence(message)
// Explicit departure: editor.leavePresence()
```

El elemento vuelve a publicar su perfil almacenado tras los cambios de selección y las
confirmaciones de edición. Devuelve `out-of-order` para las posiciones de selección
entrantes mientras haya ediciones locales pendientes; el host puede solicitar presencia
actualizada tras la sincronización. Los mensajes de salida no contienen posiciones y
pueden retirar a un participante en una versión distinta del documento, aplicando aun así
las comprobaciones de sesión y de secuencia del remitente.

Los ID de párrafos y tablas son contenido del documento y viajan en los pasos de ProseMirror. Para las transacciones locales, use `repairCollaborativeDocumentIds(state, clientId, idGenerator)` en lugar del asistente de reparación de ID no colaborativo. Cree `idGenerator` una sola vez con `createCollaborationIdGenerator(clientId)` y reutilícelo durante toda la vida de ese editor; así se evita reutilizar ID tras una eliminación y se acotan por espacio de nombres los ID de tablas, celdas y párrafos insertados localmente. El host debe asignar un ID de cliente distinto a cada encarnación activa del editor; el ID de sesión y el ID de cliente son identificadores del protocolo, no credenciales.

`stopCollaboration()` exige que las ediciones pendientes estén confirmadas. Si el host las abandona a propósito, llame a `stopCollaboration(true)` para descartar explícitamente la cola de transporte local. La desconexión o el desmontaje no deben tratarse como una confirmación; conserve la instantánea y la versión canónicas, y reanude o recupere según la política de transporte del host.

## Responsabilidades y límites de la autoridad

El host de la aplicación se encarga del transporte, la autenticación y la autorización, el almacenamiento duradero, la distribución de instantáneas, la identidad de los clientes, la política de reconexión y el momento de la exportación. La autoridad en memoria es útil para pruebas y para una demo local, no como servidor multiproceso ni de producción. La exportación a DOCX debe usar un modelo canónico suficientemente sincronizado si la aplicación necesita una semántica de guardado compartido; esta biblioteca no coordina bloqueos de archivo ni exportaciones simultáneas. Los participantes remotos solo reciben pasos de ProseMirror admitidos por el esquema, y la autoridad valida esos pasos antes de aceptarlos. Los comandos estructurales de tablas están desactivados durante la colaboración porque el editor de tablas actual reemplaza la estructura de la tabla en bloque; harán falta transacciones granulares a nivel de celda para activar esos comandos con seguridad.

Los vínculos admitidos de imágenes en línea, saltos y referencias de nota se transportan
mediante atributos de fragmento, por lo que importarlos o editarlos no depende de marcas
de elemento que Yjs descarta. El cuadro de diálogo Vínculo lee y edita esos destinos
mediante los mismos comandos del núcleo que los vínculos de texto ordinarios. El cambio
de destino y la eliminación se sincronizan y admiten deshacer local por autor; los
medios de imagen compartidos siguen disponibles en las exportaciones a DOCX. Los
vínculos de ecuaciones y la atribución de revisiones de hipervínculos quedan fuera de
esta compatibilidad.
