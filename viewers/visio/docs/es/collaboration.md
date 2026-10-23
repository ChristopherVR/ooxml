# Colaboración {#collaboration}

Archivo > Compartir inicia una sesión en vivo entre ventanas de un mismo navegador. Una aplicación anfitriona puede compartir también entre dispositivos mediante la propiedad `collaboration`, que admite las mismas opciones que cualquier otro editor de Office.

::: warning Parcial
En esta compilación no hay ningún servicio alojado. Con Archivo > Compartir nada sale de su navegador. Entre dispositivos usted aporta su propio servidor (una URL compatible con y-websocket), proveedor o sesión de Yjs; las páginas y las formas se combinan elemento a elemento y todo lo demás se comparte entero (véase más abajo).
:::

## Integración: la propiedad `collaboration` {#embedding-the-collaboration-property}

`<visio-viewer>` implementa el contrato de integración compartido de `ooxml-core/collab` (`CollaborationOptions`, `CollaborationState`, `CollaborativeEditor`), el mismo que usan Word, Excel y PowerPoint, de modo que un anfitrión que integra varios productos solo se integra una vez.

```ts
viewer.collaboration = {
	roomId: 'floor-plan', // 1 to 128 letters, digits, hyphens or underscores
	serverUrl: 'wss://example.com/collab', // optional: a y-websocket compatible server
	user: { name: 'Ada', color: '#2563eb' }, // optional; `role: 'viewer'` makes this window read-only
	intent: 'join', // optional: never seed an empty room with the drawing already open
};
viewer.addEventListener('collaboration-change', (event) => console.log(event.detail));
viewer.collaborationState; // { active, roomId, status, synced, people }
viewer.reconnectCollaboration();
viewer.stopCollaboration(); // or: viewer.collaboration = null
```

- **Transporte**: `session` (una `CollabSession` del anfitrión, que el visor nunca destruye); si no, `provider`; si no, `webrtc`; si no, `serverUrl`; si no, un `BroadcastChannel` entre ventanas de este navegador.
- **Declarativo**: mientras `collaboration` esté definida se comparte el dibujo abierto, y también uno cargado después. Con `null` se abandona la sala. Una sesión que el visor abrió por sí mismo se destruye al salir; una `session` del anfitrión no.
- **Intención (`intent`)**: con `join`, una sala vacía no se rellena con el dibujo local, de modo que quien se une tarde no pueda sobrescribir una sala que aún no ha cargado. Un dibujo cargado o editado después sí se publica. El valor predeterminado es `create`.
- **Roles**: `role: 'viewer'` deja el visor en solo lectura mientras está en la sala (`canEdit` se restablece al salir). El rol es solo orientativo; un servidor debe hacerlo cumplir.
- **Eventos**: `collaboration-change` entrega un `CollaborationState` cuando el uso compartido empieza o termina y cuando cambian el estado de la conexión, la primera sincronización o las personas. Una unión fallida se notifica como `document-error`. Archivo > Compartir usa el mismo controlador, así que también emite el evento.
- **Bindings**: todos los bindings reenvían una propiedad `collaboration` (React, Vue, Solid, Svelte, `mountViewer`/`update` de vanilla, `[collaboration]` de Angular) y retransmiten `collaboration-change` (`events['collaboration-change']`, `@collaboration-change` en Vue, `(collaborationChange)` en Angular). Al desmontar se abandona la sala.
- **Tipos**: `VisioCollaborationOptions` (las opciones compartidas más `onWriteBack`, `writeBackDebounceMs` y `writeBackRetry`) y `VisioCollaborationState` se exportan desde `ooxml-ui/visio`.

## Servidor y punto a punto: token y WebRTC {#server-and-peer-to-peer-auth-token-and-webrtc}

```ts
// A y-websocket compatible server. The token is sent as the `token` query parameter:
// wss://example.com/collab/floor-plan?token=<authToken>
viewer.collaboration = {
	roomId: 'floor-plan',
	serverUrl: 'wss://example.com/collab',
	authToken: await getToken(),
};

// Peer to peer over y-webrtc. Pass the `WebrtcProvider` class from your own y-webrtc install;
// the viewer does not bundle it. The room password defaults to `authToken`.
import { WebrtcProvider } from 'y-webrtc';
viewer.collaboration = {
	roomId: 'floor-plan',
	authToken: 'room password',
	webrtc: { WebrtcProvider, signaling: ['wss://signal.example.com'] },
};
```

El visor entrega estas opciones sin modificar a `openCollaborationSession`. El token solo autentica ante el servidor que usted ejecuta (el servidor debe comprobarlo), y una contraseña de WebRTC solo deja fuera de la sala a los pares que no la tienen; ninguno de los dos hace que los roles sean más que orientativos. Las salas punto a punto nunca confirman una primera sincronización, así que el periodo de gracia de la sesión decide cuándo puede escribir esta ventana.

## Cursores y selecciones en vivo {#live-cursors-and-selections}

Los colaboradores de una misma página ven el puntero de los demás (un cursor con el nombre de la persona en su color) y las formas seleccionadas (un contorno discontinuo en su color) sobre el lienzo. Se mueven y escalan con el zoom y el desplazamiento del dibujo, y desaparecen cuando la persona sale, cambia de página, retira el puntero de la página, borra la selección o queda inactiva.

Cada ventana publica `{ page, x?, y?, selection? }`: el id de la página, el puntero en unidades de dibujo de la página (pulgadas) y hasta 64 ids de formas seleccionadas. `sanitizeVisioPresence` (de `ooxml-core/visio/ui`) valida cada carga entrante. El visor lo aplica a las sesiones que abre él mismo. **Una `session` del anfitrión debe crearse con él** (junto con `VISIO_INITIAL_PRESENCE` como `initialPresence`); de lo contrario sus pares no llevan carga y no se dibujan cursores ni selecciones (nada se rompe).

Límites: el cursor y la selección solo se muestran en la página que ambas personas tienen abierta; los contornos de selección siguen el marco guardado de la forma, no su geometría dibujada.

## Escritura de retorno del propietario {#owner-write-back}

Un anfitrión que quiera guardar el dibujo compartido puede definir `onWriteBack` y, opcionalmente, `writeBackDebounceMs` (5000 por defecto):

```ts
viewer.collaboration = {
	roomId: 'floor-plan',
	user: { name: 'Ada', role: 'owner' },
	onWriteBack: (bytes) => upload(bytes), // VSDX bytes
	writeBackDebounceMs: 5000,
};
```

Solo escribe la ventana cuyo rol es `owner`, de modo que los colaboradores no compiten por guardar. La devolución de llamada se ejecuta cuando el dibujo lleva en reposo el tiempo de espera (tras cambios locales y remotos) y una vez más cuando el propietario sale de la sala si queda un guardado pendiente. Nunca se ejecuta para `collaborator` ni `viewer`. Devuelva una promesa desde `onWriteBack` para confirmar el almacenamiento: hay una escritura en curso a la vez, los cambios intermedios forman una escritura posterior, y un rechazo se reintenta si define `writeBackRetry` (`{ attempts, baseDelayMs?, maxDelayMs? }`). `collaborationState.writeBack` (también en `collaboration-change`) informa `{ status: 'idle' | 'pending' | 'saving' | 'saved' | 'failed', lastSavedAt?, error? }`; solo lo informa el propietario. Un guardado pendiente se completa al salir de la sala, en `pagehide` y cuando la página se oculta; una escritura asíncrona aún puede perderse al cerrar la página. Para escrituras que deban sobrevivir al cierre use `navigator.sendBeacon` o `fetch(..., { keepalive: true })` (límite de unos 64 KiB).

## Qué existe hoy {#what-exists-today}

- Una sesión en vivo basada en `ooxml-core/collab` (Yjs) sobre un `BroadcastChannel`, de modo que las ventanas y pestañas del mismo navegador pueden compartir un dibujo; los anfitriones comparten entre dispositivos mediante `collaboration`.
- Las páginas, las formas (hasta celdas, filas y secciones), los conectores, los patrones (masters) y las imágenes se combinan elemento a elemento (`ooxml-core/visio/collab`); el resto de partes del paquete (estilos, tema, datos, ajustes del documento) se comparte entero y gana la última escritura por parte. El entorno de colaboración se carga solo cuando se inicia el uso compartido.
- Las ediciones simultáneas de formas distintas se conservan ambas; dos ediciones de una misma forma dan el mismo valor en todas las ventanas; una forma o página eliminada gana frente a una edición simultánea.
- Deshacer y rehacer (Ctrl+Z, la cinta) actúan solo sobre las ediciones de esta ventana y nunca revierten las de un colaborador. Los cambios remotos se aplican fuera de su historial, y el historial empieza al unirse.
- Los colaboradores aparecen mediante el elemento de presencia compartido de `ooxml-ui`.
- El nombre de la sesión tiene de 1 a 128 letras, dígitos, guiones o guiones bajos. Una ventana que se une a una sesión cuya sala ya contiene un dibujo lo adopta; si no, rellena la sala (salvo que el anfitrión pida `intent: 'join'`).
- Las sesiones del panel conservan su nombre de BroadcastChannel original (`ooxml-collab:visio-<room>`), de modo que las ventanas de compilaciones anteriores siguen encontrándose.

## Demostración en vivo: dos ventanas, cualquier framework {#live-demo-two-windows-any-framework}

Una sesión de uso compartido es solo un nombre. Los dos paneles de abajo son áreas de pruebas independientes, creadas con enlaces de frameworks distintos, que se unen a la sesión `docs-collab` al cargarse: la ventana A es el área de pruebas vanilla y abre el ejemplo; la ventana B es el área de pruebas de React y recibe el dibujo de la sala. Edite el texto de una forma en un panel y vea cómo el otro lo sigue; deshacer revierte el paso remoto.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer window A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer window B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Combine dos frameworks cualesquiera {#pair-any-two-frameworks}

Todas las áreas de pruebas aceptan `?share=<session>`, que rellena Archivo > Compartir y lo inicia. Solo importa el nombre de la sesión, así que sirve cualquier combinación: la [demostración en vivo](/es/#live-demo) de la página de inicio tiene un selector de framework para cada ventana, y en dos pestañas puede abrir, por ejemplo, `/demo-vue/?sample=1&share=my-room` y `/demo-angular/?share=my-room`. También puede hacerlo a mano en dos ventanas cualesquiera: Archivo, luego Compartir, y el mismo nombre de sesión.

### Por qué Archivo > Compartir solo funciona en el mismo navegador {#why-it-is-same-browser-only}

El panel usa un `BroadcastChannel`, que conecta ventanas, pestañas y marcos de un perfil de navegador en un origen, con independencia del framework, y nada más. Dos equipos distintos no pueden encontrarse así; para ello un anfitrión define `collaboration` con `serverUrl`, `provider` o `session`, y esta compilación no incluye ningún servidor.

## Qué se admite y qué no {#what-is-and-is-not-supported}

| Admitido                                                                                       | No admitido                                                                            |
| ---------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| Sesiones en el mismo navegador sobre BroadcastChannel, basadas en Yjs                          | Un servidor alojado: aporte su propio servidor y-websocket, proveedor o sesión         |
| Combinación por forma, por celda y por página, gana el borrado, deshacer solo local            | Combinar dentro de una celda: dos ediciones de la misma celda dejan un solo valor      |
| Colaboradores mostrados con el elemento de presencia compartido                                | Comentarios; los roles son orientativos (un servidor debe hacerlos cumplir)            |
| Archivo > Compartir y la propiedad `collaboration` como puntos de entrada                      | Cursores entre páginas distintas (solo se dibuja la página compartida)                 |
| Cursores y selecciones en vivo de los colaboradores de la misma página                         | Presencia de una `session` del anfitrión creada sin `sanitizeVisioPresence`            |
| Escritura de retorno del propietario: con espera, confirmación por promesa, reintento y estado | Garantía de que una escritura asíncrona termine al cerrar la página                    |
| `authToken` para un servidor y `webrtc` (su propio `WebrtcProvider`)                           | y-webrtc o servidor de señalización incluidos; validar el token (tarea de su servidor) |

### Límites de la combinación a nivel de forma {#limits-of-shape-level-merging}

- Dos personas que editan propiedades distintas de la misma forma (un movimiento y un cambio de texto o relleno, celdas, filas o secciones distintas) conservan ambas su cambio. Si editan la misma celda o el mismo texto, queda un único valor, igual en todas las ventanas. Los hijos `Text` o `Data` enteros son una unidad cada uno. Las ediciones dentro de un patrón se combinan de la misma forma (por forma y celda), y un patrón que una persona elimina gana frente a una edición concurrente dentro de él.
- Las partes de página se dividen por texto, así que las celdas se conservan exactamente como las escribió el editor (se descarta el espacio entre elementos). Una parte de página que no se puede dividir (sin una lista `Shapes` utilizable) recurre a compartir sus formas como parte de la estructura de la página; gana la última escritura.
- Los patrones y los medios se combinan por elemento: los patrones que añaden dos personas a la vez se conservan ambos, y las imágenes se identifican por su contenido, de modo que las imágenes idénticas se guardan una vez y las distintas con el mismo nombre de archivo se conservan ambas (una se renombra). Los estilos, el tema, la configuración del documento, los datos y la parte de tipos de contenido siguen combinándose por parte; gana la última escritura. El estado de la ventana y la miniatura son locales.
- Si dos ventanas añaden formas con el mismo id a la vez, una se renumera; el mismo contenido de sala siempre da los mismos números en cada ventana. Las fórmulas `Sheet.n!...` siguen a la forma para la que se escribieron, dondequiera y cuando se escribieron (cada fórmula recuerda qué forma nombra, así que una fórmula escrita justo en el momento de la colisión, en una forma anterior o por otra persona, sigue apuntando a la forma que pretendía su autor). Una ventana que puede escribir guarda entonces los números en la sala, para la forma que se movió y para la que conservó su número, de modo que ninguno cambia cuando se añaden más formas después, ni siquiera una forma añadida en el mismo momento con el mismo número. Deshacer nunca revierte ese paso, aunque deshacer una edición de fórmula sí lo revierte. Solo un caso puede volver a mover un número una vez: dos ventanas guardan numeraciones distintas en el mismo momento para una sala que también contiene una adición que la otra no había visto; la sala se estabiliza en cada ventana justo después. Las páginas y los patrones se renumeran de la misma forma (las formas añadidas con un patrón lo siguen).
- Un pegado (glue) que una persona elimina y otra rehace de forma idéntica se conserva. Las imágenes de medios eliminados permanecen en la sala. Lo que una persona elimina (celda, patrón, imagen) mientras otra lo edita conserva la edición; una forma o página eliminada siempre gana.
- Una sala escrita con otra versión del esquema del formato de colaboración se rechaza: la ventana informa de un error y no lee ni escribe en ella. Actualice a la vez todas las ventanas de una sala.
- El historial es por ventana, empieza al unirse y no se conserva tras reconectar. Los cambios remotos que llegan juntos se aplican con una sola reconstrucción, y un cambio que deja el dibujo como estaba no aplica nada. Pero cuando solo cambió el contenido de algunas páginas (y no editó mientras tanto), solo se vuelven a leer esas páginas y se conservan las demás (los estilos, temas y patrones no se vuelven a leer mientras no cambien, y las partes sin cambios del paquete no se vuelven a calcular). Un cambio en la lista de páginas, los patrones, los estilos, el tema, la configuración del documento o las imágenes, el primer dibujo que adopta y el primer cambio remoto tras una de sus propias ediciones vuelven a leer todo el paquete, así que un dibujo muy grande sigue leyéndose entero en esos casos.

Las pruebas están registradas en el [registro de capacidades](/parity), en «Collaboration and comments».
