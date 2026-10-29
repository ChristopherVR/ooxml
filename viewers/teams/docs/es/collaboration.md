# Colaboración

La colaboración es lo que es OpenTeams: cada canal, mensaje, reacción y señal de presencia es un estado compartido que varias personas editan a la vez. Esta página explica cómo está construido, qué muestra la demo en vivo y qué no.

## Cómo funciona

- **El documento compartido es un CRDT de Yjs.** Los canales viven en un `Y.Map`, los mensajes en un `Y.Map` por canal y las reacciones en un mapa plano indexado por mensaje, emoji y usuario, de modo que las ediciones y reacciones concurrentes se fusionan en lugar de sobrescribirse. La lógica reside en `ooxml-core/teams`, construida sobre el área `ooxml-core/collab`, neutral respecto al formato.
- **La presencia es el awareness de Yjs**: quién está en línea, quién está escribiendo y una disponibilidad que tú fijas (disponible, ocupado, ausente). Es orientativa y no está autenticada, como cualquier awareness de Yjs.
- **Las llamadas son WebRTC.** La señalización viaja por tu servidor (o por `BroadcastChannel` en modo local); el contenido multimedia va de par a par en malla completa, por lo que las llamadas sirven para un puñado de personas.
- **El transporte es intercambiable.** Existen dos modos: `server` (tu servidor habla el protocolo de sincronización y-websocket, retransmite la señalización de llamadas y opcionalmente almacena archivos) y `local` (las pestañas de un mismo navegador lo comparten todo mediante `BroadcastChannel`).

## Demo en vivo: dos personas, una sala, cualquier framework

Los dos paneles de abajo son dos aplicaciones de demostración distintas en una misma sala local. Ada ejecuta el `<teams-app>` vanilla y Bob ejecuta el binding de React, y aun así comparten un espacio de trabajo: una sala es solo un nombre (`?room=`), y todos los bindings hablan el mismo protocolo del núcleo. Funcionan en modo local, así que se comunican entre sí únicamente a través de tu navegador. Escribe un mensaje como Ada y míralo aparecer para Bob.

En la [demo en vivo](/es/#live-demo) de la página de inicio cada panel tiene su propio selector de framework, así que puedes emparejar dos cualesquiera entre Vanilla, React, Vue, Angular, Svelte y Solid.

<iframe
	src="/ooxml/teams/demo/?name=Ada&room=docs-collab"
	title="OpenTeams demo as Ada"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/teams/demo-react/?name=Bob&room=docs-collab"
	title="OpenTeams React demo as Bob"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

Lo mismo funciona entre pestañas reales, en cualquier emparejamiento: abre [la demo vanilla como Ada](/demo/?name=Ada&room=docs-collab){target="_self"} y, por ejemplo, [la demo de Vue como Bob](/demo-vue/?name=Bob&room=docs-collab){target="_self"}. Todas las demos aceptan los mismos parámetros `name` y `room`; en modo local solo pueden encontrarse los marcos y pestañas del mismo navegador, así que para personas en máquinas distintas apuntas todos los bindings a tu propio servidor (véase más abajo).

## Qué se admite y qué no

| Admitido                                                                       | No admitido                                                                                                               |
| ------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------- |
| Chat, respuestas y reacciones concurrentes fusionados por un CRDT de Yjs       | Cifrado de extremo a extremo: tu servidor puede leer y almacenar todo                                                     |
| Presencia e indicadores de escritura                                           | Control de acceso por equipo y SSO; la identidad es una declaración salvo que actives las cuentas opcionales del servidor |
| Colaboración en el mismo navegador sin servidor (modo local, como arriba)      | Colaboración entre máquinas distintas sin un servidor que ejecutes tú                                                     |
| Colaboración entre máquinas mediante tu propio servidor de sincronización      | Un servicio alojado: no se ofrece ninguno, y las demos de Pages no tienen servidor detrás                                 |
| Llamadas con audio, vídeo, pantalla compartida y levantar la mano (malla, ~12) | Llamadas de tamaño webinar, grabación, notificaciones                                                                     |

## Entre máquinas

Ejecuta el servidor de referencia (`npx openteams-server --insecure-dev`, solo para desarrollo local; para las cuentas consulta [Aporta tu propio servidor](/es/server)) o cualquier servidor que cumpla los tres contratos (sincronización, señalización, archivos opcionales) y establece `config.mode` en `'server'`. Consulta [Aporta tu propio servidor](/es/server) y [Despliegue](/es/deploy). La coedición de archivos de Office no forma parte de OpenTeams: los archivos se comparten como enlaces, y un anfitrión puede abrirlos en el visor de Word, Excel, PowerPoint o Visio, cada uno con su propio modelo de colaboración.
