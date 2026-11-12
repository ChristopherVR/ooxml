# Arquitectura

OpenTeams se divide igual que los demás productos construidos sobre
[`ooxml`](https://github.com/ChristopherVR/ooxml) (Word, Excel, PowerPoint): la lógica y las
primitivas visuales están allí, incluido el propio elemento `<teams-app>`. Esta carpeta
(`viewers/teams` del repositorio `ooxml`) contiene los bindings de framework, el servidor de
referencia y la documentación; las demos están en `demos/teams/` y las pruebas de navegador en
`e2e/teams/`, en la raíz.

| Capa                   | Dónde                                                                   | Responsabilidad                                                                                                               |
| ---------------------- | ----------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| Lógica                 | `ooxml-core/teams` (`src/core/teams` en el repositorio ooxml)           | el CRDT del chat, la presencia, las llamadas, la señalización, la configuración del servidor, el store `createTeamsClient`    |
| Colaboración           | `ooxml-core/collab`                                                     | sesiones y transportes de Yjs (WebSocket, `BroadcastChannel`), neutral respecto al formato                                    |
| Primitivas visuales    | `ooxml-ui/teams` (`src/ui/src/teams`)                                   | avatar, barra de aplicaciones, lista de canales, lista de chats, redactor, prevuelo, cuadrícula de llamada y controles        |
| Aplicación             | `ooxml-ui/teams` (`src/ui/src/teams/app`)                               | `<teams-app>`: compone las primitivas sobre el store, el cuadro de diálogo de ajustes y las utilidades del store sin procesar |
| Capa de binding        | `ooxml-ui/teams` (`src/ui/src/teams/app/bind.ts`)                       | `bindTeams`, `pickTeamsProps` y `TEAMS_CONTRACT`: el cableado de props y eventos que usan los bindings                        |
| Componente web         | `packages/web-component` aquí (`teams-viewer` privado, nunca publicado) | reexporta `ooxml-ui/teams`, así los bindings importan un solo nombre; se incluye en cada binding                              |
| Bindings               | `packages/{react,vue,angular,svelte,solid,vanilla}` aquí                | adaptadores de ciclo de vida más un hook sin procesar en cada uno, publicados como `openteams-<framework>-viewer`             |
| Servidor de referencia | `server/` aquí                                                          | sincronización Yjs, retransmisión de señalización, almacenamiento de archivos, publicado como `openteams-server`              |

La descripción del área que hace el propio núcleo está en
[`docs/teams-area.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md).

## El cliente: un estado, acciones sencillas {#the-client-one-state-plain-actions}

Todo lo que necesita una interfaz procede de un único cliente creado con `createTeamsClient` (los
bindings lo envuelven como `createTeams`, que añade un `localStorage` protegido para las
instantáneas sin conexión):

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
```

`getState()` devuelve una única instantánea inmutable de `TeamsState`. Cada botón es una acción del
cliente. Los errores que un usuario debe ver llegan mediante `on('notice', ...)`. La interfaz nunca
toca Yjs, `RTCPeerConnection` ni el servidor directamente: **el estado entra, las acciones salen**.

## `<teams-app>` {#} {#} {#} {#} {#}

Un elemento Lit con una plantilla y un archivo CSS real (`teams-app.ts`, `teams-app.css`, en
`src/ui/src/teams/app`). Decide qué se muestra en pantalla y reenvía la intención del usuario a las
acciones del cliente; los colores, espaciados y radios proceden de los tokens de `ooxml-ui`
(`var(--office-...)`). Todo lo que envía un par se muestra como texto (semántica de `textContent`),
nunca como marcado, y el núcleo vuelve a validar todo lo que lee del documento compartido.

## Bindings {#bindings}

Un binding es un adaptador de ciclo de vida: crea `<teams-app>`, reenvía las props con `bindTeams`
(incluida `class`, según el framework `class` o `className`) y vuelve a emitir los eventos, todo ello
en `ooxml-ui/teams` (`src/ui/src/teams/app/bind.ts`, sobre el kit común `ooxml-ui/bindings`).
`onOpenFile` es `(detail, event)` en todos los casos salvo en Angular, cuya salida de valor único
emite `{ detail, event }`. Cada uno expone además el hook sin procesar sobre `createTeams`:

| Binding  | Componente                               | Hook sin procesar                                         |
| -------- | ---------------------------------------- | --------------------------------------------------------- |
| React    | `<Teams />`                              | `useTeams(options)` (`useTeamsClient`, `useTeamsState`)   |
| Vue 3    | `<Teams />`                              | `useTeams(() => options)` devuelve refs superficiales     |
| Solid    | `Teams(props)`                           | `createTeamsClient(() => options)` devuelve señales       |
| Svelte 5 | `Teams.svelte` (exportación por defecto) | `teamsStore(options)` desde `/runtime` (un store legible) |
| Angular  | `<teams-workspace>`                      | `TeamsService` (señales)                                  |
| Vanilla  | `mountTeams(el, props)`                  | `createTeams(options)`                                    |

Un error de comportamiento se corrige una sola vez, en el componente web o en el núcleo. Un error de
hook o de cableado se corrige en todos los bindings en el mismo cambio.

En tiempo de compilación, `scripts/build-packages.mjs` incluye el componente web privado, con su
CSS, en cada paquete de framework y aplana sus declaraciones, de modo que un tarball publicado
importa solo lo que declara su manifiesto (`ooxml-core`, `ooxml-ui`, `lit` y la dependencia peer del
framework).

## Chat, presencia y llamadas {#chat-presence-and-calls}

- **Chat.** Los canales viven en un único `Y.Map`, los mensajes en un `Y.Map` por canal, y las
  reacciones en un mapa plano indexado por mensaje, emoji y usuario, de modo que las reacciones
  simultáneas nunca se sobrescriben entre sí. Las ediciones y los borrados solo los puede hacer el
  autor en el cliente (un borrado es una lápida blanda).
- **La presencia** es la awareness de Yjs: personas conectadas, escritura, estado y mensaje de
  estado. Es orientativa y no autenticada. Los pares reciben también el `availability` heredado de
  tres valores.
- **Los fijados** son un mapa de Yjs por canal; cualquier miembro puede fijar o desfijar.
- **Las llamadas** son una malla completa de conexiones entre pares con el patrón de negociación
  perfecta; silenciar, la cámara y compartir pantalla cambian de pista sin volver a negociar. La
  malla es honesta hasta unas doce personas (`maxParticipants`, por defecto 12).
- **Los archivos** son enlaces, nunca bytes incrustados: se suben mediante `uploadFile` o el endpoint
  `/files` del servidor. Los archivos de Word, Excel, PowerPoint y Visio se reconocen por su nombre
  para que un host pueda abrirlos en el visor correspondiente (la prop `openers` o el evento de
  apertura de archivo).

## Transportes {#transports}

En modo `local`, el cliente sincroniza y señaliza mediante `BroadcastChannel`, de modo que solo se
ven las pestañas de un mismo navegador; las demos de GitHub Pages funcionan así. En modo `server`
usa el protocolo de y-websocket para la sincronización y un relé JSON para la señalización;
consulta [Aporta tu propio servidor](/es/server).
