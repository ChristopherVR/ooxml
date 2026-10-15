# Primeros pasos

OpenTeams es un espacio de trabajo de equipo de código abierto en el que aportas tu propio servidor: canales y chat, presencia, reuniones con vídeo y pantalla compartida sobre WebRTC, y archivos de Office compartidos en la conversación. Instalas un binding para tu framework, lo apuntas a un servidor (o empiezas sin ninguno) y obtienes la interfaz completa, o un hook sin procesar para construir la tuya.

::: warning Una implementación temprana
Esto no es Microsoft Teams y no está afiliado a Microsoft. **No hay cifrado de extremo a extremo**: quien administre el servidor puede leer todos los mensajes y archivos. Consulta las
[Limitaciones](/es/limitations) antes de depender de él.
:::

## 1. Instala un binding

Cada binding es autónomo: el elemento `<teams-app>` viene incluido, y
[`ooxml-core`](https://www.npmjs.com/package/ooxml-core) (la lógica) y
[`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui) (las primitivas visuales) se instalan como dependencias normales.

::: code-group

```bash [React]
npm install openteams-react-viewer react
```

```bash [Vue]
npm install openteams-vue-viewer vue
```

```bash [Angular]
npm install openteams-angular-viewer @angular/core
```

```bash [Svelte]
npm install openteams-svelte-viewer svelte
```

```bash [Solid]
npm install openteams-solid-viewer solid-js
```

```bash [Vanilla]
npm install openteams-vanilla-viewer
```

:::

| Paquete                    | Para                                 | Dependencia peer |
| -------------------------- | ------------------------------------ | ---------------- |
| `openteams-react-viewer`   | React 18+                            | `react`          |
| `openteams-vue-viewer`     | Vue 3.4+                             | `vue`            |
| `openteams-angular-viewer` | Angular 17+ (standalone)             | `@angular/core`  |
| `openteams-svelte-viewer`  | Svelte 5                             | `svelte`         |
| `openteams-solid-viewer`   | SolidJS 1.9                          | `solid-js`       |
| `openteams-vanilla-viewer` | sin framework (`<teams-app>`)        | ninguna          |
| `openteams-server`         | el servidor de referencia (Node 22+) | ninguna          |

Usa un solo binding por página: cada uno registra los mismos elementos personalizados. Los nombres son `openteams-*` porque el nombre simple `openteams` en npm pertenece a un proyecto ajeno.

## 2. Móntalo

React:

```tsx
import { Teams } from 'openteams-react-viewer';

const config = {
	mode: 'server' as const,
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
};

export function Workspace() {
	return <Teams workspaceId="acme" userName="Ada" config={config} />;
}
```

JavaScript puro:

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

// `config` as in the React example above.
const teams = mountTeams(document.getElementById('teams')!, {
	workspaceId: 'acme',
	userName: 'Ada',
	config,
	onOpenFile: (detail) => console.log('open', detail.attachment.kind, detail.url),
});
teams.update({ workspaceId: 'acme', userName: 'Ada Lovelace', config });
teams.destroy();
```

O con marcado simple, tras llamar a `defineTeamsApp()` desde `openteams-vanilla-viewer`:

```html
<teams-app
	workspace-id="acme"
	user-name="Ada"
	server-config='{"mode":"server","syncUrl":"wss://teams.example.com/sync","signalingUrl":"wss://teams.example.com/signal","iceServers":[{"urls":"stun:stun.example.com:3478"}]}'
></teams-app>
```

Todos los bindings aceptan las mismas props: `workspaceId` (la sala compartida: letras, dígitos, `-` y `_`), `userName`, `userId` (un id estable; por defecto usa uno recordado en este navegador), `config`, `uploadFile` (adónde van los adjuntos), `openers` (manejadores por tipo de archivo de Office) y tres eventos: listo, abrir archivo (llama a `preventDefault()` sobre él para abrir el archivo por tu cuenta) y cambio de configuración (`onReady`, `onOpenFile`, `onConfigChange` en React, Solid, Svelte y vanilla; `@ready`, `@open-file`, `@config-change` en Vue; `(ready)`, `(openFile)`, `(configChange)` en Angular). Consulta las
[guías por framework](/frameworks/react) para cada binding.

Si omites `config`, el elemento ofrece un cuadro de diálogo de ajustes (recordado en este navegador) o funciona en modo local.

## 3. Elige dónde vive el estado

| Modo     | `config`                                                        | Qué ocurre                                                                                                                                                    |
| -------- | --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `local`  | `{ mode: 'local', iceServers: [...] }`                          | Sin servidor. Las pestañas de un mismo navegador comparten chat, presencia y llamadas mediante `BroadcastChannel`. Los archivos se comparten solo por nombre. |
| `server` | `{ mode: 'server', syncUrl, signalingUrl, iceServers, token? }` | Tu servidor sincroniza el documento compartido, retransmite la señalización de llamadas y (opcionalmente) almacena archivos.                                  |

El modo local es el que usan las [demos en vivo](/demos). Para cualquier caso con más de un navegador, ejecuta un servidor:

```bash
npx openteams-server                    # the reference server on 127.0.0.1:8787
```

Después usa `ws://127.0.0.1:8787/sync` y `ws://127.0.0.1:8787/signal` (o `wss://` detrás de TLS). La [guía del servidor](/es/server) cubre el protocolo, las variables de entorno y lo que el servidor de referencia no hace; el [despliegue](/es/deploy) cubre TLS, tokens, TURN y Docker.

## 4. O construye tu propia interfaz

Cada binding expone además un hook sin procesar sobre el cliente del núcleo: el estado entra, las acciones simples salen.

```tsx
import { useTeams } from 'openteams-react-viewer';

const { client, state } = useTeams({
	workspaceId: 'acme',
	user: { id: 'u1', name: 'Ada' },
	config,
});
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

`state` es una instantánea inmutable de `TeamsState` (estado de la conexión, canales con contadores de no leídos, los mensajes y archivos del canal seleccionado, personas, escritura, resultados de búsqueda, la llamada). Las acciones incluyen `select`, `createChannel`, `send`, `startReply`, `startEdit`, `deleteMessage`, `toggleReaction`, `search`, `openCall`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`, `toggleScreenShare` y `toggleHand`. La demo de React muestra ambos estilos lado a lado.

## Ejecútalo desde el código fuente

Requiere Node 22+ y Bun. Con el repositorio [`ooxml`](https://github.com/ChristopherVR/ooxml) clonado junto a este (`../ooxml-core`), las demos y las pruebas usan su código fuente; sin él, usan los paquetes publicados `ooxml-core` y `ooxml-ui`.

```bash
bun install
bun run dev        # server on :8787, vanilla demo on :5173, React demo on :5174
```

Abre `http://127.0.0.1:5173/?name=Ada` y, en otra pestaña, `...?name=Bob`: dos personas en un mismo espacio de trabajo. `?local=1` funciona sin el servidor. `TEAMS_USE_DIST=1` hace que las demos usen los paquetes compilados (antes, `bun run build:packages`) y los paquetes ooxml publicados en lugar de cualquier código fuente. Las cámaras y los micrófonos requieren `localhost` o HTTPS.
