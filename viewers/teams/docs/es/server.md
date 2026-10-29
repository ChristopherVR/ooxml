# Aporta tu propio servidor

OpenTeams no tiene un servicio alojado. El espacio de trabajo funciona **sin servidor** (`mode: 'local'`, solo pestañas de un mismo navegador) o contra **un servidor que tú ejecutas** y que cumpla tres contratos pequeños. La implementación de referencia, `openteams-server`, cumple los tres en un único proceso de Node de unas 300 líneas; cualquier otro servidor que hable los mismos contratos funciona igual de bien.

::: danger Sin cifrado de extremo a extremo
El chat viaja como actualizaciones de Yjs a través de tu servidor, que almacena en claro el documento de la sala y todos los archivos subidos y puede leerlos. El contenido multimedia de las llamadas va cifrado entre pares (DTLS-SRTP) y un relé TURN no puede descifrarlo, pero la señalización pasa por el servidor.
:::

## Ejecuta el servidor de referencia

Requiere Node 22 o posterior.

```bash
TEAMS_USERS_FILE=users.json TEAMS_SESSION_SECRET=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
npx openteams-server --insecure-dev     # solo para desarrollo local: escucha en 127.0.0.1:8787, sin cuentas
```

El servidor es **seguro por defecto**: sin archivo de cuentas (`TEAMS_USERS_FILE`), hook `authenticate` ni hook `authorizeSync`
se niega a arrancar. `--insecure-dev` (opción `insecureDev: true`) lo arranca igualmente, con una advertencia, para desarrollo
local. Se configura mediante el entorno más los indicadores `--insecure-dev`, `--trust-proxy[=N]` e `--insecure` (véanse
`TEAMS_INSECURE` y `TEAMS_TRUST_PROXY` en la [guía de despliegue](./deploy.md)). `SIGTERM` y `SIGINT` guardan todas las salas antes de salir.

| Variable        | Valor por defecto | Significado                                                                                                                        |
| --------------- | ----------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`          | `8787`            | Puerto de escucha.                                                                                                                 |
| `HOST`          | `127.0.0.1`       | Dirección de escucha. Enlaza `0.0.0.0` solo detrás de TLS.                                                                         |
| `TEAMS_TOKEN`   | ninguno           | Si se define, los sockets lo reciben como `?token=` y las peticiones de archivos como `Authorization: Bearer`.                     |
| `TEAMS_ORIGINS` | ninguno           | `Origin` permitidos, separados por comas, para sockets y archivos. Sin valor: cualquier origen con token, solo loopback sin token. |
| `TEAMS_DATA`    | `./data`          | Directorio de datos: instantáneas `rooms/<room>.bin` y `files/<workspace>/`.                                                       |

El servidor lee más variables para las cuentas (`TEAMS_USERS_FILE`, `TEAMS_SESSION_SECRET`, `TEAMS_ADMIN_TOKEN`, `TEAMS_AUTH_RATE_*`) y para los límites (`TEAMS_INSECURE`, `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP`, `TEAMS_MAX_ROOMS`, `TEAMS_MAX_DOC_MB`, `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`, `TEAMS_FILES_QUOTA_MB`). La [guía de despliegue](/es/deploy) lista los límites; el README del servidor es la referencia completa. Las opciones correspondientes de `createTeamsServer` son `users`, `authenticate`, `sessionSecret`, `adminToken`, `sessionRate`, `authorizeSync`, `maxConnections`, `maxRooms`, `messageRate`, `filesQuotaBytes`, `insecure`, `insecureDev` y `trustProxy`.

O incrústalo:

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({
	users: { alice: 'long-random-secret' },
	token: process.env.TEAMS_TOKEN,
	origins: ['https://app.example.com'],
	dataDir: '/var/lib/openteams',
	maxFileBytes: 25 * 1024 * 1024, // default
	log: console.log,
});
await server.listen(8787, '127.0.0.1');
// server.http is the node:http server; server.close() flushes every room to disk.
```

Después apunta el cliente hacia él:

```ts
const config = {
	mode: 'server',
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	token: 'change-me', // only if TEAMS_TOKEN is set
};
```

[Despliegue del servidor de referencia](/es/deploy) cubre TLS, el token, TURN, Docker y un ejemplo con Caddy.

## Endpoints

| Endpoint                              | Qué hace                                                                                                                                   |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `ws /sync/<room>`                     | Sincronización compatible con y-websocket: un `Y.Doc` y su awareness por sala, persistido en `rooms/<room>.bin`. Mensajes de hasta 32 MB.  |
| `ws /signal/<room>`                   | Relé JSON para la señalización de llamadas: reenvía cada mensaje de texto a los demás sockets de la sala. Mensajes de hasta 256 KB.        |
| `POST /files/<workspace>/<name>`      | Almacena un adjunto (límite de 25 MB) y responde `201 { url, size }`.                                                                      |
| `GET /files/<workspace>/<name>`       | Lo descarga como adjunto (`Content-Disposition: attachment`, `Referrer-Policy: no-referrer`, `nosniff`).                                   |
| `POST /files/link/<workspace>/<name>` | Devuelve `{ url }`: un enlace firmado para ese único archivo, válido durante cinco minutos, de modo que el token nunca aparece en una URL. |
| `GET /health`                         | `{ ok, rooms, calls, auth }`.                                                                                                              |
| `POST /auth/session`                  | Firma una sesión para `{ user, credential }` (con límite de frecuencia). Consulta "Accounts and sessions" en el README del servidor.       |
| `POST /auth/check`, `/auth/signout`   | Comprueba una `{ session }` (204 o 401); cierra sesión, lo que revoca las sesiones del usuario.                                            |
| `POST /auth/revoke`                   | Revocación por el operador de `{ user }` con `TEAMS_ADMIN_TOKEN`; un 404 sin token de administrador.                                       |

Los nombres de sala deben cumplir `^[\w-]{1,128}$` y los ids de espacio de trabajo en las rutas de archivos `^[\w-]{1,64}$`; los nombres de archivo se depuran a letras, dígitos, `.`, `-`, `_`, espacios y paréntesis, y se limitan a 120 caracteres.

## Los contratos

Cualquier servidor que haga estas tres cosas funciona con todos los bindings.

### Sincronización

`ws(s)://host/<base>/<room>` habla el formato de transmisión de y-websocket (sync step 1 y 2, actualizaciones, awareness, query awareness). Un servidor y-websocket estándar sirve. Un servidor que mantenga un `Y.Doc` por sala y lo persista ofrece el historial a quienes se unan más tarde.

### Señalización

`ws(s)://host/<base>/<room>` reenvía cada mensaje de texto JSON a los **demás** sockets de la sala y les envía `{ "type": "bye", "from": <id> }` cuando se cierra un socket. Un socket habla en nombre de un único id de par (el primer `from` que usa). El servidor no necesita entender los mensajes.

### Archivos

Opcional. `POST|GET <origin>/files/<workspace>/<name>` con el token en una cabecera `Authorization: Bearer`, más `POST <origin>/files/link/<workspace>/<name>`, que devuelve un `{ url }` firmado de corta duración para abrir un archivo en una pestaña del navegador. El origen se toma de `syncUrl`. Sin servidor de archivos, entrega al elemento una función `uploadFile` (para guardar los adjuntos en tu propio almacén de objetos); de lo contrario, los adjuntos se comparten solo por nombre.

Ambos sockets reciben `?token=<token>` cuando hay un token configurado, porque un navegador no puede establecer cabeceras en un WebSocket. Aplícalo, junto con una lista de orígenes permitidos, en el servidor.

## Llamadas: STUN y TURN

`iceServers` contiene las URL de STUN y TURN. El STUN público basta en redes abiertas y nunca basta detrás de un NAT simétrico o de cortafuegos estrictos: ejecuta tu propio TURN (por ejemplo coturn) y da credenciales a los clientes. El `parseServerConfig` del núcleo rechaza las entradas TURN sin credenciales. Las credenciales TURN estáticas son visibles para cualquiera que pueda abrir la aplicación; para cualquier uso público, emite credenciales de corta duración. La [guía de despliegue](/es/deploy#llamadas-a-traves-de-nat-turn) incluye un ejemplo de coturn.

## Lo que el servidor de referencia no hace

Ten esto claro antes de poner personas en él:

- **Sin cifrado de extremo a extremo.** Almacena y puede leer todos los mensajes y archivos.
- **Cuentas opcionales y pequeñas.** Con `users` (o `TEAMS_USERS_FILE`) o una función `authenticate`, el servidor firma sesiones (`POST /auth/session`), toma el id de usuario de la sesión y puede revocarla. Sin ellas, `TEAMS_TOKEN` es un único secreto compartido para todos, almacenado en el navegador, y los ids de usuario son declaraciones. No hay restablecimiento de contraseñas, gestión de cuentas, SSO ni control de acceso por equipo. Las comprobaciones de autor y de rol en el cliente pueden eludirse con un cliente hostil, así que el servidor es el límite de confianza: usa `authenticate` con tu proveedor de identidad y `authorizeSync` para permisos por sala.
- **Sin TLS propio.** Ejecútalo detrás de un proxy inverso que termine TLS y actualice las conexiones a WebSocket; los navegadores rechazan `ws://` desde una página `https://`.
- **Un proceso, salas en memoria.** No escala horizontalmente. Una sala se descarga 30 segundos después de que se cierre su último socket.
- **Una instantánea por sala, sin compactación.** El documento crece con el historial y nunca se poda.
- **Almacenamiento de archivos simple.** Sin listado, sin borrado, sin cuota por espacio de trabajo; subir un archivo con un nombre existente lo reemplaza. Los archivos están limitados a 25 MB cada uno.
- **Sin limitación de tasa** y sin servidor multimedia: las llamadas son de par a par (malla).
