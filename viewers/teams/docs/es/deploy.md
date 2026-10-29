# Ejecutar tu propio servidor

## El servidor de referencia

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable                                                | Significado                                                                                                                                                                                                                                                                                      |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `PORT`, `HOST`                                          | Dirección de escucha. Enlaza `0.0.0.0` solo detrás de TLS.                                                                                                                                                                                                                                       |
| `TEAMS_TOKEN`                                           | Si se define, los sockets lo reciben como `?token=` y las peticiones de archivos como `Authorization: Bearer`.                                                                                                                                                                                   |
| `TEAMS_ORIGINS`                                         | `Origin` permitidos, separados por comas, para sockets y archivos. Vacío: cualquier origen si hay token; sin token, solo orígenes de loopback y clientes que no envían `Origin`.                                                                                                                 |
| `TEAMS_DATA`                                            | Directorio de datos (por defecto `./data`): instantáneas `rooms/<room>.bin` y `files/`.                                                                                                                                                                                                          |
| `TEAMS_TRUST_PROXY`, `--trust-proxy[=N]`                | Agrupa el limitador de tasa de autenticación y el límite de conexiones por dirección por `X-Forwarded-For` (la entrada del N-ésimo proxy de confianza, por defecto 1). Desactivado por defecto; actívalo solo detrás de un proxy que establezca la cabecera, o los clientes pueden falsificarla. |
| `--insecure-dev`                                        | Arranca sin cuentas y sin hook `authenticate` (opción `insecureDev`). Sin él, el servidor solo arranca si se define `TEAMS_USERS_FILE`. Solo para desarrollo.                                                                                                                                    |
| `TEAMS_INSECURE`                                        | `1` (o `--insecure`) permite un `HOST` que no sea loopback sin token. Si no, el servidor se niega a arrancar.                                                                                                                                                                                    |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Sockets abiertos en total (por defecto 500) y por dirección remota (20).                                                                                                                                                                                                                         |
| `TEAMS_MAX_ROOMS`                                       | Salas de sincronización y de llamada (por defecto 1000 cada una).                                                                                                                                                                                                                                |
| `TEAMS_MAX_DOC_MB`                                      | Mayor documento de sala persistido en MB (por defecto 64); las actualizaciones que lo superen se rechazan.                                                                                                                                                                                       |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Mensajes por segundo (por defecto 100) y ráfaga (200) por socket; un socket que los supere se cierra.                                                                                                                                                                                            |
| `TEAMS_FILES_QUOTA_MB`                                  | Límite opcional para todo el directorio de archivos en MB (desactivado por defecto); las subidas que lo superen reciben 507.                                                                                                                                                                     |
| `TEAMS_WORKSPACE_QUOTA_MB`                              | Archivos por espacio de trabajo en MB (por defecto 1024; un chat privado cuenta para su espacio de trabajo); las subidas que lo superen reciben 507.                                                                                                                                             |
| `TEAMS_COMPACT_INTERVAL_MIN`                            | Minutos entre compactaciones de las instantáneas de las salas (por defecto 360, `0` desactiva el temporizador). Una sala también se compacta al cargarse.                                                                                                                                        |
| `TEAMS_METRICS_TOKEN`                                   | Activa `GET /metrics` (texto de Prometheus), servido solo con `Authorization: Bearer <este token>`. Desactivado sin él.                                                                                                                                                                          |
| `TEAMS_LOG_FORMAT`, `--log-format=json`                 | `text` (por defecto) o `json`: un objeto JSON por línea con `time`, `level` y `msg`.                                                                                                                                                                                                             |

Sirve `ws /sync/<room>` (compatible con y-websocket, un `Y.Doc` persistido por sala), `ws /signal/<room>` (un relé JSON) y `POST|GET /files/<workspace>/<name>` (límite de 25 MB, nombres depurados, descargas enviadas como adjuntos con `Referrer-Policy: no-referrer`). Con un token, el navegador abre los archivos mediante `POST /files/link/<workspace>/<name>` (token en la cabecera), que devuelve un enlace firmado para ese único archivo y válido durante cinco minutos. `GET /health` informa del número de salas y de llamadas.

**Ponlo detrás de TLS.** Los navegadores rechazan `ws://` desde una página `https://`, y un token o la señalización de llamadas sobre HTTP simple pueden leerse en la red. Sirve cualquier proxy inverso que actualice a WebSocket (nginx, Caddy, Traefik). Ejemplo con Caddy:

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

Después apunta la aplicación a `wss://teams.example.com/sync` y `wss://teams.example.com/signal`.

**El token es un secreto compartido** almacenado en el navegador: mantiene fuera a los extraños del servidor de un equipo pequeño, no es autenticación por usuario. Para cuentas reales, da al servidor `users` (o `TEAMS_USERS_FILE`) o una función `authenticate` para tu proveedor de identidad, y define `TEAMS_SESSION_SECRET` para que las sesiones sobrevivan a los reinicios (consulta "Accounts and sessions" en el README del servidor). Los tres contratos siguen siendo los mismos.

## Modelo de seguridad

- **Un único token compartido abre todas las salas y todos los archivos.** La identidad por usuario solo existe si activas las cuentas, y los permisos por sala solo si pasas `authorizeSync`. Trata el token como la contraseña del equipo.
- **Todos los miembros de una sala confían entre sí.** El servidor solo mueve bytes de Yjs, así que
  cualquier miembro puede editar o borrar los mensajes de otros y falsificar identificadores de autor
  y roles. No invites a quien no sea de confianza para toda la sala.
- **El token y la sesión viajan en la cabecera `Sec-WebSocket-Protocol`** (`openteams.session.<session>`,
  `openteams.token.<token>`, junto a un `openteams` sin más), que no aparece en URLs ni registros. El cliente incluido
  solo recurre a la cadena de consulta si un valor contiene caracteres que un subprotocolo no puede llevar. Ejecuta el
  servidor detrás de TLS. Los archivos nunca llevan el token en una URL: usan la cabecera `Authorization` o un enlace
  firmado.
- **La autoría se aplica** a los usuarios con sesión: nadie puede escribir en nombre de otra persona, editar o borrar
  sus mensajes, ni cambiar canales y pestañas que no creó, salvo que la cuenta tenga `"role": "moderator"` en el
  archivo de usuarios. `TEAMS_ENFORCE_AUTHORSHIP=0` lo desactiva.
- **Seguro por defecto.** Sin cuentas, un hook `authenticate` ni `authorizeSync`, el servidor se niega a arrancar;
  `--insecure-dev` es la excepción explícita y con advertencia. Un token compartido por sí solo no es autenticación.
- **Los valores por defecto son locales.** Sin token, el servidor rechaza un `HOST` que no sea loopback
  (salvo con `TEAMS_INSECURE=1`) y solo acepta orígenes de loopback, de modo que un sitio web que
  visites no pueda controlar un servidor de tu propia máquina. Define `TEAMS_TOKEN` y `TEAMS_ORIGINS`
  en cualquier despliegue compartido.
- **Usa credenciales TURN de corta duración** (véase más abajo); las estáticas las ve cualquiera que abra la aplicación.
- **Los límites** de conexiones, salas, tamaño de documento, tasa de mensajes y almacenamiento de
  archivos (las variables anteriores) acotan el abuso, pero no sustituyen la limitación de tasa propia
  de un proxy inverso.
- Un socket solo puede hablar por los identificadores de presencia y el identificador de par de llamada
  que anunció él mismo, y una señal de llamada dirigida se entrega solo a ese par. Una subida nunca
  sobrescribe un archivo: un nombre ya usado se guarda con otro nuevo, que devuelve el campo `url` de la respuesta.

## Llamadas a través de NAT: TURN

Las llamadas son de par a par. El STUN público funciona en redes abiertas; detrás de un NAT simétrico o de cortafuegos estrictos necesitas un relé TURN, que además transporta el contenido multimedia solo como texto cifrado (DTLS-SRTP). Un coturn mínimo (`turnserver.conf`):

```
listening-port=3478
tls-listening-port=5349
realm=turn.example.com
fingerprint
lt-cred-mech
user=teams:CHANGE_ME
no-multicast-peers
cert=/etc/letsencrypt/live/turn.example.com/fullchain.pem
pkey=/etc/letsencrypt/live/turn.example.com/privkey.pem
```

y, en los ajustes de servidor de la aplicación, un servidor ICE por línea:

```
stun:turn.example.com:3478
turn:turn.example.com:3478,turns:turn.example.com:5349 teams CHANGE_ME
```

Las credenciales TURN estáticas son visibles para cualquiera que pueda abrir la aplicación. Para cualquier uso público, emite credenciales de corta duración (coturn `use-auth-secret` más un pequeño endpoint que las firme) y pásalas en los `iceServers` que entregas al elemento.

## Docker

```dockerfile
FROM node:22-alpine
WORKDIR /app
COPY server/package.json ./
RUN npm install --omit=dev
COPY server/bin.mjs server/index.mjs server/auth.mjs server/passwords.mjs server/guards.mjs \
  server/rate-limit.mjs server/chat-access.mjs server/validator.generated.mjs server/files.mjs server/metrics.mjs server/room-store.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data TEAMS_USERS_FILE=/config/users.json
RUN mkdir /data && chown node:node /data
VOLUME /data
EXPOSE 8787
USER node
CMD ["node", "bin.mjs"]
```

Constrúyelo desde `viewers/teams`. Copia todos los módulos que importa el punto de entrada (los `files` de `server/package.json`); copiar solo `index.mjs` falla por un módulo ausente. La receta se comprobó ejecutando esos archivos con `npm install --omit=dev` en Node; la imagen en sí no se construyó.

Ejecútalo con un archivo de cuentas montado, `TEAMS_SESSION_SECRET` y `TEAMS_ORIGINS` definidos, detrás de tu proxy TLS:

```bash
docker run -d --name openteams -p 127.0.0.1:8787:8787 \
  -v openteams-data:/data -v "$PWD/users.json:/config/users.json:ro" \
  -e TEAMS_SESSION_SECRET=change-me -e TEAMS_ORIGINS=https://app.example.com \
  openteams-server
```

Sin el archivo de cuentas montado, el contenedor termina con el mensaje "refusing to start"; no añadas `--insecure-dev` a un
contenedor accesible desde una red. Añade `--trust-proxy` al comando (`CMD ["node", "bin.mjs", "--trust-proxy"]`) solo si un
proxy inverso establece `X-Forwarded-For`. `docker stop` envía `SIGTERM`, a lo que el servidor responde guardando primero
todas las salas (Docker espera diez segundos por defecto, lo que basta). La imagen crea `/data` propiedad de `node`; con un
montaje por enlace, cambia el propietario del directorio del host a uid 1000 con `chown`.

## Operaciones

### Copia de seguridad y restauración

Todo lo que guarda el servidor está bajo `TEAMS_DATA` (por defecto `./data`):

| Ruta                       | Contiene                                                                                                     |
| -------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `rooms/<room>.bin`         | Una instantánea Yjs por sala: todo el chat, canales, calendarios y datos de pestañas.                        |
| `files/<workspace>/`       | Archivos subidos.                                                                                            |
| `file-meta/<scope>.json`   | Quién subió cada archivo (decide quién puede borrarlo). Cópialo junto con `files/`.                          |
| `session-revocations.json` | Generaciones de revocación de sesiones (si se pierde, las sesiones revocadas vuelven a valer hasta caducar). |

Guarda también lo que no está en el directorio de datos: tu archivo de cuentas (`TEAMS_USERS_FILE`) y `TEAMS_SESSION_SECRET`.
Cambiar o perder el secreto cierra la sesión de todos, lo cual no es grave.

Para hacer una copia de seguridad, copia todo el directorio de datos. Las instantáneas de las salas se escriben en un archivo
temporal y se renombran, así que una copia nunca ve una instantánea a medio escribir; pero una copia tomada mientras el servidor
funciona puede perder los cambios de la última media sección (las escrituras tienen un retardo de 500 ms) y capturar un
archivo más nuevo que los de su sala. Para una copia coherente, detén primero el servidor (`SIGTERM` guarda todas las salas antes
de salir), copia y vuelve a arrancarlo; o haz una instantánea del volumen a nivel de sistema de archivos.

Para restaurar, detén el servidor, sustituye el directorio de datos por la copia (manteniendo juntos `rooms/`, `files/` y
`session-revocations.json`), restaura el archivo de cuentas y el secreto, y arráncalo. Los clientes conectados se reconectan y
combinan: Yjs conserva lo que tengan y falte en la copia, así que una restauración no descarta trabajo que siga abierto en un navegador.

### Archivos, compactación y métricas

- **Permisos de archivos.** Un archivo vive en un ámbito: el id del espacio de trabajo para los archivos de canal
  (cualquier usuario con sesión del servidor) o el nombre de sala de un chat privado (`chat-<hash>-<chat id>`), que solo
  los miembros de ese chat pueden listar, leer, subir o enlazar. El servidor comprueba la lista de miembros del chat igual
  que para su sala de sincronización. Con cuentas, cada petición de archivo debe llevar la sesión firmada en
  `X-Teams-Session` (el cliente lo hace); un enlace de descarga firmado solo se emite tras esa comprobación y abre su
  único archivo durante cinco minutos. Con `--insecure-dev` vale el `X-Teams-User` declarado, como toda declaración allí.
  Los archivos de canal no se restringen por canal: todos los usuarios del servidor pueden leerlos.
- **Listar y borrar.** `GET /files/<scope>` responde `{ files: [{ name, size, modified, uploader }] }`.
  `DELETE /files/<scope>/<name>` lo permite quien lo subió o un moderador (`204`, si no `403`). Los archivos subidos antes
  de esta versión no tienen quien los subió registrado, así que solo un moderador puede borrarlos.
- **Limpieza.** Una subida se escribe en un archivo oculto `.part` que una subida fallida o cancelada elimina. Al arrancar,
  el servidor borra todos los `.part` sobrantes (y los temporales de salas obsoletos), y cada 15 minutos los de más de una
  hora. Los archivos subidos pero nunca adjuntados a un mensaje permanecen hasta que alguien los borre.
- **Cuotas.** El límite es por espacio de trabajo (`TEAMS_WORKSPACE_QUOTA_MB`, por defecto 1024) y cuenta sus archivos de
  canal y los de sus chats privados; borrar libera espacio. `TEAMS_FILES_QUOTA_MB` es un límite opcional para todo el directorio.
- **Compactación.** Al cargar una sala, y cada `TEAMS_COMPACT_INTERVAL_MIN` minutos para todas, el servidor recodifica la
  sala como una única actualización de Yjs sobre un documento nuevo y la escribe solo si es menor y describe exactamente el
  mismo estado. Se escribe en un archivo temporal, se vuelca a disco y se renombra sobre la instantánea, de modo que un
  fallo deja la instantánea antigua o la nueva, nunca una a medias; los temporales interrumpidos se eliminan al arrancar.
  Las copias de seguridad anteriores y posteriores a una compactación se restauran ambas.
- **Métricas.** Define `TEAMS_METRICS_TOKEN` y consulta `GET /metrics` con `Authorization: Bearer <token>`:
  `teams_rooms_open`, `teams_connections{kind}`, `teams_stored_bytes{kind="rooms|files"}`,
  `teams_updates_rejected_total{code}`, `teams_auth_failures_total{kind}`, `teams_rate_limited_total{scope}` y
  `teams_room_compactions_total`. Sin el token responde 404. No lo expongas a internet.
- **Registros.** `--log-format=json` (o `TEAMS_LOG_FORMAT=json`) escribe un objeto JSON por línea: `{ "time", "level", "msg" }`.

### Actualizar

1. Lee el `CHANGELOG.md` del servidor para el rango de versiones que cruzas; los cambios incompatibles están marcados.
2. Haz una copia de seguridad del directorio de datos.
3. Detén el servidor (`SIGTERM`), instala la nueva versión (`npx openteams-server@<version>`, o reconstruye la imagen) y arráncalo. Las instantáneas de las salas son actualizaciones Yjs sencillas, y las antiguas se leen tal cual.
4. Comprueba `GET /health` y abre una sala.

Actualizar desde una versión anterior al valor seguro por defecto: un servidor que arrancaba sin cuentas ahora se niega a hacerlo.
Añade cuentas (`TEAMS_USERS_FILE`), un hook `authenticate` o `authorizeSync`, o pasa `--insecure-dev` si (y solo si) es un servidor
de desarrollo local. Un `TEAMS_TOKEN` compartido por sí solo ya no arranca el servidor.

### Lista de verificación para producción

- [ ] Cuentas configuradas (`TEAMS_USERS_FILE` con credenciales con hash de `openteams-hash-password`, o un hook `authenticate`); no se usa `--insecure-dev`.
- [ ] `TEAMS_SESSION_SECRET` definido con un valor aleatorio largo, guardado fuera del repositorio, para que las sesiones sobrevivan a los reinicios.
- [ ] TLS termina delante del servidor; la aplicación usa `wss://`. El propio servidor escucha en loopback o en una red privada.
- [ ] `TEAMS_ORIGINS` enumera los orígenes exactos de tu aplicación.
- [ ] `--trust-proxy` solo si el proxy establece `X-Forwarded-For` y los clientes solo pueden llegar al servidor a través de él; si no, déjalo desactivado (todos los clientes tras el proxy compartirán entonces una clave de límite de tasa, lo cual es seguro pero grueso).
- [ ] Los registros de acceso del proxy no guardan cadenas de consulta (los sockets llevan `?token=` y `?session=`), o lo aceptas.
- [ ] `TEAMS_ADMIN_TOKEN` definido si necesitas `POST /auth/revoke`, y mantenido en secreto.
- [ ] El directorio de datos está en un volumen persistente, escribible solo por el usuario del servidor, y tiene copias de seguridad.
- [ ] Límites revisados (`TEAMS_MAX_*`, `TEAMS_WORKSPACE_QUOTA_MB`) y el proxy aplica su propio límite de tasa.
- [ ] Credenciales TURN de corta duración, no un secreto estático en la página.
- [ ] El proceso está supervisado y se detiene con `SIGTERM` para que las salas se guarden.
- [ ] Has leído [Limitaciones](./limitations.md): sin cifrado de extremo a extremo, los miembros de una sala confían entre sí.

## Usar un servidor distinto

Funciona cualquier cosa que cumpla los contratos de `docs/teams-area.md` del núcleo: un servidor `y-websocket` estándar para la sincronización, cualquier difusión por WebSocket para la señalización y cualquier almacén de objetos para los archivos (entrega a `<teams-app>` una función `uploadFile`, o define `uploadFile` en el elemento, y los adjuntos irán allí).
