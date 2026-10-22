# Ejecutar tu propio servidor

## El servidor de referencia

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable                                                | Significado                                                                                                                                                                      |
| ------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`                                          | Dirección de escucha. Enlaza `0.0.0.0` solo detrás de TLS.                                                                                                                       |
| `TEAMS_TOKEN`                                           | Si se define, los sockets lo reciben como `?token=` y las peticiones de archivos como `Authorization: Bearer`.                                                                   |
| `TEAMS_ORIGINS`                                         | `Origin` permitidos, separados por comas, para sockets y archivos. Vacío: cualquier origen si hay token; sin token, solo orígenes de loopback y clientes que no envían `Origin`. |
| `TEAMS_DATA`                                            | Directorio de datos (por defecto `./data`): instantáneas `rooms/<room>.bin` y `files/`.                                                                                          |
| `TEAMS_INSECURE`                                        | `1` (o `--insecure`) permite un `HOST` que no sea loopback sin token. Si no, el servidor se niega a arrancar.                                                                    |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Sockets abiertos en total (por defecto 500) y por dirección remota (20).                                                                                                         |
| `TEAMS_MAX_ROOMS`                                       | Salas de sincronización y de llamada (por defecto 1000 cada una).                                                                                                                |
| `TEAMS_MAX_DOC_MB`                                      | Mayor documento de sala persistido en MB (por defecto 64); las actualizaciones que lo superen se rechazan.                                                                       |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Mensajes por segundo (por defecto 100) y ráfaga (200) por socket; un socket que los supere se cierra.                                                                            |
| `TEAMS_FILES_QUOTA_MB`                                  | Tamaño total del directorio de archivos en MB (por defecto 1024); las subidas que lo superen reciben 507.                                                                        |

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
- **El token viaja en la cadena de consulta del WebSocket** (`?token=`), que los proxies y los
  registros de acceso pueden guardar y los navegadores conservar en el historial. Ejecuta el servidor
  detrás de TLS y no registres las cadenas de consulta. Los archivos nunca llevan el token en una URL:
  usan la cabecera `Authorization` o un enlace firmado.
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
  server/rate-limit.mjs server/chat-access.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data
VOLUME /data
EXPOSE 8787
CMD ["node", "bin.mjs"]
```

Constrúyelo desde `viewers/teams`. Copia todos los módulos que importa el punto de entrada (los `files` de `server/package.json`); copiar solo `index.mjs` falla por un módulo ausente. La receta se comprobó ejecutando esos archivos con `npm install --omit=dev` en Node; la imagen en sí no se construyó.

Ejecútalo con `TEAMS_TOKEN` y `TEAMS_ORIGINS` definidos, detrás de tu proxy TLS.

## Usar un servidor distinto

Funciona cualquier cosa que cumpla los contratos de `docs/teams-area.md` del núcleo: un servidor `y-websocket` estándar para la sincronización, cualquier difusión por WebSocket para la señalización y cualquier almacén de objetos para los archivos (entrega a `<teams-app>` una función `uploadFile`, o define `uploadFile` en el elemento, y los adjuntos irán allí).
