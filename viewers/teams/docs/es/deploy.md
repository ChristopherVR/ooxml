# Ejecutar tu propio servidor

## El servidor de referencia

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable        | Significado                                                                                                    |
| --------------- | -------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`  | Dirección de escucha. Enlaza `0.0.0.0` solo detrás de TLS.                                                     |
| `TEAMS_TOKEN`   | Si se define, los sockets lo reciben como `?token=` y las peticiones de archivos como `Authorization: Bearer`. |
| `TEAMS_ORIGINS` | `Origin` permitidos, separados por comas, para sockets y archivos. Vacío permite cualquiera.                   |
| `TEAMS_DATA`    | Directorio de datos (por defecto `./data`): instantáneas `rooms/<room>.bin` y `files/`.                        |

Sirve `ws /sync/<room>` (compatible con y-websocket, un `Y.Doc` persistido por sala), `ws /signal/<room>` (un relé JSON) y `POST|GET /files/<workspace>/<name>` (límite de 25 MB, nombres depurados, descargas enviadas como adjuntos con `Referrer-Policy: no-referrer`). Con un token, el navegador abre los archivos mediante `POST /files/link/<workspace>/<name>` (token en la cabecera), que devuelve un enlace firmado para ese único archivo y válido durante cinco minutos. `GET /health` informa del número de salas y de llamadas.

**Ponlo detrás de TLS.** Los navegadores rechazan `ws://` desde una página `https://`, y un token o la señalización de llamadas sobre HTTP simple pueden leerse en la red. Sirve cualquier proxy inverso que actualice a WebSocket (nginx, Caddy, Traefik). Ejemplo con Caddy:

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

Después apunta la aplicación a `wss://teams.example.com/sync` y `wss://teams.example.com/signal`.

**El token es un secreto compartido** almacenado en el navegador: mantiene fuera a los extraños del servidor de un equipo pequeño, no es autenticación por usuario. Para cuentas reales, coloca delante tu propia autenticación (un proxy que compruebe una sesión, o sustituye `tokenOk` por una verificación de tokens) y conserva los mismos tres contratos.

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
COPY server/index.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data
VOLUME /data
EXPOSE 8787
CMD ["node", "index.mjs"]
```

Ejecútalo con `TEAMS_TOKEN` y `TEAMS_ORIGINS` definidos, detrás de tu proxy TLS.

## Usar un servidor distinto

Funciona cualquier cosa que cumpla los contratos de `docs/teams-area.md` del núcleo: un servidor `y-websocket` estándar para la sincronización, cualquier difusión por WebSocket para la señalización y cualquier almacén de objetos para los archivos (entrega a `<teams-app>` una función `uploadFile`, o define `uploadFile` en el elemento, y los adjuntos irán allí).
