# Den eigenen Server betreiben

## Der Referenzserver

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable        | Bedeutung                                                                                       |
| --------------- | ----------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`  | Listen-Adresse. Binden Sie `0.0.0.0` nur hinter TLS.                                            |
| `TEAMS_TOKEN`   | Falls gesetzt, erwarten Sockets es als `?token=` und Dateianfragen als `Authorization: Bearer`. |
| `TEAMS_ORIGINS` | Kommagetrennte erlaubte `Origin`s für Sockets und Dateien. Leer erlaubt alle.                   |
| `TEAMS_DATA`    | Datenverzeichnis (Standard `./data`): `rooms/<room>.bin`-Snapshots und `files/`.                |

Er stellt `ws /sync/<room>` (kompatibel mit y-websocket, ein gespeichertes `Y.Doc` pro Raum),
`ws /signal/<room>` (ein JSON-Relay) und `POST|GET /files/<workspace>/<name>` (Limit 25 MB, Namen
bereinigt, Downloads als Anhang mit `Referrer-Policy: no-referrer`) bereit. Mit einem Token öffnet ein
Browser Dateien über `POST /files/link/<workspace>/<name>` (Token im Header), das
einen für genau diese Datei signierten, fünf Minuten gültigen Link liefert. `GET /health` meldet die Anzahl von Räumen und Anrufen.

**Stellen Sie ihn hinter TLS.** Browser verweigern `ws://` von einer `https://`-Seite, und ein Token oder die Anruf-
Signalisierung über einfaches HTTP ist im Netzwerk lesbar. Jeder Reverse-Proxy, der WebSockets hochstuft,
funktioniert (nginx, Caddy, Traefik). Caddy-Beispiel:

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

Verweisen Sie die App dann auf `wss://teams.example.com/sync` und `wss://teams.example.com/signal`.

**Das Token ist ein gemeinsames Geheimnis**, das im Browser gespeichert wird: Es hält Fremde vom Server eines kleinen Teams fern,
es ist keine Authentifizierung pro Benutzer. Für echte Konten schalten Sie Ihre eigene Authentifizierung davor (ein Proxy,
der eine Sitzung prüft, oder ersetzen Sie `tokenOk` durch eine Token-Prüfung) und behalten dieselben drei
Verträge bei.

## Anrufe über NATs hinweg: TURN

Anrufe laufen Peer-to-Peer. Öffentliches STUN funktioniert in offenen Netzen; hinter symmetrischem NAT oder strikten
Firewalls benötigen Sie ein TURN-Relay, das Medien ebenfalls nur als Chiffretext (DTLS-SRTP) transportiert. Ein minimales
coturn (`turnserver.conf`):

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

und in den Servereinstellungen der App, ein ICE-Server pro Zeile:

```
stun:turn.example.com:3478
turn:turn.example.com:3478,turns:turn.example.com:5349 teams CHANGE_ME
```

Statische TURN-Zugangsdaten sind für jeden sichtbar, der die App öffnen kann. Stellen Sie für alles Öffentliche
kurzlebige Zugangsdaten aus (coturn `use-auth-secret` plus ein kleiner Endpunkt, der sie signiert) und übergeben Sie
sie in den `iceServers`, die Sie dem Element geben.

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

Starten Sie ihn mit gesetzten `TEAMS_TOKEN` und `TEAMS_ORIGINS`, hinter Ihrem TLS-Proxy.

## Einen anderen Server verwenden

Alles, was die Verträge aus `docs/teams-area.md` des Core einhält, funktioniert: ein unveränderter
`y-websocket`-Server für die Synchronisierung, ein beliebiger WebSocket-Fan-out für die Signalisierung, ein beliebiger Objektspeicher für Dateien (geben Sie
`<teams-app>` eine `uploadFile`-Funktion oder setzen Sie `uploadFile` am Element, und Anhänge gehen
dorthin).
