# Den eigenen Server betreiben

## Der Referenzserver

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable                                                | Bedeutung                                                                                                                                                            |
| ------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`                                          | Listen-Adresse. Binden Sie `0.0.0.0` nur hinter TLS.                                                                                                                 |
| `TEAMS_TOKEN`                                           | Falls gesetzt, erwarten Sockets es als `?token=` und Dateianfragen als `Authorization: Bearer`.                                                                      |
| `TEAMS_ORIGINS`                                         | Kommagetrennte erlaubte `Origin`s für Sockets und Dateien. Leer: jede Origin, wenn ein Token gesetzt ist; ohne Token nur Loopback-Origins und Clients ohne `Origin`. |
| `TEAMS_DATA`                                            | Datenverzeichnis (Standard `./data`): `rooms/<room>.bin`-Snapshots und `files/`.                                                                                     |
| `TEAMS_INSECURE`                                        | `1` (oder `--insecure`) erlaubt einen Nicht-Loopback-`HOST` ohne Token. Sonst startet der Server nicht.                                                              |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Offene Sockets insgesamt (Standard 500) und je Remote-Adresse (20).                                                                                                  |
| `TEAMS_MAX_ROOMS`                                       | Sync-Räume und Anrufräume (Standard je 1000).                                                                                                                        |
| `TEAMS_MAX_DOC_MB`                                      | Größtes persistiertes Raumdokument in MB (Standard 64); größere Updates werden abgelehnt.                                                                            |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Nachrichten pro Sekunde (Standard 100) und Burst (200) je Socket; ein Socket darüber wird geschlossen.                                                               |
| `TEAMS_FILES_QUOTA_MB`                                  | Gesamtgröße des Dateiverzeichnisses in MB (Standard 1024); weitere Uploads erhalten 507.                                                                             |

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

## Sicherheitsmodell

- **Ein gemeinsames Token öffnet jeden Raum und jede Datei.** Es gibt keine Rechte pro Benutzer oder
  Raum. Behandeln Sie das Token wie ein Team-Passwort.
- **Alle Mitglieder eines Raums vertrauen einander.** Der Server bewegt nur Yjs-Bytes, daher kann jedes
  Mitglied Nachrichten anderer bearbeiten oder löschen und Autoren-IDs und Rollen fälschen. Laden Sie
  nur Personen ein, denen Sie den ganzen Raum anvertrauen.
- **Das Token wandert in der WebSocket-Query-Zeichenfolge** (`?token=`), die Proxys und Zugriffsprotokolle
  aufzeichnen und Browser im Verlauf behalten können. Betreiben Sie den Server hinter TLS und
  protokollieren Sie keine Query-Zeichenfolgen. Dateien nehmen das Token nie in einer URL: sie nutzen
  den `Authorization`-Header oder einen signierten Link.
- **Die Standardwerte sind lokal.** Ohne Token lehnt der Server einen Nicht-Loopback-`HOST` ab (außer mit
  `TEAMS_INSECURE=1`) und akzeptiert nur Loopback-Origins, sodass eine besuchte Website keinen Server auf
  Ihrem eigenen Rechner steuern kann. Setzen Sie für jede gemeinsame Bereitstellung `TEAMS_TOKEN` und
  `TEAMS_ORIGINS`.
- **Verwenden Sie kurzlebige TURN-Zugangsdaten** (siehe unten); statische sind für jeden sichtbar, der die App öffnet.
- **Grenzwerte** für Verbindungen, Räume, Dokumentgröße, Nachrichtenrate und Dateispeicher (die Variablen
  oben) begrenzen Missbrauch, ersetzen aber nicht das eigene Rate-Limiting eines Reverse-Proxys.
- Ein Socket kann nur für die Awareness-IDs und die Anruf-Peer-ID sprechen, die er selbst angekündigt
  hat, und ein adressiertes Anrufsignal wird nur an diesen Peer zugestellt. Ein Upload überschreibt nie
  eine Datei: ein bereits vergebener Name wird unter einem neuen gespeichert, den die Antwort in `url`
  zurückgibt.

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
