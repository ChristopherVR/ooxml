# Den eigenen Server betreiben

## Der Referenzserver

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable                                                | Bedeutung                                                                                                                                                                                                                                                                     |
| ------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`                                          | Listen-Adresse. Binden Sie `0.0.0.0` nur hinter TLS.                                                                                                                                                                                                                          |
| `TEAMS_TOKEN`                                           | Falls gesetzt, erwarten Sockets es als `?token=` und Dateianfragen als `Authorization: Bearer`.                                                                                                                                                                               |
| `TEAMS_ORIGINS`                                         | Kommagetrennte erlaubte `Origin`s für Sockets und Dateien. Leer: jede Origin, wenn ein Token gesetzt ist; ohne Token nur Loopback-Origins und Clients ohne `Origin`.                                                                                                          |
| `TEAMS_DATA`                                            | Datenverzeichnis (Standard `./data`): `rooms/<room>.bin`-Snapshots und `files/`.                                                                                                                                                                                              |
| `TEAMS_TRUST_PROXY`, `--trust-proxy[=N]`                | Schlüsselt den Auth-Ratenbegrenzer und die Verbindungsgrenze pro Adresse auf `X-Forwarded-For` (den Eintrag des N-ten vertrauenswürdigen Proxys, Standard 1). Standardmäßig aus; nur hinter einem Proxy aktivieren, der den Header setzt, sonst kann ein Client ihn fälschen. |
| `--insecure-dev`                                        | Startet ohne Konten und ohne `authenticate`-Hook (Option `insecureDev`). Ohne diese Option startet der Server nur, wenn `TEAMS_USERS_FILE` gesetzt ist. Nur für die Entwicklung.                                                                                              |
| `TEAMS_INSECURE`                                        | `1` (oder `--insecure`) erlaubt einen Nicht-Loopback-`HOST` ohne Token. Sonst startet der Server nicht.                                                                                                                                                                       |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Offene Sockets insgesamt (Standard 500) und je Remote-Adresse (20).                                                                                                                                                                                                           |
| `TEAMS_MAX_ROOMS`                                       | Sync-Räume und Anrufräume (Standard je 1000).                                                                                                                                                                                                                                 |
| `TEAMS_MAX_DOC_MB`                                      | Größtes persistiertes Raumdokument in MB (Standard 64); größere Updates werden abgelehnt.                                                                                                                                                                                     |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Nachrichten pro Sekunde (Standard 100) und Burst (200) je Socket; ein Socket darüber wird geschlossen.                                                                                                                                                                        |
| `TEAMS_FILES_QUOTA_MB`                                  | Optionale Obergrenze für das gesamte Dateiverzeichnis in MB (standardmäßig aus); weitere Uploads erhalten 507.                                                                                                                                                                |
| `TEAMS_WORKSPACE_QUOTA_MB`                              | Dateien pro Workspace in MB (Standard 1024; ein privater Chat zählt zu seinem Workspace); weitere Uploads erhalten 507.                                                                                                                                                       |
| `TEAMS_COMPACT_INTERVAL_MIN`                            | Minuten zwischen den Verdichtungen der Raum-Snapshots (Standard 360, `0` schaltet den Timer ab). Ein Raum wird außerdem beim Laden verdichtet.                                                                                                                                |
| `TEAMS_METRICS_TOKEN`                                   | Aktiviert `GET /metrics` (Prometheus-Text), nur mit `Authorization: Bearer <dieses Token>`. Ohne Token aus.                                                                                                                                                                   |
| `TEAMS_LOG_FORMAT`, `--log-format=json`                 | `text` (Standard) oder `json`: ein JSON-Objekt pro Zeile mit `time`, `level` und `msg`.                                                                                                                                                                                       |

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
es ist keine Authentifizierung pro Benutzer. Für echte Konten geben Sie dem Server `users` (oder `TEAMS_USERS_FILE`) oder eine
`authenticate`-Funktion für Ihren Identitätsanbieter und setzen `TEAMS_SESSION_SECRET`, damit Sitzungen Neustarts überstehen
(siehe "Accounts and sessions" in der README des Servers). Die drei Verträge bleiben dieselben.

## Sicherheitsmodell

- **Ein gemeinsames Token öffnet jeden Raum und jede Datei.** Identität pro Benutzer gibt es nur mit aktivierten Konten und Rechte pro Raum nur mit `authorizeSync`. Behandeln Sie das Token wie ein Team-Passwort.
- **Alle Mitglieder eines Raums vertrauen einander.** Der Server bewegt nur Yjs-Bytes, daher kann jedes
  Mitglied Nachrichten anderer bearbeiten oder löschen und Autoren-IDs und Rollen fälschen. Laden Sie
  nur Personen ein, denen Sie den ganzen Raum anvertrauen.
- **Token und Sitzung wandern im `Sec-WebSocket-Protocol`-Header** (`openteams.session.<session>`,
  `openteams.token.<token>`, neben einem einfachen `openteams`), der nicht in URLs oder Protokollen landet. Der
  mitgelieferte Client weicht nur auf die Query-Zeichenfolge aus, wenn ein Wert Zeichen enthält, die ein Subprotokoll
  nicht tragen kann. Betreiben Sie den Server hinter TLS. Dateien nehmen das Token nie in einer URL: sie nutzen den
  `Authorization`-Header oder einen signierten Link.
- **Autorenschaft wird durchgesetzt** für angemeldete Benutzer: niemand kann als eine andere Person schreiben, deren
  Nachrichten bearbeiten oder löschen oder Kanäle und Tabs ändern, die er nicht erstellt hat, außer das Konto hat
  `"role": "moderator"` in der Benutzerdatei. `TEAMS_ENFORCE_AUTHORSHIP=0` schaltet das ab.
- **Sicher per Standard.** Ohne Konten, einen `authenticate`-Hook oder `authorizeSync` startet der Server nicht;
  `--insecure-dev` ist die ausdrückliche Abweichung, die mit einer Warnung quittiert wird. Ein gemeinsames Token
  allein ist keine Authentifizierung.
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
COPY server/bin.mjs server/index.mjs server/auth.mjs server/passwords.mjs server/guards.mjs \
  server/rate-limit.mjs server/chat-access.mjs server/validator.generated.mjs server/files.mjs server/metrics.mjs server/room-store.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data TEAMS_USERS_FILE=/config/users.json
RUN mkdir /data && chown node:node /data
VOLUME /data
EXPOSE 8787
USER node
CMD ["node", "bin.mjs"]
```

Bauen Sie es in `viewers/teams`. Es kopiert jedes Modul, das der Einstiegspunkt importiert (die `files` aus `server/package.json`); nur `index.mjs` zu kopieren scheitert mit einem fehlenden Modul. Das Rezept wurde durch Ausführen dieser Dateien mit `npm install --omit=dev` unter Node geprüft, das Image selbst wurde nicht gebaut.

Starten Sie ihn mit eingehängter Kontendatei, gesetztem `TEAMS_SESSION_SECRET` und `TEAMS_ORIGINS`, hinter Ihrem TLS-Proxy:

```bash
docker run -d --name openteams -p 127.0.0.1:8787:8787 \
  -v openteams-data:/data -v "$PWD/users.json:/config/users.json:ro" \
  -e TEAMS_SESSION_SECRET=change-me -e TEAMS_ORIGINS=https://app.example.com \
  openteams-server
```

Ohne eingehängte Kontendatei beendet sich der Container mit der Meldung „refusing to start“; fügen Sie
`--insecure-dev` keinem Container hinzu, der aus einem Netzwerk erreichbar ist. Fügen Sie `--trust-proxy` dem
Befehl hinzu (`CMD ["node", "bin.mjs", "--trust-proxy"]`), aber nur, wenn ein Reverse-Proxy `X-Forwarded-For`
setzt. `docker stop` sendet `SIGTERM`, worauf der Server zuerst jeden Raum speichert (Docker wartet standardmäßig
zehn Sekunden, das genügt). Das Image legt `/data` im Besitz von `node` an; bei einem Bind-Mount setzen Sie das
Host-Verzeichnis mit `chown` auf uid 1000.

## Betrieb

### Sicherung und Wiederherstellung

Alles, was der Server speichert, liegt unter `TEAMS_DATA` (Standard `./data`):

| Pfad                       | Enthält                                                                                                        |
| -------------------------- | -------------------------------------------------------------------------------------------------------------- |
| `rooms/<room>.bin`         | Ein Yjs-Snapshot pro Raum: alle Chats, Kanäle, Kalender und Tab-Daten.                                         |
| `files/<workspace>/`       | Hochgeladene Dateien.                                                                                          |
| `file-meta/<scope>.json`   | Wer welche Datei hochgeladen hat (bestimmt, wer sie löschen darf). Zusammen mit `files/` sichern.              |
| `session-revocations.json` | Widerrufsgenerationen der Sitzungen (fehlt sie, werden widerrufene Sitzungen wieder gültig, bis sie ablaufen). |

Bewahren Sie außerdem auf, was nicht im Datenverzeichnis liegt: Ihre Kontendatei (`TEAMS_USERS_FILE`) und
`TEAMS_SESSION_SECRET`. Ein geändertes oder verlorenes Geheimnis meldet alle ab, was harmlos ist.

Zum Sichern kopieren Sie das ganze Datenverzeichnis. Raum-Snapshots werden in eine temporäre Datei geschrieben und
umbenannt, sodass eine Kopie nie einen halb geschriebenen Snapshot sieht. Eine Kopie, die während des Betriebs
entsteht, kann die Änderungen der letzten halben Sekunde verpassen (Schreibvorgänge sind um 500 ms verzögert) und
eine Datei erfassen, die neuer ist als die Nachbarn ihres Raums. Für eine konsistente Kopie stoppen Sie den Server
zuerst (`SIGTERM` speichert vor dem Beenden jeden Raum), kopieren und starten ihn wieder; oder Sie erstellen einen
Snapshot des Volumes auf Dateisystemebene.

Zum Wiederherstellen stoppen Sie den Server, ersetzen das Datenverzeichnis durch die Sicherung (`rooms/`, `files/`
und `session-revocations.json` bleiben zusammen), stellen die Kontendatei und das Geheimnis wieder her und starten
ihn. Verbundene Clients verbinden sich neu und führen ihre Daten zusammen: Yjs behält, was sie halten und der
Sicherung fehlt, sodass eine Wiederherstellung keine Arbeit verwirft, die noch in einem Browser offen ist.

### Dateien, Verdichtung und Metriken

- **Dateiberechtigungen.** Eine Datei liegt in einem Geltungsbereich: der Workspace-ID für Kanaldateien (jede
  angemeldete Person des Servers) oder dem Raumnamen eines privaten Chats (`chat-<hash>-<chat id>`), den nur dessen
  Mitglieder auflisten, lesen, befüllen oder verlinken dürfen. Der Server prüft die Mitgliederliste des Chats wie bei
  dessen Sync-Raum. Mit Konten muss jede Dateianfrage die signierte Sitzung in `X-Teams-Session` tragen (der Client
  tut das); ein signierter Download-Link wird erst nach dieser Prüfung ausgestellt und öffnet fünf Minuten lang
  genau seine Datei. Mit `--insecure-dev` gilt der behauptete `X-Teams-User`, wie jede Angabe dort. Kanaldateien
  sind nicht pro Kanal eingeschränkt: alle Nutzer des Servers können sie lesen.
- **Auflisten und Löschen.** `GET /files/<scope>` antwortet mit `{ files: [{ name, size, modified, uploader }] }`.
  `DELETE /files/<scope>/<name>` dürfen die hochladende Person oder ein Moderator (`204`, sonst `403`). Vor dieser
  Version hochgeladene Dateien haben keinen erfassten Uploader; nur ein Moderator kann sie löschen.
- **Bereinigung.** Ein Upload wird in eine verborgene `.part`-Datei geschrieben, die ein fehlgeschlagener oder
  abgebrochener Upload entfernt. Beim Start löscht der Server alle übrig gebliebenen `.part`-Dateien (und veraltete
  temporäre Raumdateien), danach alle 15 Minuten jene, die älter als eine Stunde sind. Hochgeladene, aber nie an eine
  Nachricht angehängte Dateien bleiben, bis sie jemand löscht.
- **Kontingente.** Das Limit gilt pro Workspace (`TEAMS_WORKSPACE_QUOTA_MB`, Standard 1024) und zählt die Kanaldateien
  und die Dateien seiner privaten Chats; Löschen gibt Platz frei. `TEAMS_FILES_QUOTA_MB` ist eine optionale Obergrenze
  für das ganze Verzeichnis.
- **Verdichtung.** Beim Laden eines Raums und alle `TEAMS_COMPACT_INTERVAL_MIN` Minuten für jeden Raum kodiert der
  Server den Raum als einziges Yjs-Update auf einem frischen Dokument neu und schreibt es nur, wenn es kleiner ist und
  exakt denselben Zustand beschreibt. Geschrieben wird in eine temporäre Datei, auf den Datenträger geschrieben und über
  den Snapshot umbenannt; ein Absturz hinterlässt den alten oder den neuen Snapshot, nie einen halben. Unterbrochene
  temporäre Dateien werden beim Start entfernt. Sicherungen vor und nach einer Verdichtung lassen sich beide wiederherstellen.
- **Metriken.** Setzen Sie `TEAMS_METRICS_TOKEN` und rufen Sie `GET /metrics` mit `Authorization: Bearer <token>` ab:
  `teams_rooms_open`, `teams_connections{kind}`, `teams_stored_bytes{kind="rooms|files"}`,
  `teams_updates_rejected_total{code}`, `teams_auth_failures_total{kind}`, `teams_rate_limited_total{scope}` und
  `teams_room_compactions_total`. Ohne Token antwortet der Endpunkt mit 404. Nicht ins öffentliche Internet stellen.
- **Logs.** `--log-format=json` (oder `TEAMS_LOG_FORMAT=json`) schreibt ein JSON-Objekt pro Zeile: `{ "time", "level", "msg" }`.

### Aktualisieren

1. Lesen Sie die `CHANGELOG.md` des Servers für den Versionsbereich, den Sie überschreiten; Breaking Changes sind markiert.
2. Sichern Sie das Datenverzeichnis.
3. Stoppen Sie den Server (`SIGTERM`), installieren Sie die neue Version (`npx openteams-server@<version>` oder das Image neu bauen) und starten Sie ihn. Raum-Snapshots sind einfache Yjs-Updates; ältere werden unverändert gelesen.
4. Prüfen Sie `GET /health` und öffnen Sie einen Raum.

Aktualisierung von einer Version vor dem sicheren Standard: Ein Server, der ohne Konten startete, startet jetzt nicht
mehr. Fügen Sie Konten (`TEAMS_USERS_FILE`), einen `authenticate`-Hook oder `authorizeSync` hinzu, oder übergeben Sie
`--insecure-dev`, wenn (und nur wenn) es ein lokaler Entwicklungsserver ist. Ein gemeinsames `TEAMS_TOKEN` allein
startet den Server nicht mehr.

### Checkliste für den Produktionsbetrieb

- [ ] Konten eingerichtet (`TEAMS_USERS_FILE` mit gehashten Zugangsdaten aus `openteams-hash-password` oder ein `authenticate`-Hook); `--insecure-dev` wird nicht verwendet.
- [ ] `TEAMS_SESSION_SECRET` auf einen langen Zufallswert gesetzt, außerhalb des Repositorys aufbewahrt, damit Sitzungen Neustarts überstehen.
- [ ] TLS endet vor dem Server; die App nutzt `wss://`. Der Server selbst lauscht auf Loopback oder einem privaten Netz.
- [ ] `TEAMS_ORIGINS` listet die genauen Origins Ihrer App.
- [ ] `--trust-proxy` nur gesetzt, wenn der Proxy `X-Forwarded-For` setzt und Clients den Server nur über ihn erreichen können; sonst aus lassen (alle Clients teilen sich dann hinter dem Proxy einen Rate-Limit-Schlüssel, was sicher, aber grob ist).
- [ ] Proxy-Zugriffsprotokolle zeichnen keine Query-Zeichenfolgen auf (Sockets tragen `?token=` und `?session=`), oder Sie akzeptieren das.
- [ ] `TEAMS_ADMIN_TOKEN` gesetzt, falls Sie `POST /auth/revoke` brauchen, und geheim gehalten.
- [ ] Das Datenverzeichnis liegt auf einem dauerhaften Volume, ist nur für den Benutzer des Servers schreibbar und wird gesichert.
- [ ] Grenzwerte geprüft (`TEAMS_MAX_*`, `TEAMS_WORKSPACE_QUOTA_MB`), und der Proxy wendet eigenes Rate-Limiting an.
- [ ] Kurzlebige TURN-Zugangsdaten, kein statisches Geheimnis auf der Seite.
- [ ] Der Prozess wird überwacht und mit `SIGTERM` beendet, damit Räume gespeichert werden.
- [ ] Sie haben die [Einschränkungen](./limitations.md) gelesen: keine Ende-zu-Ende-Verschlüsselung, Raummitglieder vertrauen einander.

## Einen anderen Server verwenden

Alles, was die Verträge aus `docs/teams-area.md` des Core einhält, funktioniert: ein unveränderter
`y-websocket`-Server für die Synchronisierung, ein beliebiger WebSocket-Fan-out für die Signalisierung, ein beliebiger Objektspeicher für Dateien (geben Sie
`<teams-app>` eine `uploadFile`-Funktion oder setzen Sie `uploadFile` am Element, und Anhänge gehen
dorthin).
