# Bringen Sie Ihren eigenen Server mit

OpenTeams hat keinen gehosteten Dienst. Der Arbeitsbereich läuft entweder **ohne Server** (`mode: 'local'`,
nur Tabs eines Browsers) oder gegen **einen Server, den Sie betreiben** und der drei kleine Verträge einhält. Die
Referenzimplementierung `openteams-server` erledigt alle drei in einem einzigen Node-Prozess; jeder andere Server, der dieselben Verträge spricht, funktioniert genauso gut.

::: danger Keine Ende-zu-Ende-Verschlüsselung
Der Chat läuft als Yjs-Updates über Ihren Server, der das Raumdokument und jede hochgeladene
Datei im Klartext speichert und lesen kann. Anruf-Medien sind zwischen den Teilnehmern verschlüsselt (DTLS-SRTP), und ein TURN-Relay
kann sie nicht entschlüsseln, die Signalisierung läuft jedoch über den Server.
:::

## Den Referenzserver starten

Erfordert Node 22 oder neuer.

```bash
TEAMS_USERS_FILE=users.json TEAMS_SESSION_SECRET=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
npx openteams-server --insecure-dev     # nur für die lokale Entwicklung: lauscht auf 127.0.0.1:8787, ohne Konten
```

Der Server ist **standardmäßig sicher**: Ohne Kontendatei (`TEAMS_USERS_FILE`), `authenticate`-Hook oder
`authorizeSync`-Hook startet er nicht. `--insecure-dev` (Option `insecureDev: true`) startet ihn trotzdem, mit einer
Warnung, für die lokale Entwicklung. Er wird über die Umgebung konfiguriert, ergänzt um die Flags `--insecure-dev`,
`--trust-proxy[=N]` und `--insecure` (siehe `TEAMS_INSECURE` und `TEAMS_TRUST_PROXY` im
[Bereitstellungsleitfaden](./deploy.md)). `SIGTERM` und `SIGINT` speichern jeden Raum, bevor der Server beendet wird.

| Variable        | Standard    | Bedeutung                                                                                                               |
| --------------- | ----------- | ----------------------------------------------------------------------------------------------------------------------- |
| `PORT`          | `8787`      | Listen-Port.                                                                                                            |
| `HOST`          | `127.0.0.1` | Listen-Adresse. Binden Sie `0.0.0.0` nur hinter TLS.                                                                    |
| `TEAMS_TOKEN`   | keiner      | Falls gesetzt, erwarten Sockets es als `?token=` und Dateianfragen als `Authorization: Bearer`.                         |
| `TEAMS_ORIGINS` | keine       | Kommagetrennte erlaubte `Origin`s für Sockets und Dateien. Ohne Angabe: jede Origin mit Token, ohne Token nur Loopback. |
| `TEAMS_DATA`    | `./data`    | Datenverzeichnis: `rooms/<room>.bin`-Snapshots und `files/<workspace>/`.                                                |

Der Server liest weitere Variablen für Konten (`TEAMS_USERS_FILE`, `TEAMS_SESSION_SECRET`, `TEAMS_ADMIN_TOKEN`, `TEAMS_AUTH_RATE_*`) und für Limits (`TEAMS_INSECURE`, `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP`, `TEAMS_MAX_ROOMS`, `TEAMS_MAX_DOC_MB`, `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`, `TEAMS_FILES_QUOTA_MB`, `TEAMS_WORKSPACE_QUOTA_MB`, `TEAMS_COMPACT_INTERVAL_MIN`, `TEAMS_METRICS_TOKEN`, `TEAMS_LOG_FORMAT`). Die Limits stehen im [Bereitstellungsleitfaden](/de/deploy); die README des Servers ist die vollständige Referenz. Die entsprechenden Optionen von `createTeamsServer` sind `users`, `authenticate`, `sessionSecret`, `adminToken`, `sessionRate`, `authorizeSync`, `maxConnections`, `maxRooms`, `messageRate`, `filesQuotaBytes`, `workspaceQuotaBytes`, `metricsToken`, `compactIntervalMs`, `insecure`, `insecureDev` und `trustProxy`.

Oder binden Sie ihn ein:

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

Verweisen Sie dann den Client darauf:

```ts
const config = {
	mode: 'server',
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	token: 'change-me', // only if TEAMS_TOKEN is set
};
```

[Den Referenzserver bereitstellen](/de/deploy) behandelt TLS, das Token, TURN, Docker und ein Caddy-Beispiel.

## Endpunkte

| Endpunkt                              | Was er tut                                                                                                                                                                                  |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ws /sync/<room>`                     | Zu y-websocket kompatible Synchronisierung: ein `Y.Doc` und Awareness pro Raum, gespeichert in `rooms/<room>.bin`. Nachrichten bis 32 MB.                                                   |
| `ws /signal/<room>`                   | JSON-Relay für die Anruf-Signalisierung: leitet jede Textnachricht an die anderen Sockets des Raums weiter. Nachrichten bis 256 KB.                                                         |
| `POST /files/<workspace>/<name>`      | Speichert einen Anhang (Limit 25 MB) und antwortet mit `201 { url, size }`.                                                                                                                 |
| `GET /files/<workspace>/<name>`       | Lädt ihn als Anhang herunter (`Content-Disposition: attachment`, `Referrer-Policy: no-referrer`, `nosniff`).                                                                                |
| `POST /files/link/<workspace>/<name>` | Liefert `{ url }`: einen für genau diese Datei signierten Link, fünf Minuten gültig, sodass das Token nie in einer URL erscheint.                                                           |
| `GET /files/<scope>`                  | Listet die Dateien eines Geltungsbereichs als `{ files: [{ name, size, modified, uploader }] }`. Ein Bereich ist eine Workspace-ID oder der Raumname eines privaten Chats (nur Mitglieder). |
| `DELETE /files/<scope>/<name>`        | Löscht eine Datei (hochladende Person oder Moderator). `204`, `403` oder `404`.                                                                                                             |
| `GET /metrics`                        | Prometheus-Text, nur mit `TEAMS_METRICS_TOKEN` als Bearer-Token (404, wenn nicht gesetzt).                                                                                                  |
| `GET /health`                         | `{ ok, rooms, calls, auth }`.                                                                                                                                                               |
| `POST /auth/session`                  | Signiert eine Sitzung für `{ user, credential }` (mit Ratenbegrenzung). Siehe "Accounts and sessions" in der README des Servers.                                                            |
| `POST /auth/check`, `/auth/signout`   | Prüfen einer `{ session }` (204 oder 401); Abmelden, was die Sitzungen des Benutzers widerruft.                                                                                             |
| `POST /auth/revoke`                   | Widerruf durch den Betreiber für `{ user }` mit `TEAMS_ADMIN_TOKEN`; ohne Admin-Token ein 404.                                                                                              |

Raumnamen müssen `^[\w-]{1,128}$` entsprechen und Workspace-IDs in Dateipfaden `^[\w-]{1,64}$`; Dateinamen
werden auf Buchstaben, Ziffern, `.`, `-`, `_`, Leerzeichen und Klammern bereinigt und auf 120
Zeichen begrenzt.

## Die Verträge

Jeder Server, der diese drei Dinge tut, funktioniert mit jedem Binding.

### Synchronisierung

`ws(s)://host/<base>/<room>` spricht das y-websocket-Wire-Format (Sync Step 1 und 2, Updates,
Awareness, Query Awareness). Ein unveränderter y-websocket-Server funktioniert. Ein Server, der pro
Raum ein `Y.Doc` hält und es speichert, liefert Verlauf für Personen, die später beitreten.

### Signalisierung

`ws(s)://host/<base>/<room>` leitet jede JSON-Textnachricht an die **anderen** Sockets des Raums weiter
und sendet ihnen `{ "type": "bye", "from": <id> }`, wenn ein Socket schließt. Ein Socket spricht für eine
Peer-ID (das erste `from`, das er verwendet). Der Server muss die Nachrichten nicht verstehen.

### Dateien

Optional. `POST|GET <origin>/files/<workspace>/<name>` mit dem Token in einem `Authorization: Bearer`-Header,
dazu `POST <origin>/files/link/<workspace>/<name>`, das ein kurzlebiges signiertes `{ url }` liefert,
um eine Datei in einem Browser-Tab zu öffnen. Der Origin wird aus `syncUrl` übernommen. Ohne Dateiserver geben Sie
dem Element eine `uploadFile`-Funktion (um Anhänge in Ihrem eigenen Objektspeicher abzulegen), sonst werden Anhänge
nur per Name geteilt.

Dateien haben einen Geltungsbereich. `<scope>` ist die Workspace-ID für Kanaldateien oder der Raumname eines
privaten Chats (`chat-<hash>-<chat id>`), dessen Dateien nur dessen Mitglieder erreichen. Ein Server mit Konten erwartet
bei jeder Dateianfrage die signierte Sitzung im Header `X-Teams-Session` (ein Server ohne Konten liest den
behaupteten Nutzer aus `X-Teams-User`); der Client sendet beides. Ebenfalls optional: `GET <origin>/files/<scope>` zum
Auflisten und `DELETE <origin>/files/<scope>/<name>` für die hochladende Person oder einen Moderator.

Beide Sockets erhalten `?token=<token>`, wenn ein Token konfiguriert ist, weil ein Browser bei einem WebSocket keine
Header setzen kann. Erzwingen Sie es auf dem Server, ebenso wie eine Origin-Allowlist.

## Anrufe: STUN und TURN

`iceServers` enthält STUN- und TURN-URLs. Öffentliches STUN reicht in offenen Netzen und ist hinter symmetrischem NAT
oder strikten Firewalls nie genug: Betreiben Sie Ihr eigenes TURN (zum Beispiel coturn) und geben Sie den Clients
Zugangsdaten. TURN-Einträge ohne Zugangsdaten werden von `parseServerConfig` des Core abgelehnt.
Statische TURN-Zugangsdaten sind für jeden sichtbar, der die App öffnen kann; stellen Sie für alles Öffentliche
kurzlebige aus. Die [Bereitstellungsanleitung](/de/deploy#anrufe-uber-nats-hinweg-turn) enthält ein coturn-Beispiel.

## Was der Referenzserver nicht kann

Machen Sie sich dies klar, bevor Sie Personen darauf arbeiten lassen:

- **Keine Ende-zu-Ende-Verschlüsselung.** Er speichert jede Nachricht und jede Datei und kann sie lesen.
- **Konten sind minimal gehalten.** Mit `users` (oder `TEAMS_USERS_FILE`) oder einer `authenticate`-Funktion signiert der Server Sitzungen (`POST /auth/session`), übernimmt die Benutzer-ID aus der Sitzung und kann sie widerrufen. Ohne sie ist `TEAMS_TOKEN` ein gemeinsames Geheimnis für alle, im Browser gespeichert, und Benutzer-IDs sind Behauptungen. Es gibt kein Zurücksetzen von Passwörtern, keine Kontenverwaltung, kein SSO und keine Zugriffssteuerung pro Team. Autor- und Rollenprüfungen im Client lassen sich von einem feindlichen Client umgehen, daher ist der Server die Vertrauensgrenze: Nutzen Sie `authenticate` mit Ihrem Identitätsanbieter und `authorizeSync` für Rechte pro Raum.
- **Kein eigenes TLS.** Betreiben Sie ihn hinter einem Reverse-Proxy, der TLS terminiert und WebSockets
  hochstuft; Browser verweigern `ws://` von einer `https://`-Seite.
- **Ein Prozess, Räume im Arbeitsspeicher.** Er skaliert nicht horizontal. Ein Raum wird 30 Sekunden
  nach dem Schließen seines letzten Sockets entladen.
- **Ein Snapshot pro Raum.** Er wird beim Laden eines Raums und per Timer (`TEAMS_COMPACT_INTERVAL_MIN`) als einziges
  kompaktes Update neu kodiert, der Verlauf wird aber nie beschnitten, ein aktiver Raum wächst also weiter; die
  Größenobergrenze (`TEAMS_MAX_DOC_MB`) begrenzt ihn.
- **Einfache Dateispeicherung.** Dateien sind nach Workspace und privatem Chat getrennt (nur dessen Mitglieder erreichen
  seine Dateien), können von der hochladenden Person oder einem Moderator aufgelistet und gelöscht werden und werden pro
  Workspace gezählt (`TEAMS_WORKSPACE_QUOTA_MB`). Kanaldateien kann jede angemeldete Person lesen; es gibt keine
  Freigabe pro Datei, keine Versionierung und keinen Virenscan. Ein Upload überschreibt nie eine Datei: ein vergebener
  Name wird unter einem neuen gespeichert. Dateien sind auf je 25 MB begrenzt.
- **Kein Rate Limiting** und kein Medienserver: Anrufe laufen Peer-to-Peer (Mesh).
