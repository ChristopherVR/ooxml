# Bringen Sie Ihren eigenen Server mit

OpenTeams hat keinen gehosteten Dienst. Der Arbeitsbereich läuft entweder **ohne Server** (`mode: 'local'`,
nur Tabs eines Browsers) oder gegen **einen Server, den Sie betreiben** und der drei kleine Verträge einhält. Die
Referenzimplementierung `openteams-server` erledigt alle drei in einem einzigen Node-Prozess mit etwa 300
Zeilen; jeder andere Server, der dieselben Verträge spricht, funktioniert genauso gut.

::: danger Keine Ende-zu-Ende-Verschlüsselung
Der Chat läuft als Yjs-Updates über Ihren Server, der das Raumdokument und jede hochgeladene
Datei im Klartext speichert und lesen kann. Anruf-Medien sind zwischen den Teilnehmern verschlüsselt (DTLS-SRTP), und ein TURN-Relay
kann sie nicht entschlüsseln, die Signalisierung läuft jedoch über den Server.
:::

## Den Referenzserver starten

Erfordert Node 22 oder neuer.

```bash
npx openteams-server                    # listens on 127.0.0.1:8787
TEAMS_TOKEN=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
```

Er wird ausschließlich über die Umgebung konfiguriert; der Befehl kennt keine Optionen.

| Variable        | Standard     | Bedeutung                                                                                       |
| --------------- | ------------ | ----------------------------------------------------------------------------------------------- |
| `PORT`          | `8787`       | Listen-Port.                                                                                    |
| `HOST`          | `127.0.0.1`  | Listen-Adresse. Binden Sie `0.0.0.0` nur hinter TLS.                                            |
| `TEAMS_TOKEN`   | keiner       | Falls gesetzt, erwarten Sockets es als `?token=` und Dateianfragen als `Authorization: Bearer`. |
| `TEAMS_ORIGINS` | keine (alle) | Kommagetrennte erlaubte `Origin`s für Sockets und Dateien.                                      |
| `TEAMS_DATA`    | `./data`     | Datenverzeichnis: `rooms/<room>.bin`-Snapshots und `files/<workspace>/`.                        |

Oder binden Sie ihn ein:

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({
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

| Endpunkt                              | Was er tut                                                                                                                                |
| ------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `ws /sync/<room>`                     | Zu y-websocket kompatible Synchronisierung: ein `Y.Doc` und Awareness pro Raum, gespeichert in `rooms/<room>.bin`. Nachrichten bis 32 MB. |
| `ws /signal/<room>`                   | JSON-Relay für die Anruf-Signalisierung: leitet jede Textnachricht an die anderen Sockets des Raums weiter. Nachrichten bis 256 KB.       |
| `POST /files/<workspace>/<name>`      | Speichert einen Anhang (Limit 25 MB) und antwortet mit `201 { url, size }`.                                                               |
| `GET /files/<workspace>/<name>`       | Lädt ihn als Anhang herunter (`Content-Disposition: attachment`, `Referrer-Policy: no-referrer`, `nosniff`).                              |
| `POST /files/link/<workspace>/<name>` | Liefert `{ url }`: einen für genau diese Datei signierten Link, fünf Minuten gültig, sodass das Token nie in einer URL erscheint.         |
| `GET /health`                         | `{ ok, rooms, calls, auth }`.                                                                                                             |

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
- **Keine Benutzerkonten.** `TEAMS_TOKEN` ist ein gemeinsames Geheimnis für alle, im Browser gespeichert. Es
  hält Fremde vom Server eines kleinen Teams fern; es ist keine Authentifizierung pro Benutzer. Autor- und Rollenprüfungen
  im Client lassen sich von einem feindlichen Client umgehen, daher ist der Server die Vertrauensgrenze:
  Schalten Sie für den echten Einsatz Ihre eigene Authentifizierung davor (ein Proxy, der eine Sitzung prüft).
- **Kein eigenes TLS.** Betreiben Sie ihn hinter einem Reverse-Proxy, der TLS terminiert und WebSockets
  hochstuft; Browser verweigern `ws://` von einer `https://`-Seite.
- **Ein Prozess, Räume im Arbeitsspeicher.** Er skaliert nicht horizontal. Ein Raum wird 30 Sekunden
  nach dem Schließen seines letzten Sockets entladen.
- **Ein Snapshot pro Raum, keine Verdichtung.** Das Dokument wächst mit dem Verlauf und wird nie bereinigt.
- **Einfache Dateispeicherung.** Kein Auflisten, kein Löschen, kein Kontingent pro Workspace; das Hochladen einer Datei mit
  einem vorhandenen Namen ersetzt sie. Dateien sind auf je 25 MB begrenzt.
- **Kein Rate Limiting** und kein Medienserver: Anrufe laufen Peer-to-Peer (Mesh).
