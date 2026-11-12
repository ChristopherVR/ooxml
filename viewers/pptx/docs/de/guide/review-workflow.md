---
title: Prüfung in Ihrer Anwendung
description: Ein Rezept für die gemeinsame Prüfung von Präsentationen in einer Host-Anwendung, die Konten, Berechtigungen, die Sitzung, die Speicherung und den KI-Schlüssel verwaltet.
---

# Prüfung in Ihrer Anwendung

Zusammenarbeit ist meist dort am nützlichsten, wo Ihre Benutzer bereits arbeiten: an einem Deal-Desk, in einer Compliance-Warteschlange, in einem CRM-Datensatz oder auf einem Projektboard. Mehrere Personen öffnen dieselbe Präsentation aus diesem System, nehmen kleine Änderungen von Hand oder mit dem KI-Assistenten vor, und das System hält fest, wer was tun darf und welche Version aktuell ist.

Dieses Rezept teilt die Arbeit an einer Linie auf:

| Ihre Anwendung verwaltet                                   | Der Viewer übernimmt                                                   |
| ---------------------------------------------------------- | ---------------------------------------------------------------------- |
| Anmeldung, Konten sowie Anzeigename und Farbe der Person   | Rendern und Bearbeiten der Präsentation                                |
| Wer jedes Element öffnen, bearbeiten oder nur ansehen darf | Das Folienschema im gemeinsamen Yjs-Dokument                           |
| Das Yjs-Dokument, Awareness, Provider und Relay            | Präsenz (Cursor, Auswahl, "einem Mitwirkenden folgen")                 |
| Wann die Sitzung bereit ist und wann nicht mehr            | Zurückhalten von Änderungen, bis Sie den Raum als bereit melden        |
| Wohin gespeicherte Versionen gehen und wer sie schreibt    | Umwandlung der gemeinsamen Präsentation in PPTX-Bytes für den Besitzer |
| Das KI-Modell, sein Schlüssel und wer es nutzen darf       | Das Assistenten-Panel, seine Präsentationswerkzeuge und die Vorschläge |

Die Beispiel-App in
[`demos/pptx/demo-host-review`](https://github.com/ChristopherVR/ooxml/tree/main/demos/pptx/demo-host-review)
setzt jeden der folgenden Schritte mit der Vanilla-Bindung um. Jede Bindung akzeptiert dieselben Optionen `collaboration` und `ai`, daher sind die Schritte in React, Vue, Angular und Svelte gleich.

## 1. Berechtigungen einer Rolle zuordnen {#_1-map-your-permissions-to-a-role}

Behalten Sie Ihr eigenes Berechtigungsmodell. Leiten Sie beim Einhängen des Editors daraus die `role` des Viewers ab:

| Ihre Berechtigung (Beispiel) | `role`            | Hinweise                                                          |
| ---------------------------- | ----------------- | ----------------------------------------------------------------- |
| Besitzer des Elements        | `'owner'`         | Bearbeitet und erhält `onWriteBack`-Snapshots zum Speichern       |
| Prüfer                       | `'collaborator'`  | Bearbeitet live zusammen mit den anderen                          |
| Beobachter                   | `'viewer'`        | Schreibgeschützt; übergeben Sie außerdem `editable: false`        |
| Kein Zugriff                 | (nicht einhängen) | Ihre API sollte nicht einmal verraten, dass das Element existiert |

`role: 'viewer'` wird **nur im Browser** durchgesetzt. Setzen Sie es zusätzlich auf dem Server durch (Schritt 3).

## 2. Die Sitzung selbst verwalten {#_2-own-the-session}

Erzeugen Sie das Yjs-Dokument, die Awareness und den Provider selbst und übergeben Sie sie als `collaboration.externalSession` (den vollständigen Vertrag finden Sie unter [Vom Host verwaltete Zusammenarbeit](/de/guide/host-owned-collaboration)). Der Viewer verbindet, verbindet erneut oder zerstört sie nie.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const grant = await api.requestGrant(reviewId); // your API: ticket, role, intent, user
const doc = new Y.Doc();
const awareness = new Awareness(doc);
const provider = new WebsocketProvider(`wss://${location.host}/collab`, reviewId, doc, {
	awareness,
	params: { ticket: grant.ticket },
});

const viewer = createPptxViewer(container, {
	source: sourceBytes, // the same source PPTX for every participant
	editable: grant.role !== 'viewer',
	collaboration: {
		roomId: reviewId,
		serverUrl: '', // unused with an external session
		userName: grant.user.name, // from your directory, not typed by the user
		userColor: grant.user.color,
		role: grant.role,
		sessionIntent: grant.intent,
		externalSession: { doc, awareness, getSnapshot, subscribe },
	},
});
```

`getSnapshot()` gibt `{ status, synced }` zurück. Melden Sie `synced: true` nur, wenn der Provider synchronisiert ist **und** der Benutzer noch den Zugriff besitzt, mit dem der Editor eingehängt wurde. Setzen Sie den Wert auf `false`, wird der Editor sofort schreibgeschützt.

**Wer den Raum befüllt.** Einen leeren Raum füllt der erste Teilnehmer mit `sessionIntent: 'create'`; alle anderen übergeben `'join'` und übernehmen, was dort bereits liegt. Lassen Sie Ihren Server entscheiden, nicht die Clients: Das Beispiel gewährt `'create'` jeweils nur einem Schreibenden mit einer kurzen Lease, und nur solange der Raum noch leer ist. So veröffentlichen zwei Personen, die das Element im selben Moment öffnen, nie beide ihre Startpräsentation.

## 3. Sockets auf dem Server zulassen und überwachen {#_3-admit-and-police-sockets-on-the-server}

Eine WebSocket-Verbindung kann das Sitzungscookie Ihrer Website nicht zu einem Relay auf einer anderen Origin übertragen, und ein langlebiges Token in einer URL landet in Protokollen. Stellen Sie stattdessen ein **kurzlebiges, signiertes Ticket** aus einem gewöhnlichen, authentifizierten API-Aufruf aus, und lassen Sie das Relay nur dem vertrauen, was das Ticket aussagt:

- **Beim Upgrade**: Prüfen Sie Signatur, Ablauf und Raum des Tickets. Lehnen Sie alles andere ab.
- **Bei jeder Nachricht**: Eine Verbindung nur zum Lesen darf das Dokument anfordern (Sync-Schritt 1) und Präsenz teilen, aber ihre Sync-Schritt-2- und Update-Nachrichten werden verworfen. Das ist die serverseitige Hälfte von `role: 'viewer'`.
- **Awareness**: Lassen Sie jede Verbindung nur für die Awareness-Client-IDs sprechen, die sie zuerst beansprucht hat, damit niemand den Cursor eines anderen löschen oder sich als dieser ausgeben kann.
- **Ticket-Erneuerung**: Ein Ticket lässt nur eine Verbindung zu. Erneuern Sie es vor dem Ablauf und aktualisieren Sie `provider.params.ticket`, damit ein automatisches Wiederverbinden ein gültiges Ticket verwendet.

Sie müssen dieses Relay nicht selbst schreiben. `openteams-server` (der Referenzserver des OpenTeams-Viewers, ein mit y-websocket kompatibler Yjs-Server) nimmt einen Hook `authorizeSync` entgegen, der pro Socket `'write'`, `'read'` oder eine Ablehnung zurückgibt, bindet Awareness-IDs bereits fest und stellt `closeSync`, `syncPrincipals`, `isSyncRoomEmpty` und `handleUpgrade` für die folgenden Schritte bereit. Die Datei `server/relay.ts` des Beispiels umfasst rund sechzig Zeilen um ihn herum.

## 4. Zugriffe ändern, während Personen arbeiten {#_4-change-access-while-people-are-working}

Wenn ein Besitzer den Zugriff einer Person ändert, schließen Sie deren offene Sockets mit einem Anwendungs-Schließcode (das Beispiel verwendet `4403`). Der Client reagiert dann so:

1. Er meldet `synced: false`, wodurch der Editor eingefroren wird.
2. Er zerstört den Viewer, danach sein eigenes Dokument und seinen Provider.
3. Er fordert eine neue Berechtigung an und hängt eine frische Sitzung mit der neuen Rolle ein, oder zeigt "Zugriff entfernt", wenn es keine gibt.

Beginnen Sie nach einer Herabstufung mit einem **neuen** Y.Doc. Änderungen, die der Benutzer kurz vor der Umstellung gemacht hat, hat das Relay verworfen, und ein weiterverwendetes altes Dokument würde sie lokal weiter anzeigen.

## 5. Versionen mit einem einzigen Schreibenden speichern {#_5-save-versions-with-one-writer}

Hängen Sie genau einen Teilnehmer ein, normalerweise den Besitzer des Elements, mit `role: 'owner'` und `onWriteBack(bytes)`. Der Viewer verzögert Änderungen (`writeBackDebounceMs`, standardmäßig 5 Sekunden) und übergibt Ihnen einen PPTX-Snapshot der gemeinsamen Präsentation, den Sie als Version speichern. Andere Teilnehmer schreiben nie zurück, daher gibt es kein Wettrennen, bei dem die letzte Speicherung gewinnt.

`onWriteBack` ist keine Bestätigung einer dauerhaften Speicherung. Wiederholungen, Fehler und die Wahl eines Schreibenden, wenn der Besitzer offline ist, bleiben Ihre Entscheidungen. Wenn der Raum einen Neustart des Relays überleben muss, speichern Sie die Yjs-Updates zusätzlich auf dem Server (etwa mit einem eigenen Speicher oder einem Provider-Backend wie y-leveldb).

## 6. KI-gestützte Änderungen hinzufügen {#_6-add-ai-assisted-changes}

Übergeben Sie die Option `ai` nur an Konten, die sie nutzen dürfen, und richten Sie sie auf **Ihre** Route, damit der Modellschlüssel den Browser nie erreicht:

```ts
ai: {
	connection: {
		kind: 'endpoint',
		api: `/api/ai/chat?review=${reviewId}`,
		credentials: 'same-origin',
	},
	writePolicy: 'stage', // every change is a proposal the reviewer accepts or rejects
},
```

Die Route prüft Konto und Berechtigung des Aufrufers und führt dann das Modell mit den nur auf dem Schema basierenden Präsentationswerkzeugen aus `ooxml-ui/pptx/ai` aus (siehe [KI-Assistent](/de/guide/ai-assistant)). Werkzeugaufrufe laufen im Browser des Prüfers gegen die Live-Präsentation. Ein angenommener Vorschlag ist eine gewöhnliche Bearbeitung durch diesen Prüfer: Sie läuft über den Editor-Verlauf (ein Strg+Z macht sie rückgängig) und erreicht den Raum so, wie Tippen es tut. Ihr Audit-Log kann festhalten, dass der Assistent genutzt wurde; das gemeinsame Dokument kennzeichnet nicht, welche Änderungen von ihm stammen.

## Das Beispiel ausführen {#run-the-example}

```sh
bun install                                   # at the repository root
cd demos/pptx/demo-host-review
bun run start                                 # http://localhost:4190
```

Öffnen Sie die Seite in zwei Browsern (oder einem normalen und einem privaten Fenster) und melden Sie sich als verschiedene Personen an: Ana besitzt das Element, Ben und Chen prüfen, Dee beobachtet. Ändern Sie als Ana im Seitenpanel den Zugriff von Dee und sehen Sie zu, wie Dees Editor zwischen schreibgeschützt, bearbeitbar und ohne Zugriff wechselt, ohne neu geladen zu werden. Um den Assistenten zu testen, setzen Sie vor dem Start `REVIEW_AI_MODEL` (und `REVIEW_AI_API_KEY` oder `REVIEW_AI_BASE_URL` für einen lokalen OpenAI-kompatiblen Server).

## Grenzen {#limits}

- Das Beispiel hält alles im Speicher: Konten, Versionen und die gemeinsamen Räume sind weg, sobald der Prozess endet. Seine Anmeldung ist eine Kontoauswahl.
- Awareness-Inhalte (Anzeigename und Farbe) stammen vom Client. Das Relay bindet Awareness-IDs an Verbindungen, prüft den Inhalt aber nicht gegen das Konto im Ticket.
- Kommentare, ein Freigabestatus der Prüfung und die Zuordnung einzelner Änderungen sind in diesem Beispiel nicht Teil des gemeinsamen Modells. Prüfen Sie diese Abläufe gesondert, bevor Sie sich darauf verlassen.
- Dass Browser-Teilnehmer zusammenlaufen, ist nicht dasselbe wie PowerPoint, das die gespeicherte Version ohne Reparatur wieder öffnet. Prüfen Sie repräsentative Präsentationen in PowerPoint, bevor Sie das versprechen.
