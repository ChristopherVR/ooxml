---
title: Vom Host verwaltete Zusammenarbeit
description: Stellen Sie ein vorhandenes Yjs-Dokument, Awareness und den Synchronisationszustand bereit, ohne dem Viewer die Verantwortung für den Transport zu übergeben.
---

# Vom Host verwaltete Zusammenarbeit

Verwenden Sie `collaboration.externalSession`, wenn Ihre Anwendung bereits Authentifizierung, einen Yjs-Provider, Offline-Speicher oder eine gemeinsame Verbindung verwaltet. React, Vue, Angular, Svelte und Vanilla akzeptieren denselben Vertrag. Der Viewer übernimmt sein Folienschema und die Anwesenheit. Der Host übernimmt das Verbinden, das erneute Verbinden, die Autorisierung und die Lebensdauer der Ressourcen.

```ts
import type { CollaborationConfig, ExternalCollaborationSession } from 'pptx-react-viewer';

// These resources and the store belong to your application, not to a viewer mount.
const session: ExternalCollaborationSession = {
	doc: hostDocument,
	awareness: hostAwareness,
	getSnapshot: () => connectionStore.getSnapshot(),
	subscribe: (listener) => connectionStore.subscribe(listener),
};

const collaboration: CollaborationConfig = {
	roomId: 'presentation-123',
	serverUrl: '', // Required by the shared config, but unused for an external session.
	userName: 'Alice',
	sessionIntent: 'join',
	externalSession: session,
};
```

Importieren Sie die Typen aus dem Paket für Ihr Framework. Übergeben Sie `collaboration` an die übliche Prop bzw. Option des Viewers. Das Dokument muss ein echtes `Y.Doc` sein, und die Awareness muss zu diesem Dokument gehören. Halten Sie das Sitzungsobjekt stabil und benachrichtigen Sie die Abonnenten, wenn sich `status` oder `synced` ändert. Ein Abonnement gibt eine Funktion zum Abmelden zurück.

Der Host und der Viewer müssen dieselbe Yjs-Laufzeit verwenden. Der Viewer lässt Yjs außerhalb seines Bundles und deklariert es als optionale Peer-Abhängigkeit. Installieren Sie `yjs` in der Anwendung, sowohl für integrierte als auch für vom Host verwaltete Zusammenarbeit:

```sh
npm install yjs
```

Ohne Zusammenarbeit, also beim reinen Anzeigen und Bearbeiten, wird diese Peer-Abhängigkeit nicht benötigt. Der Host installiert außerdem seinen gewählten Provider und seine Awareness-Implementierung. Deduplizieren Sie Yjs, wenn Ihr Abhängigkeitsbaum mehrere Versionen enthält: Typen aus getrennten Laufzeitkopien lassen sich nicht sicher mischen.

## Den gemeinsamen Kollaborationsvertrag verwenden {#using-the-shared-collaboration-contract}

Word, Excel und Visio übernehmen eine einzige `collaboration`-Option aus `ooxml-core/collab` (`CollaborationOptions`: `roomId`, `serverUrl`, `provider`, `session`, `user`, `intent`). PowerPoint akzeptiert dieselbe Form in seiner `collaboration`-Prop neben seiner eigenen `CollaborationConfig`. So kann ein Host, der mehrere Produkte einbettet, allen eine Sitzung übergeben:

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({ roomId, provider: transportProvider({ transport }), user });
// <PowerPointViewer collaboration={{ session, user: { name: 'Ada' }, intent: 'join' }} />
```

- `session` wird nur geliehen: Sie wird zur `externalSession`, und der Viewer zerstört sie nie. Der Host erstellt, verbindet und zerstört sie. `roomId` übernimmt standardmäßig den Raum der Sitzung.
- `user.name`, `color`, `avatar` und `role` werden auf `userName`, `userColor`, `userAvatar` und `role` abgebildet. `intent` wird auf `sessionIntent` abgebildet. Wenn beide gesetzt sind, gewinnen die expliziten Legacy-Felder.
- Jedes Binding führt an der Stelle, an der es die Prop liest, dieselbe `resolveCollaborationConfig` aus (aus `ooxml-ui/pptx`). Sie verhalten sich daher identisch. Eine Legacy-`CollaborationConfig` wird unverändert zurückgegeben.

Ohne `session` decken dieselben Optionen auch die Fälle ab, für die ein Host keine Sitzung erstellen möchte:

```ts
// Your own transport: the viewer opens and destroys the session around it.
collaboration = { roomId, provider: transportProvider({ transport }), user: { name: 'Ada' } };
// No server: tabs and windows of this browser share over a BroadcastChannel.
collaboration = { roomId, user: { name: 'Ada' } };
```

Der Transport wird in dieser gemeinsamen Reihenfolge gewählt: `session` (geliehen, wird nie zerstört) vor `provider` vor `serverUrl` vor BroadcastChannel. Für `provider` und den BroadcastChannel-Raum ruft der Viewer `openCollaborationSession` auf. Die Laufzeit `ooxml-core/collab` wird erst dann importiert, und sie teilt sich die eine Yjs-Kopie. Der Viewer übergibt die Sitzung über denselben Pfad für geliehene Sitzungen, den auch eine Host-`session` nimmt, und zerstört sie beim Beenden, beim Unmount oder wenn sich die Konfiguration ändert. Die Sitzung ist beschreibbar, sobald sie synchronisiert ist, oder nach der Schonfrist für einen einzelnen Teilnehmer in einem Mesh. `intent: 'join'` wartet trotzdem auf die Präsentation des Raums, und `role: 'viewer'` bleibt schreibgeschützt. Der Zustand meldet `connecting`, während die Sitzung geöffnet wird, und `error` bei einem ungültigen Raumnamen oder einem Browser ohne BroadcastChannel.

Eine Legacy-`CollaborationConfig` behält ihre Transporte: `serverUrl` wählt y-websocket, und `transport: 'webrtc'` oder eine leere `serverUrl` wählt das serverlose y-webrtc, das auch Tabs im selben Browser synchronisiert. Nur Eingaben des gemeinsamen Vertrags, die weder `serverUrl` noch `transport` nennen, verwenden den BroadcastChannel-Raum. Die Anwesenheit des Viewers und der Identitäts-Heartbeat der Sitzung teilen sich das Awareness-Feld `presence`. Der Viewer behält seine Folien- und Cursordaten über die Heartbeats hinweg bei, daher bleiben entfernte Cursor sichtbar.

## Den gemeinsamen Zustand melden {#reporting-the-shared-state}

Jedes Binding meldet denselben `CollaborationState`, den Word, Excel und Visio melden (`active`, `roomId`, `status`, `synced`, `people`, wobei die lokale Person mit `self` markiert ist). Der Typ stammt aus `ooxml-core/collab` und wird außerdem von jedem Paket re-exportiert. Die Meldung erfolgt, wenn eine Sitzung startet oder endet (der Aus-Zustand `status: 'off'`), und wenn sich Status, Synchronisiert-Flag oder Personen ändern. Das gilt gleichermaßen für vom Host und für vom Viewer verwaltete Sitzungen. Ein unveränderter Zustand wird nie wiederholt:

| Binding | Benachrichtigung                                             |
| ------- | ------------------------------------------------------------ |
| React   | Prop `onCollaborationChange`                                 |
| Vue     | Ereignis `collaboration-change`                              |
| Angular | Output `collaborationChange`                                 |
| Svelte  | Callback-Prop `oncollaborationchange`                        |
| Vanilla | Option `onCollaborationChange` und `getCollaborationState()` |

Vanilla behält `onCollaborationStatus` und `getCollaborationStatus()` für den rohen Verbindungsstatus.

## Die Bereitschaft entscheidet der Host {#readiness-is-the-host-s-decision}

`getSnapshot()` gibt `{ status, synced }` zurück. Der Status ist `connecting`, `connected`, `disconnected` oder `error`. `synced` bedeutet, dass das maßgebliche Dokument verfügbar ist und der Viewer Bearbeitungen veröffentlichen darf. Das Flag ist unabhängig vom Netzwerkstatus:

- Halten Sie `synced: false`, solange das gemeinsame Dokument geladen oder befüllt wird. Kein Timeout aktiviert Schreibzugriffe im Namen des Hosts.
- Setzen Sie `synced: true` nach der initialen Synchronisation. Vorhandene Folien des Raums werden übernommen, bevor ein lokales Start-Deck veröffentlicht werden kann.
- Lassen Sie es während einer Offline-Phase auf `true`, wenn Bearbeitungen im Yjs-Provider des Hosts in eine Warteschlange gestellt werden dürfen.
- Setzen Sie es auf `false`, um den Editor schreibgeschützt zu machen und das dauerhafte Zurückschreiben während einer neuen Synchronisation auszusetzen. Das Bearbeiten wird fortgesetzt, sobald der maßgebliche Zustand des Raums übernommen wurde.

Angenommene Texteingaben und Gestenvorschauen werden synchron in das geliehene Dokument geschrieben. So verwirft eine spätere Änderung der Bereitschaft einen noch ausstehenden letzten Tastendruck nicht. Der Provider des Hosts kann den Netzwerkverkehr weiterhin bündeln. Eine aktive Textbearbeitung wird geschlossen, wenn die Bereitschaft entzogen wird. Beim Fortsetzen wird der maßgebliche Raumzustand übernommen, statt einen veralteten lokalen Entwurf zu wiederholen.

Verwenden Sie `sessionIntent: 'create'` für den Teilnehmer, der einen leeren Raum befüllen darf, und `'join'` für Teilnehmer, deren Start-Deck den Raum nicht befüllen darf. Ein ausdrückliches **File > Open** (Datei > Öffnen) ist weiterhin eine bewusste Ersetzung. Stimmen Sie die Erstellung des Raums im Host ab. Dies ist keine serverseitige Autorisierung und kein Wahlprotokoll.

## Eigentum und Persistenz {#ownership-and-persistence}

Der Viewer erstellt nie einen zweiten Provider, trennt nicht den Provider des Hosts und zerstört das übergebene Dokument oder die Awareness nicht. Beim Unmount entfernt er seine Listener und stellt nur das Anwesenheitsfeld wieder her, das ihm noch gehört. Nicht zugehörige Awareness-Felder und spätere Schreibvorgänge des Hosts bleiben erhalten. Verwenden Sie pro Awareness-Instanz genau einen aktiven Schreiber für die Folienanwesenheit.

Der Host darf Yjs-Updates direkt speichern. Alternativ bestimmen Sie einen Teilnehmer mit `role: 'owner'` und `onWriteBack(bytes)`, der entprellte PPTX-Snapshots erhält. Diese verwenden dieselbe Quell-PPTX und dieselben Speicheroptionen wie der Editor. Ausstehende und laufende Snapshots werden abgebrochen, wenn die Sitzung endet oder der Host die Bereitschaft entzieht. Die Wahl des Schreibers bleibt Aufgabe des Hosts. Geben Sie aus `onWriteBack` ein Promise zurück, um die Speicherung zu bestätigen: Der Viewer meldet `saving`, `saved` oder `failed` in `collaborationState.writeBack` (nur für Owner). Er serialisiert Schreibvorgänge (einer ist in Bearbeitung, spätere Änderungen werden zu einem Folgeschreiben zusammengefasst) und wiederholt eine Ablehnung als `writeBackRetry` (`{ attempts, baseDelayMs?, maxDelayMs? }`), wie es beschrieben ist. Eine einfache Rückgabe zählt sofort als bestätigt. Ein ausstehendes Zurückschreiben wird bei `pagehide` und dann geleert, wenn die Seite ausgeblendet wird. Browser warten beim Schließen einer Seite nicht auf asynchrone Arbeit. Ein Schreibvorgang, der einen Netzwerk-Roundtrip benötigt, kann daher beim Schließen trotzdem verloren gehen. Ein Host, dessen `write` ein Schließen überstehen muss, sollte die Bytes mit `navigator.sendBeacon(url, bytes)` oder `fetch(url, { method: 'PUT', body: bytes, keepalive: true })` senden. Der Browser schließt diese Anfragen ab, nachdem die Seite verschwunden ist. Keepalive-Bodys sind auf etwa 64 KiB begrenzt, daher kann ein großes Dokument trotzdem nicht ankommen.
Eine aktive Sitzung mit `role: 'viewer'` macht den Editor schreibgeschützt und blockiert die Veröffentlichung, sowohl für integrierte als auch für vom Host verwaltete Zusammenarbeit. Der Server muss die Zugriffskontrolle unabhängig davon erzwingen.

## Grenzen der Integration {#integration-boundaries}

Diese Option ändert die Verantwortung für die Sitzung, nicht das bestehende Datenmodell der Zusammenarbeit. Der Viewer besitzt `pptx:slides` und seine verschachtelten Yjs-Typen. Geben Sie allen Teilnehmern dieselbe Quell-PPTX für Paketressourcen, Designs und Medien. Andere Anwendungsdaten können eigene Yjs-Schlüssel belegen. Gleichen Sie vollständige Folien-Snapshots nicht selbst im Host ab.

Prüfen Sie die Nachrichten- und Speicherlimits des Hosts anhand repräsentativer Präsentationen. Dazu gehören das initiale Dokument-Update und `Y.encodeStateAsUpdate(doc)` nach dem Bearbeiten. Die Dateigröße einer komprimierten PPTX ist nicht ihre Yjs-Snapshot-Größe: Das gemeinsame Schema trägt außerdem geparstes XML, Form- und Gruppenmetadaten sowie Asset-Nutzdaten. Allein das Auslagern von Bildern aus dem Dokument reicht möglicherweise nicht aus. Diese Sitzungs-API fügt kein Transport-Chunking hinzu und ändert keine Limits des Hosts.

Sie ergänzt keine Anbindung an ein Kommentar-Backend, keine neue lokale Rückgängig-Implementierung und keine Synchronisation für jedes Feld auf Präsentationsebene. Prüfen Sie diese Abläufe separat, bevor Sie sie in einem kollaborativen Produkt anbieten. Ein erneutes Öffnen in nativem PowerPoint ist ebenfalls eine eigene Treueprüfung, getrennt davon, ob zwei Browser-Teilnehmer konvergieren.

## Die Demos ausprobieren {#trying-the-demos}

Starten Sie das Demo-Relais (`bun run collab`) und eine Framework-Demo. Öffnen Sie dann:

```text
/?externalSession=1&room=example&sample=1&server=ws://localhost:1234
```

Ein zweiter Teilnehmer verwendet denselben Raum ohne `sample=1`. Die Host-Leiste stellt Steuerelemente für die Bereitschaft und das Einhängen des Editors bereit und behält Dokument und Provider bei. Die Demo deaktiviert BroadcastChannel, damit die Synchronisation über das WebSocket-Relais läuft.

Fügen Sie `&headless=1` hinzu, um eigene Oberflächen in einer der fünf Demos auszuprobieren. Diese Beispiele verwenden die öffentlichen Bausteine des Bindings, ohne einen versteckten vollständigen Viewer einzubinden. Die vom Host verwaltete Schaltfläche Speichern ruft den öffentlichen Serialisierungspfad auf. Mit der Zoomauswahl prüfen Sie entfernte Cursor und Auswahlumrisse bei verschiedenen Maßstäben.

Siehe [React-Hooks](/react/hooks), [Vue-Zusammenarbeit](/vue/collaboration), [Angular-Zusammenarbeit](/angular/collaboration), [Svelte-Zusammenarbeit](/svelte/collaboration) und [Vanilla-Eigene-Oberflächen](/vanilla/api#custom-shell-collaboration) für die jeweilige Komposition.

Für eine vollständige Anwendung um diese Option herum (Konten, ein berechtigungsgeprüftes Relais, Live-Änderungen der Zugriffsrechte, Owner-Versionen und eine geschützte KI-Route) siehe [Prüfung innerhalb Ihrer Anwendung](/de/guide/review-workflow).
