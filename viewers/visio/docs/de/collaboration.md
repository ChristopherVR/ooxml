# Zusammenarbeit {#collaboration}

Datei > Freigeben startet eine Live-Sitzung zwischen Fenstern eines Browsers. Eine Host-Anwendung kann über die Eigenschaft `collaboration` auch geräteübergreifend teilen; sie nimmt dieselben Optionen wie jeder andere Office-Editor.

::: warning Teilweise umgesetzt
In diesem Build gibt es keinen gehosteten Dienst. Mit Datei > Freigeben verlässt nichts Ihren Browser. Für den Austausch zwischen Geräten bringen Sie Ihren eigenen Server (eine y-websocket-kompatible URL), Provider oder Yjs-Sitzung mit; Seiten und Formen werden elementweise zusammengeführt, alles andere wird als Ganzes geteilt (siehe unten).
:::

## Einbettung: die Eigenschaft `collaboration` {#embedding-the-collaboration-property}

`<visio-viewer>` implementiert den gemeinsamen Einbettungsvertrag von `ooxml-core/collab` (`CollaborationOptions`, `CollaborationState`, `CollaborativeEditor`), denselben wie Word, Excel und PowerPoint; ein Host, der mehrere Produkte einbettet, integriert also nur einmal.

```ts
viewer.collaboration = {
	roomId: 'floor-plan', // 1 to 128 letters, digits, hyphens or underscores
	serverUrl: 'wss://example.com/collab', // optional: a y-websocket compatible server
	user: { name: 'Ada', color: '#2563eb' }, // optional; `role: 'viewer'` makes this window read-only
	intent: 'join', // optional: never seed an empty room with the drawing already open
};
viewer.addEventListener('collaboration-change', (event) => console.log(event.detail));
viewer.collaborationState; // { active, roomId, status, synced, people }
viewer.reconnectCollaboration();
viewer.stopCollaboration(); // or: viewer.collaboration = null
```

- **Transport**: `session` (eine vom Host verwaltete `CollabSession`, die der Viewer nie zerstört), sonst `provider`, sonst `webrtc`, sonst `serverUrl`, sonst ein `BroadcastChannel` zwischen Fenstern dieses Browsers.
- **Deklarativ**: Solange `collaboration` gesetzt ist, wird die geöffnete Zeichnung geteilt, ebenso eine später geladene. Mit `null` verlässt der Viewer den Raum. Eine vom Viewer selbst geöffnete Sitzung wird beim Verlassen zerstört, eine `session` des Hosts nicht.
- **Absicht (`intent`)**: Mit `join` wird ein leerer Raum nicht mit der lokalen Zeichnung befüllt, sodass ein später Beitretender keinen noch nicht geladenen Raum überschreiben kann. Eine danach geladene oder bearbeitete Zeichnung wird weiterhin veröffentlicht. Standard ist `create`.
- **Rollen**: `role: 'viewer'` macht den Viewer im Raum schreibgeschützt (`canEdit` wird beim Verlassen wiederhergestellt). Die Rolle ist nur ein Hinweis; ein Server muss sie durchsetzen.
- **Ereignisse**: `collaboration-change` liefert einen `CollaborationState`, wenn das Teilen beginnt oder endet und wenn sich Verbindungsstatus, erste Synchronisierung oder Teilnehmer ändern. Ein fehlgeschlagener Beitritt wird als `document-error` gemeldet. Datei > Freigeben nutzt denselben Controller und löst daher ebenfalls das Ereignis aus.
- **Bindings**: Jedes Binding reicht eine Eigenschaft `collaboration` durch (React, Vue, Solid, Svelte, Vanilla `mountViewer`/`update`, Angular `[collaboration]`) und leitet `collaboration-change` weiter (`events['collaboration-change']`, Vue `@collaboration-change`, Angular `(collaborationChange)`). Das Unmounten verlässt den Raum.
- **Typen**: `VisioCollaborationOptions` (die gemeinsamen Optionen plus `onWriteBack`, `writeBackDebounceMs` und `writeBackRetry`) und `VisioCollaborationState` werden aus `ooxml-ui/visio` exportiert.

## Server und Peer-to-Peer: Token und WebRTC {#server-and-peer-to-peer-auth-token-and-webrtc}

```ts
// A y-websocket compatible server. The token is sent as the `token` query parameter:
// wss://example.com/collab/floor-plan?token=<authToken>
viewer.collaboration = {
	roomId: 'floor-plan',
	serverUrl: 'wss://example.com/collab',
	authToken: await getToken(),
};

// Peer to peer over y-webrtc. Pass the `WebrtcProvider` class from your own y-webrtc install;
// the viewer does not bundle it. The room password defaults to `authToken`.
import { WebrtcProvider } from 'y-webrtc';
viewer.collaboration = {
	roomId: 'floor-plan',
	authToken: 'room password',
	webrtc: { WebrtcProvider, signaling: ['wss://signal.example.com'] },
};
```

Der Viewer reicht diese Optionen unverändert an `openCollaborationSession` weiter. Das Token authentifiziert nur gegenüber Ihrem eigenen Server (der Server muss es prüfen), und ein WebRTC-Passwort hält lediglich Peers ohne Passwort aus dem Raum; beides macht Rollen nicht verbindlicher als einen Hinweis. Peer-to-Peer-Räume bestätigen nie eine erste Synchronisierung, daher entscheidet die Karenzzeit der Sitzung, wann dieses Fenster schreiben darf.

## Live-Cursor und Auswahl {#live-cursors-and-selections}

Mitwirkende auf derselben Seite sehen gegenseitig ihren Zeiger (ein Cursor mit dem Namen in der Farbe der Person) und ihre ausgewählten Formen (eine gestrichelte Kontur in dieser Farbe) über der Zeichenfläche. Sie bewegen und skalieren sich mit Zoom und Verschiebung der Zeichnung und verschwinden, wenn die Person den Raum verlässt, die Seite wechselt, den Zeiger von der Seite nimmt, die Auswahl aufhebt oder veraltet ist.

Jedes Fenster veröffentlicht `{ page, x?, y?, selection? }`: die Seiten-ID, den Zeiger in Zeichnungseinheiten der Seite (Zoll) und bis zu 64 ausgewählte Form-IDs. `sanitizeVisioPresence` (aus `ooxml-core/visio/ui`) prüft jede eingehende Nutzlast. Der Viewer wendet es auf Sitzungen an, die er selbst öffnet. **Eine `session` des Hosts sollte damit erstellt werden** (zusammen mit `VISIO_INITIAL_PRESENCE` als `initialPresence`), sonst tragen ihre Peers keine Nutzlast und es werden weder Cursor noch Auswahlen gezeichnet (nichts geht kaputt).

Grenzen: Cursor und Auswahl werden nur für die Seite angezeigt, die beide geöffnet haben; Auswahlkonturen folgen dem gespeicherten Rahmen der Form, nicht ihrer gerenderten Geometrie.

## Besitzer-Rückschreiben {#owner-write-back}

Ein Host, der die geteilte Zeichnung speichern möchte, setzt `onWriteBack` und optional `writeBackDebounceMs` (Standard 5000):

```ts
viewer.collaboration = {
	roomId: 'floor-plan',
	user: { name: 'Ada', role: 'owner' },
	onWriteBack: (bytes) => upload(bytes), // VSDX bytes
	writeBackDebounceMs: 5000,
};
```

Nur das Fenster mit der Rolle `owner` schreibt, sodass Mitwirkende nicht um das Speichern konkurrieren. Der Callback wird ausgelöst, nachdem die Zeichnung für die Entprellzeit ruhig war (nach lokalen und entfernten Änderungen), und einmal beim Verlassen des Raums, falls ein Speichern aussteht. Für `collaborator` und `viewer` wird er nie ausgelöst. Geben Sie aus `onWriteBack` ein Promise zurück, um das Speichern zu bestätigen: Es läuft jeweils ein Schreibvorgang, Änderungen währenddessen werden zu einem Folgeschreibvorgang, und eine Ablehnung wird wiederholt, wenn Sie `writeBackRetry` setzen (`{ attempts, baseDelayMs?, maxDelayMs? }`). `collaborationState.writeBack` (auch in `collaboration-change`) meldet `{ status: 'idle' | 'pending' | 'saving' | 'saved' | 'failed', lastSavedAt?, error? }`; nur der Besitzer meldet es. Ein ausstehendes Speichern wird beim Verlassen des Raums, bei `pagehide` und beim Ausblenden der Seite abgeschlossen; ein asynchroner Schreibvorgang kann beim Schließen der Seite dennoch verloren gehen. Für Schreibvorgänge, die ein Schließen überstehen müssen, verwenden Sie `navigator.sendBeacon` oder `fetch(..., { keepalive: true })` (ca. 64 KiB Grenze).

## Was es heute gibt {#what-exists-today}

- Eine Live-Sitzung auf Basis von `ooxml-core/collab` (Yjs) über einen `BroadcastChannel`, sodass Fenster und Tabs desselben Browsers eine Zeichnung gemeinsam nutzen können; Hosts teilen über `collaboration` auch über Geräte hinweg.
- Seiten, Formen (bis hinunter zu Zellen, Zeilen und Abschnitten), Verbinder, Schablonen (Masters) und Bilder werden elementweise zusammengeführt (`ooxml-core/visio/collab`); alle übrigen Teile des Pakets (Stile, Design, Daten, Dokumenteinstellungen) werden als Ganzes geteilt, pro Teil gewinnt der letzte Schreibvorgang. Die Collab-Laufzeit wird erst geladen, wenn das Teilen beginnt.
- Gleichzeitige Änderungen an verschiedenen Formen bleiben beide erhalten; zwei Änderungen an derselben Form ergeben in jedem Fenster denselben Wert; eine gelöschte Form oder Seite gewinnt gegen eine gleichzeitige Bearbeitung.
- Rückgängig und Wiederholen (Strg+Z, das Menüband) wirken nur auf die Änderungen dieses Fensters und machen nie die eines Mitwirkenden rückgängig. Entfernte Änderungen werden außerhalb Ihres Verlaufs angewendet, und der Verlauf beginnt beim Beitritt.
- Mitwirkende erscheinen über das gemeinsame `ooxml-ui`-Präsenzelement.
- Der Sitzungsname besteht aus 1 bis 128 Buchstaben, Ziffern, Bindestrichen oder Unterstrichen. Ein Fenster, das einer Sitzung beitritt, deren Raum bereits eine Zeichnung enthält, übernimmt diese; andernfalls befüllt es den Raum (außer der Host verlangt `intent: 'join'`).
- Sitzungen aus dem Menü behalten ihren ursprünglichen BroadcastChannel-Namen (`ooxml-collab:visio-<room>`), sodass Fenster älterer Builds sich weiterhin treffen.

## Live-Demo: zwei Fenster, beliebiges Framework {#live-demo-two-windows-any-framework}

Eine Freigabesitzung ist nur ein Name. Beide Bereiche unten sind eigenständige Playgrounds, die mit unterschiedlichen Framework-Bindings gebaut wurden und beim Laden der Sitzung `docs-collab` beitreten: Fenster A ist das Vanilla-Playground und öffnet das Beispiel, Fenster B ist das React-Playground und erhält die Zeichnung aus dem Raum. Bearbeiten Sie den Formtext in einem Bereich und beobachten Sie, wie der andere folgt; Rückgängig macht den entfernten Schritt rückgängig.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer Fenster A (Vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer Fenster B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Zwei beliebige Frameworks koppeln {#pair-any-two-frameworks}

Jedes Playground akzeptiert `?share=<session>`, das Datei > Freigeben ausfüllt und startet. Nur der Sitzungsname zählt, daher funktioniert jede Kombination: Die [Live-Demo](/de/#live-demo) auf der Startseite bietet für jedes Fenster eine Framework-Auswahl, und in zwei Tabs können Sie zum Beispiel `/demo-vue/?sample=1&share=my-room` und `/demo-angular/?share=my-room` öffnen. Sie können es auch von Hand in zwei beliebigen Fenstern tun: Datei, dann Freigeben, dann derselbe Sitzungsname.

### Warum Datei > Freigeben nur im selben Browser funktioniert {#why-it-is-same-browser-only}

Das Menü nutzt einen `BroadcastChannel`. Er verbindet Fenster, Tabs und Frames eines Browserprofils auf einer Origin, unabhängig vom verwendeten Framework, und sonst nichts. Zwei verschiedene Rechner können sich so nicht treffen; dafür setzt ein Host `collaboration` mit `serverUrl`, `provider` oder `session`, und dieser Build liefert keinen Server mit.

## Was unterstützt wird und was nicht {#what-is-and-is-not-supported}

| Unterstützt                                                                             | Nicht unterstützt                                                                              |
| --------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| Sitzungen im selben Browser über BroadcastChannel, auf Basis von Yjs                    | Ein gehosteter Server: bringen Sie Ihren eigenen y-websocket-Server, Provider oder Sitzung mit |
| Zusammenführen pro Form, pro Zelle und pro Seite, Löschen gewinnt, Rückgängig nur lokal | Zusammenführen innerhalb einer Zelle: zwei Änderungen derselben Zelle behalten einen Wert      |
| Mitwirkende werden mit dem gemeinsamen Präsenzelement angezeigt                         | Kommentare; Rollen sind nur Hinweise (ein Server muss sie durchsetzen)                         |
| Datei > Freigeben und die Eigenschaft `collaboration` als Einstiegspunkte               | Cursor über verschiedene Seiten hinweg (nur die gemeinsame Seite wird gezeichnet)              |
| Live-Cursor und Auswahlen von Mitwirkenden auf derselben Seite                          | Präsenz einer Host-`session`, die ohne `sanitizeVisioPresence` erstellt wurde                  |
| Besitzer-Rückschreiben: entprellt, Promise-Bestätigung, Wiederholung, Status            | Eine Garantie, dass ein asynchroner Schreibvorgang beim Schließen der Seite abgeschlossen wird |
| `authToken` für einen Server und `webrtc` (Ihr eigener `WebrtcProvider`)                | Ein mitgeliefertes y-webrtc oder Signaling-Server; Token-Prüfung (Aufgabe Ihres Servers)       |

### Grenzen des Zusammenführens auf Formebene {#limits-of-shape-level-merging}

- Zwei Personen, die verschiedene Eigenschaften derselben Form ändern (Verschieben und Text- oder Füllungsänderung, verschiedene Zellen, Zeilen oder Abschnitte), behalten beide ihre Änderung. Ändern zwei Personen dieselbe Zelle oder denselben Text, bleibt in jedem Fenster ein Wert. Ganze `Text`- oder `Data`-Kinder sind je eine Einheit. Änderungen innerhalb eines Masters werden auf dieselbe Weise zusammengeführt (pro Form und Zelle), und ein Master, den eine Person löscht, gewinnt gegenüber einer gleichzeitigen Änderung darin.
- Seitenteile werden am Text geteilt, sodass Zellen genau so erhalten bleiben, wie der Editor sie geschrieben hat (Leerraum zwischen Elementen entfällt). Ein Seitenteil, der sich nicht teilen lässt (keine brauchbare `Shapes`-Liste), wird als Teil der Seitenhülle mit seinen Formen geteilt; der letzte Schreibvorgang gewinnt.
- Master und Medien werden je Element zusammengeführt: Master, die zwei Personen gleichzeitig hinzufügen, bleiben beide erhalten. Bilder werden über ihren Inhalt zugeordnet, sodass gleiche Bilder einmal gespeichert werden und verschiedene Bilder mit gleichem Dateinamen beide erhalten bleiben (eines wird umbenannt). Formate, Design, Dokumenteinstellungen, Daten und der Teil mit den Inhaltstypen werden weiterhin je Teil zusammengeführt, der letzte Schreibvorgang gewinnt. Fensterzustand und Vorschaubild bleiben lokal.
- Haben zwei Fenster gleichzeitig Formen mit derselben ID hinzugefügt, wird eine neu nummeriert; derselbe Raum-Inhalt ergibt auf jedem Fenster dieselben Nummern. Die `Sheet.n!...`-Formeln folgen der Form, für die sie geschrieben wurden, unabhängig davon, wann und wo sie eingegeben wurden (jede Formel merkt sich, welche Form sie benennt, sodass auch eine Formel, die genau im Moment der Kollision in einer älteren Form oder von einer anderen Person eingegeben wird, auf die gemeinte Form zeigt). Ein Fenster, das schreiben kann, speichert danach die Nummern im Raum, für die verschobene Form und für die, die ihre Nummer behalten hat, sodass sich keine von beiden ändert, wenn später weitere Formen hinzukommen, auch nicht eine im selben Moment mit derselben Nummer hinzugefügte Form. Rückgängig macht diesen Schritt nie, das Rückgängigmachen einer Formelbearbeitung nimmt ihn aber zurück. Nur ein Fall kann eine Nummer noch einmal verschieben: Zwei Fenster speichern im selben Moment unterschiedliche Nummerierungen für einen Raum, der außerdem eine Hinzufügung enthält, die das andere Fenster noch nicht gesehen hat; der Raum pendelt sich danach auf jedem Fenster ein. Seiten und Master werden auf dieselbe Weise neu nummeriert (die mit einem Master hinzugefügten Formen folgen ihm).
- Eine Verbindung (Glue), die eine Person entfernt und eine andere identisch neu erstellt, bleibt erhalten. Bilder entfernter Medien bleiben im Raum. Was eine Person löscht, während eine andere es bearbeitet (Zelle, Master, Bild), behält die Bearbeitung; eine gelöschte Form oder Seite gewinnt immer.
- Ein Raum, der von einer anderen Schemaversion des Collab-Formats geschrieben wurde, wird abgelehnt: Das Fenster meldet einen Fehler und liest oder schreibt nicht. Aktualisieren Sie alle Fenster eines Raums gemeinsam.
- Der Verlauf gilt pro Fenster, beginnt beim Beitritt und bleibt über einen Neuverbindungsvorgang hinweg nicht erhalten. Remote-Änderungen, die zusammen ankommen, werden mit einer Neuberechnung angewendet, und eine Änderung, die die Zeichnung unverändert lässt, wendet nichts an. Haben sich aber nur die Inhalte einiger Seiten geändert (und Sie haben in der Zwischenzeit nichts bearbeitet), werden nur diese Seiten erneut gelesen und die übrigen beibehalten (Formate, Designs und Master werden, solange sie unverändert sind, nicht erneut gelesen, und unveränderte Teile des Pakets werden nicht erneut geprüft). Eine Änderung der Seitenliste, der Master, Formate, Designs, Dokumenteinstellungen oder Bilder, die erste übernommene Zeichnung und die erste Remote-Änderung nach einer eigenen Bearbeitung lesen das gesamte Paket erneut; eine sehr große Zeichnung wird in diesen Fällen also weiterhin vollständig gelesen.

Die Nachweise sind im [Funktionsverzeichnis](/parity) unter „Collaboration and comments“ festgehalten.
