# Protokoll zur Zusammenarbeit

Die Web Component unterstützt zwei sich gegenseitig ausschließende Modi der Zusammenarbeit: durch eine Autorität geordnete ProseMirror-Schritte und das optionale Zusammenführen mit Yjs über eine von der Anwendung verwaltete `CollabSession`. Beide verwenden denselben Editor und dieselben sechs Framework-Adapter. Der Host ist für Transport, Authentifizierung, Persistenz und den Zeitpunkt des Exports zuständig.

## Live-Demo: zwei Fenster, beliebiges Framework

Eine Sitzung ist nur ein Name. Die beiden Bereiche unten sind getrennte Demo-Apps, die aus verschiedenen Framework-Adaptern (React und Vue) gebaut wurden und der Sitzung `docs-collab` beitreten: Der erste Bereich öffnet das Beispieldokument und hostet die Sitzung, der zweite tritt über den Namen bei und erhält das Dokument, jede Änderung und den Cursor der anderen Person. Tippen Sie in einem der beiden Bereiche.

<iframe
	src="/ooxml/docx/demo/?sample=1&room=docs-collab&name=Ada"
	title="docx-viewer-Sitzungshost (React)"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/docx/demo-vue/?room=docs-collab&name=Grace"
	title="docx-viewer-Sitzungsgast (Vue)"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Beliebige zwei Frameworks kombinieren

Jede Demo akzeptiert `?room=<session>`. Das Fenster, das zusätzlich `sample=1` angibt, hostet die Sitzung; jedes andere Fenster mit demselben `room` tritt ihr bei, in welchem Framework es auch gebaut wurde. Die [Live-Demo](/de/#live-demo) auf der Startseite bietet für jedes Fenster eine Framework-Auswahl, und in zwei Tabs können Sie zum Beispiel `/demo-angular/?sample=1&room=my-room` und `/demo-svelte/?room=my-room` öffnen.

### Wie die Demo das umsetzt und warum sie nur im selben Browser funktioniert

Das Protokoll des Editors ist transportneutral, daher liefert die Demo den Transport: einen `BroadcastChannel`, benannt nach der Sitzung. Das Hostfenster betreibt die Referenzautorität im Arbeitsspeicher (`createCollaborationAuthority`), ordnet die Schrittstapel aller Beteiligten und sendet die akzeptierten per Broadcast; ein Gast startet mit dem Ausgangsdokument des Hosts und spielt die akzeptierten Stapel nach, sodass das Framework des jeweiligen Fensters keine Rolle spielt. Ein `BroadcastChannel` verbindet Fenster, Tabs und Frames eines einzigen Browserprofils und sonst nichts; die Sitzung endet, wenn ihr Hostfenster geschlossen wird, und es wird nichts gespeichert. Für Personen auf verschiedenen Rechnern muss Ihre Anwendung Transport, Autorität, Identität und Speicherung bereitstellen (siehe unten).

Es gibt außerdem eine [Einzelseiten-Demo](/demo/collaboration.html){target="_self"} mit zwei Editoren auf einer Seite und einer Schaltfläche **Pause delivery**, um gleichzeitige Änderungen auszuprobieren; mit `?guest=vue` läuft ihr zweiter Editor in einem anderen Framework.

## Was unterstützt wird und was nicht

| Unterstützt                                                                                              | Nicht unterstützt                                                                                             |
| -------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| Geordnete, validierte ProseMirror-Schrittstapel über eine Autorität                                      | Ein gehosteter Server, Netzwerk, Authentifizierung oder Persistenz (dafür ist der Host zuständig)             |
| Yjs-Zusammenführung von Text und Formatierung, relative Cursor, Kommentarfäden, lokales Rückgängigmachen | Gleichzeitige Bearbeitung von Formatvorlagen, Nummerierungsdefinitionen und Notizen außerhalb des Textkörpers |
| Eine Referenzautorität im Arbeitsspeicher für Tests und die lokale Demo                                  | Strukturelle Tabellenbefehle während der Zusammenarbeit                                                       |

## Yjs-Modus

Laden Sie in jedem Editor dasselbe DOCX-Quellpaket, einschließlich seiner Medien, Formatvorlagen, Kommentare, Notizen und opaken Teile. Der Host stellt eine stabile `documentId` für dieses Paket bereit und bestimmt genau einen Ersteller des Raums. Diese Kennung ist ein Vertrag des Hosts, kein Authentifizierungstoken und keine automatische Paketprüfsumme.

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({
	roomId,
	provider: transportProvider({ transport }), // Application-owned transport
	user: { name: 'Ada', role: 'collaborator' },
});
if (!session.synced) {
	await new Promise<void>((resolve) => {
		const off = session.on('synced', (synced) => {
			if (synced) {
				off();
				resolve();
			}
		});
	});
}
editor.startYjsCollaboration(session, {
	documentId: sourcePackageId,
	initializeIfEmpty: designatedCreator,
});
editor.publishPresence({ name: 'Ada', color: '#2563eb' });
```

Warten Sie vor dem Beitritt auf die tatsächliche anfängliche Provider-Synchronisierung und nicht auf das Schreib-Gate der Karenzzeit. Der Ersteller befüllt einen leeren Raum; andere Teilnehmer übernehmen den vorhandenen Textkörper sowie die Seiten- und Abschnittsattribute, statt sie mit einem lokalen Snapshot zu überschreiben. Eine abweichende Dokument-ID oder ein abweichendes Raumformat lehnt den Beitritt ab. Eine Autoritätssitzung und eine Yjs-Sitzung können einen Editor nicht gleichzeitig steuern.

Das aktuelle gemeinsame Word-Schema initialisiert Räume vom Typ `word-yjs-v3`. Jede modellierte direkte Eigenschaft eines Bildes, eines Notizverweises, einer Feldmarkierung oder eines Umbruchs hat ein eigenes Attribut über ihrer importierten Eigenschaftsbasis. Gleichzeitige Änderungen verschiedener Eigenschaften werden zusammengeführt, einschließlich des Entfernens von Eigenschaften; das lokale Rückgängigmachen behält die Änderungen des anderen Autors bei. Konflikte bei derselben Eigenschaft folgen der Yjs-Reihenfolge. Der Formatierungsverlauf hat weiterhin einen vorherigen Snapshot und einen Autor pro Run; die vollständige gleichzeitige Zuordnung von Überarbeitungen bei mehreren Autoren ist noch nicht abgeschlossen.

Das ändert die Raumkompatibilität: Der aktualisierte Viewer lehnt Räume der Versionen v1 und v2 ab, und alte Clients lehnen v3-Räume ab. Core-Konsumenten mit älteren Schemas können weiterhin passende v1/v2-Räume verwenden, nur mit Textkommentaren. Um einen persistierten Raum zu aktualisieren, exportieren Sie sein aktuelles DOCX mit dem passenden alten Client, laden Sie dieses Paket in aktualisierten Clients und lassen Sie den bestimmten Ersteller einen neuen leeren v3-Raum initialisieren. Halten Sie alle Teilnehmer auf dem passenden Schema. Das Binding schreibt einen aktiven Raum nicht um und migriert die Provider-Persistenz nicht automatisch.

Der Yjs-Modus verwendet das stabile Yjs-13-Binding `y-prosemirror` 1.3.7. Attribute auf Dokumentebene haben eine eigene Zuordnung, weil das Binding sie nicht synchronisiert. Text, Formatierung und nachverfolgte Überarbeitungsmarkierungen werden über das Textkörper-Fragment übertragen. Neu eingefügte Bildteile in PNG, JPEG, GIF, BMP und SVG nutzen das gemeinsame Asset-Routing des Cores mit rohen binären Nutzdaten; Namen sind pro Client eindeutig und Teile sind unveränderlich. Bildbytes bleiben nach dem Beenden der Zusammenarbeit für Rückgängigmachen/Wiederholen und Exporte verfügbar. Vorhandene Paket-Assets verbleiben in der jeweiligen passend geladenen Quelle jedes Editors. Der Host steuert weiterhin die Exportsynchronisierung und die Raumpersistenz; dies koordiniert keine gleichzeitigen Speichervorgänge und belegt keinen verlustfreien Export.

Auch die Aufzeichnungseinstellung für die Änderungsverfolgung wird über Dokumenttransaktionen übertragen. Das Umschalten aktualisiert die anderen Teilnehmer, nimmt am lokalen Rückgängigmachen/Wiederholen teil und wird als `w:trackRevisions` exportiert. Beim Beitritt zu einem kompatiblen Raum ohne diese Einstellung bleibt der Wert des geladenen Pakets erhalten, bis ein Teilnehmer ihn ändert. Überarbeitungsmarkierungen behalten den Namen des bearbeitenden Autors und den UTC-Zeitstempel der aufzeichnenden Transaktion. Die anderen Teilnehmer behalten diese Werte, statt fremde Änderungen neu zuzuordnen; vom Client gelieferte Namen und Uhren authentifizieren weder Urheberschaft noch Zeit.

Strg/Cmd+Z und das Rückgängigmachen im Menüband nutzen den lokalen Yjs-Verlauf. Auch normales Bearbeiten hält Fadendatensätze und Anker in rückgängig machbaren Transaktionen; lokale Fadensnapshots sind von den Yjs-Dokumentattributen ausgenommen. Nach dem Beenden von Yjs werden seine aktuellen Fäden zum normalen Verlaufssnapshot. Der Verlauf übersteht das Lösen und erneute Einbinden des Editors; getrennte Ansichten übernehmen Änderungen beim erneuten Einbinden. Awareness verwendet relative Positionen für die Zuordnung von Cursor und Auswahl, mit der vorhandenen lokalisierten Cursor-Oberfläche. `publishPresence` aktualisiert das Awareness-Profil und liefert im Yjs-Modus `null`; `leavePresence` löscht den Cursor. Der Provider überträgt Awareness direkt, daher bleiben `presence-send` und `receivePresence` APIs des Autoritätsmodus.

`reconnectCollaboration()` startet den Yjs-Provider neu und behält dabei das gemeinsame Dokument. `resyncCollaboration()` fordert einen Zustandsaustausch an und liefert `false`, wenn der Provider dies nicht kann. Der Host kann außerdem `session.connect()`, `disconnect()`, `peers()`, Abonnements für Status/Synchronisierung/Fehler und `destroy()` verwenden. `stopCollaboration()` löst das Word-Binding und behält den aktuellen Inhalt sowie neue Bild-Assets, auch während der Editor nicht eingebunden ist; die vom Host verwaltete Sitzung wird dabei nicht zerstört. Zerstören Sie diese Sitzung, wenn der Host den Raum verlässt.

Schreibgeschützte Editoren und Betrachterrollen empfangen Remote-Aktualisierungen, blockieren aber lokale Dokumentschreibvorgänge. Client-Rollen sind nur beratend: Erzwingen Sie die Autorisierung zusätzlich im Server oder Provider. Das Bearbeiten außerhalb des Textkörpers und strukturelle Tabellenbefehle bleiben deaktiviert. Seiten- und Abschnittsattribute werden erhalten und synchronisiert, aber die Zusammenarbeit über mehrere Stories hinweg und ein vollständiger Paket-Bootstrap sind noch nicht abgeschlossen.

Neue Yjs-Räume teilen außerdem Kommentaranker, Antworten, Auflösung und Löschung über den vorhandenen Überprüfungsbereich. Der Ersteller befüllt die geladenen Kommentare. Jeder Kommentar hat unabhängige Textattribute, elementrelative Bereiche und einen Datensatz; gleichzeitige Antworten behalten ihre eigenen IDs. Die Auflösung folgt der Konfliktreihenfolge von Yjs-Maps. Das Löschen eines Stammkommentars blendet seine Antworten aus, auch solche, die gleichzeitig offline geschrieben wurden. Rückgängigmachen stellt den Stamm und seinen Anker gemeinsam wieder her und macht überlebende Antworten sichtbar; das Rückgängigmachen jedes Autors behält die Operationen der anderen bei. Der Export erfasst die neuesten gemeinsamen Fäden, auch wenn der Editor gelöst ist.

V3-Räume führen in ihrer Word-Identitätszuordnung die Kommentarfähigkeit `inline-relative-v1` mit. Bilder, Umbrüche, Notizverweise und Feldmarkierungen behalten unabhängige Anker bei gleichzeitigem Hinzufügen, Löschen, benachbarten Textänderungen und lokalem Rückgängigmachen. Importierte Anker befüllen diese Bereiche, wenn ein neuer Raum erstellt wird. Das Kommentieren eines Ergebnisses oder einer Markierung eines ausgewogenen komplexen Feldes deckt das gesamte Feld ab, entsprechend dem aufgezeichneten nativen Word-Verhalten. Passende ältere Räume behalten ihre reine Textfähigkeit `independent-v1`. Kompatible Räume ohne sie behalten ihre geladenen Kommentare und deaktivieren das Bearbeiten von Kommentaren. Erstellen Sie zum Aktualisieren einen neuen Raum aus einem kanonischen gespeicherten Snapshot; ändern Sie die Fähigkeit nicht in einem aktiven Raum. Der Autoritätsschritt-Modus synchronisiert keine Kommentarmetadaten, daher bleibt dort das Bearbeiten von Kommentaren deaktiviert. Diese Client-Verhaltensweisen bieten keine serverseitige Autorisierung, keine dauerhafte Persistenz und keine modernen Word-Kommentarerwähnungen, Benachrichtigungen oder Aufgabenzuweisungen.

Probieren Sie die [Yjs-Demo mit zwei Peers](/demo/collaboration.html?mode=yjs){target="_self"}. Ihre Aktionen Pause delivery, Reconnect providers und Resync providers üben die Wiederherstellung im Arbeitsspeicher. Sie bietet weder Produktionsnetzwerk noch dauerhafte Speicherung.

## Protokoll

Die Web Component stellt `startCollaboration`, `getPendingCollaboration`, `receiveCollaboration` und `stopCollaboration` bereit. Hosts können mit `createCollaborationAuthority(model, options)` die Referenzautorität im Arbeitsspeicher mit demselben Schema und derselben Modellkonvertierung wie der Editor erzeugen.

## Ablauf bei Client und Host

Alle Teilnehmer, die dasselbe Dokument bearbeiten, verwenden dieselbe `sessionId`; jede Editor-Inkarnation verwendet eine eindeutige `clientId`. Die Wiederverwendung einer Client-ID kann dazu führen, dass Peers die Schritte der anderen fälschlich für Bestätigungen halten. Alle Teilnehmer müssen dasselbe Schema und dasselbe Ausgangsdokument verwenden. Um einer bestehenden Autorität beizutreten, initialisieren Sie den Client mit deren aktueller Version und dem kanonischen Dokument:

```ts
const client = new CollaborationClient({ sessionId, clientId, version: currentVersion });
const state = EditorState.create({ doc: canonicalDoc, plugins: [history(), client.plugin] });
```

Für das Custom Element rufen Sie `editor.startCollaboration({ sessionId, clientId, version })` auf, nachdem Sie den passenden Modell-Snapshot geladen haben. Halten Sie jeden Empfänger eingebunden, solange er teilnimmt: `receiveCollaboration` wirkt über die aktive Editoransicht. Hören Sie auf `collaboration-send`; reichen Sie jeden `StepBatch` bei der Autorität ein und liefern Sie den akzeptierten Stapel mit `receiveCollaboration(batch)` an jeden Editor aus. Ablehnungen wie `stale`, `out-of-order` oder `wrong-session` sollten eine Wiederherstellung auf Host-Ebene auslösen und nicht stillschweigend verworfen werden.

Rufen Sie nach lokalen Transaktionen `client.createPendingBatch(state)` auf. Es liefert `null`, wenn keine unbestätigten Schritte vorliegen, und sonst einen unveränderlichen Umschlag. Wiederholen Sie genau diesen Umschlag, bis die Autorität ihn akzeptiert; starten Sie keine weitere Anfrage, solange eine unterwegs ist. Die Methode liefert denselben Umschlag, bis der zugehörige akzeptierte Stapel empfangen wurde.

Der Host reicht den Umschlag bei seiner Autorität ein und leitet den akzeptierten `batch` an alle Teilnehmer weiter, einschließlich des Absenders. Die `version` eines akzeptierten Stapels ist die Version der Autorität unmittelbar vor den Schritten des Stapels. Teilnehmer übergeben den Umschlag an `client.receive(state, batch)`. Bei `applied` senden Sie die zurückgegebene Transaktion über den üblichen Statusaktualisierungspfad des Editors ab. So laufen gewöhnliche Modellaktualisierungen, Statusereignisse und das Speicherverhalten auch bei Remote-Änderungen. Das Anwenden von Remote-Stapeln berücksichtigt die Schreibschutzeinstellung der Ansicht nicht; diese Einstellung blockiert nur lokale Änderungen.

Das Protokoll akzeptiert höchstens 500 Schritte und 1.000.000 serialisierte JSON-Zeichen pro Stapel. Die Referenz-`CollaborationAuthority` ist eine Ordnungs- und Rebase-Implementierung im Arbeitsspeicher. Sie akzeptiert Anfragen der aktuellen Version und bildet veraltete Anfragen über zwischenzeitlich akzeptierte Schritte ab. Ihr Standardverlauf behält 2.000 Schritte, und sie puffert 5.000 Anfragebestätigungen für Wiederholungen. Eine veraltete Anfrage, die älter als der behaltene Verlauf ist, wird abgelehnt; eine zukünftige Version wird als `out-of-order` abgelehnt. Der Client merkt sich die zuletzt empfangenen 1.000 Stapel-IDs, um Duplikate zu ignorieren. Hosts sollten nach einer Lücke, einer abgelaufenen Version oder einem abgelehnten überdimensionierten Stapel einen Abgleich oder ein Neuladen aus ihrem kanonischen Snapshot anbieten.

## Präsenz

`PresenceClient` bietet einen getrennten flüchtigen Kanal für Namen, Farben, Cursor und Auswahlen von Peers. Er verwendet den Schrittstrom des Dokuments weder noch verändert er ihn. Fügen Sie sein `.plugin` neben dem Plugin der Zusammenarbeit hinzu. `publish(state, { name, color })` liefert eine Nachricht für den Host-Transport oder `null`, solange lokale Dokumentschritte noch auf Bestätigung warten. `leave(state)` liefert eine Entfernungsnachricht. Nach dem Empfang einer Nachricht senden Sie die Transaktion aus `receive(state, message)` ab, wenn ihr Status `applied` ist; ignorieren Sie `duplicate` und beheben Sie `stale`, `out-of-order` oder `wrong-session` gemäß der Richtlinie des Hosts.

Präsenznachrichten tragen die gemeinsame Sitzungs-ID, die Client-ID des Absenders, die bestätigte Version der Zusammenarbeit und eine pro Absender monoton steigende Sequenznummer. Leiten Sie Nachrichten in Reihenfolge weiter und veröffentlichen Sie nach der Bestätigung einer Änderung erneut; Positionen sind relativ zur bestätigten Version der Nachricht definiert. Das Plugin bildet vorhandene Cursor und Auswahlen über nachfolgende Dokumenttransaktionen ab, auch über Änderungen, die im schreibgeschützten Zustand des Editors empfangen werden. Es akzeptiert Namen mit bis zu 80 druckbaren Zeichen, Farben aus `PRESENCE_PALETTE` und höchstens 100 sichtbare Peers. Hosts sollten in ihrer Transportschicht Raten begrenzen und getrennte Peers ablaufen lassen. Präsenz wird weder gespeichert noch authentifiziert noch in den DOCX-Export aufgenommen.

## Stabile Identitäten

Für das gemeinsame Custom Element wird die Präsenz durch `startCollaboration` aktiviert:

```ts
import { PRESENCE_PALETTE } from 'docx-vanilla-viewer';

editor.addEventListener('presence-send', (event) => {
	transport.sendPresence(event.detail); // Application-owned transport
});
editor.publishPresence({ name: 'Ada', color: PRESENCE_PALETTE[0] });
// Incoming messages: editor.receivePresence(message)
// Explicit departure: editor.leavePresence()
```

Das Element veröffentlicht sein gespeichertes Profil nach Auswahländerungen und Bestätigungen von Änderungen erneut. Für eingehende Auswahlpositionen liefert es `out-of-order`, solange lokale Änderungen ausstehen; der Host kann nach der Synchronisierung neue Präsenz anfordern. Abmeldenachrichten enthalten keine Positionen und können einen Peer bei einer anderen Dokumentversion entfernen, wobei die Prüfungen von Sitzung und Absender-Sequenz weiterhin gelten.

Absatz- und Tabellen-IDs sind Dokumentinhalt und werden in ProseMirror-Schritten übertragen. Verwenden Sie für lokale Transaktionen `repairCollaborativeDocumentIds(state, clientId, idGenerator)` statt der Hilfsfunktion zur ID-Reparatur ohne Zusammenarbeit. Erzeugen Sie `idGenerator` einmal mit `createCollaborationIdGenerator(clientId)` und verwenden Sie ihn für die Lebensdauer dieses Editors weiter; das verhindert die Wiederverwendung von IDs nach dem Löschen und versieht lokal eingefügte Tabellen-, Zellen- und Absatz-IDs mit einem Namensraum. Der Host muss jeder aktiven Editor-Inkarnation eine eigene Client-ID zuweisen; Sitzungs-ID und Client-ID sind Protokollkennungen, keine Zugangsdaten.

`stopCollaboration()` setzt voraus, dass ausstehende Änderungen bestätigt wurden. Verwirft der Host sie absichtlich, rufen Sie `stopCollaboration(true)` auf, um die lokale Transportwarteschlange ausdrücklich zu verwerfen. Eine Trennung oder das Entfernen des Editors sollte nicht als Bestätigung gelten; behalten Sie den kanonischen Snapshot samt Version und setzen Sie die Arbeit gemäß der Transportrichtlinie des Hosts fort oder stellen Sie sie wieder her.

## Zuständigkeiten und Grenzen der Autorität

Der Anwendungshost ist zuständig für Transport, Authentifizierung und Autorisierung, dauerhafte Speicherung, Verteilung von Snapshots, Client-Identität, Wiederverbindungsrichtlinie und den Zeitpunkt des Exports. Die Autorität im Arbeitsspeicher eignet sich für Tests und eine lokale Demo, nicht als Mehrprozess- oder Produktionsserver. Der DOCX-Export sollte ein kanonisches, ausreichend synchronisiertes Modell verwenden, wenn die Anwendung Semantik für gemeinsames Speichern braucht; diese Bibliothek koordiniert keine Dateisperren und keine gleichzeitigen Exporte. Remote-Peers erhalten nur vom Schema unterstützte ProseMirror-Schritte, und die Autorität validiert diese Schritte, bevor sie sie akzeptiert. Strukturelle Tabellenbefehle sind während der Zusammenarbeit deaktiviert, weil der aktuelle Tabelleneditor die Tabellenstruktur als Ganzes ersetzt; vor einer sicheren Aktivierung dieser Befehle sind granulare Transaktionen auf Zellebene nötig.

Unterstützte Verknüpfungen von Inline-Bildern, Umbrüchen und Notizen werden über Run-Attribute getragen, sodass Import oder Bearbeitung nicht von Element-Marks abhängen, die Yjs verwirft. Der Dialog „Link“ liest und bearbeitet diese Ziele über dieselben Core-Befehle wie gewöhnliche Textlinks. Das Ändern des Ziels und das Entfernen werden synchronisiert und unterstützen das lokale Rückgängigmachen durch den Autor; gemeinsame Bildmedien bleiben in DOCX-Exporten verfügbar. Gleichungslinks und die Zuordnung von Hyperlink-Überarbeitungen sind von dieser Unterstützung ausgenommen.
