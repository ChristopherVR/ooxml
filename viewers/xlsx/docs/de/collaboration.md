# Zusammenarbeit

Mehrere Personen können gleichzeitig dieselbe Arbeitsmappe bearbeiten. Der Editor bildet die Arbeitsmappe über `ooxml-core/xlsx/collab` auf ein Yjs-Dokument ab und tritt über die gemeinsame Sitzung `ooxml-core/collab` einem Raum bei (dieselbe Sitzung und dieselben Provider, die auch die Word- und Visio-Editoren verwenden). Jede Änderung wird sofort gesendet, gleichzeitige Änderungen an verschiedenen Zellen werden zusammengeführt, und jede Person sieht, wo sich die anderen befinden.

## Freigabe starten

Im Editor: **Datei > Freigeben** (oder die Schaltfläche „Freigeben“ am rechten Ende der Registerkartenzeile des Menübands), einen Sitzungsnamen eingeben und **Freigabe starten** wählen. Ein weiteres Fenster oder ein weiterer Tab desselben Browsers, in dem eine Arbeitsmappe geöffnet ist und das die Freigabe mit demselben Sitzungsnamen startet, tritt ihr bei. Ohne einen von der Seite festgelegten Server verlässt nichts den Browser: Die Sitzung läuft über einen `BroadcastChannel`.

Im Code setzen Sie die Eigenschaft `collaboration` (in jedem Binding als Prop):

```ts
editor.newWorkbook(); // a window joining an existing room still needs a workbook open
editor.collaboration = {
	roomId: 'budget-2026',
	serverUrl: 'wss://collab.example.com/rooms', // y-websocket compatible; omit for same-browser only
	user: { name: 'Ada', color: '#2563eb' },
};
editor.addEventListener('collaboration-change', (event) => console.log(event.detail));
```

| Option      | Hinweise                                                                                                                                              |
| ----------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `roomId`    | 1 bis 128 Buchstaben, Ziffern, Bindestriche oder Unterstriche.                                                                                        |
| `serverUrl` | Basis-URL eines y-websocket-kompatiblen Servers; die Raum-ID wird angehängt.                                                                          |
| `provider`  | Eine eigene Provider-Factory (WebRTC-Signalisierung oder ein vorhandener Yjs-Provider über `externalProvider`). Hat Vorrang vor `serverUrl`.          |
| `session`   | Eine vorhandene `CollabSession`, erstellt mit `sanitizePayload: sanitizeXlsxPresence`. Der Editor zerstört nie eine Sitzung, die ihm übergeben wurde. |
| `user`      | `name` (Standard: `authorName`), `color` (Standard: eine Farbe aus der Palette), `role` (`viewer` macht den Editor schreibgeschützt).                 |

Solange `collaboration` gesetzt ist, tritt eine später geöffnete Arbeitsmappe demselben Raum bei; enthält der Raum bereits eine Arbeitsmappe, ersetzt deren Inhalt die geöffnete. Wird die Eigenschaft auf `null` gesetzt, verlässt der Editor den Raum. Die imperativen Formen sind `startCollaboration(options)` (wird aufgelöst, sobald der Raum beigetreten ist), `stopCollaboration()`, `reconnectCollaboration()`, `collaborationState` und `share()` (öffnet Datei > Freigeben). Jedes Binding leitet `collaboration` weiter, meldet `collaboration-change` (`onCollaborationChange`, Vue `collaboration-change`, Angular `collaborationChange`, Svelte `oncollaborationchange`) und stellt `share()` an seinem Handle bereit; das Entfernen des Editors verlässt den Raum.

## Was geteilt wird

- Zellwerte, Formeln (mit ihren zwischengespeicherten Werten), Zahlenformate und Zellformatierung.
- Blätter: hinzufügen, umbenennen, verschieben, löschen, ausblenden und Registerfarbe.
- Zeilenhöhen, Spaltenbreiten, ausgeblendete Zeilen und Spalten, verbundene Zellen und definierte Namen.
- Die Auswahl jeder Person, in den Rastern der anderen als Umriss in ihrer Farbe mit ihrem Namen gezeichnet, sowie die Personen der Sitzung als Avatare in der Titelleiste.

Rückgängig (Strg+Z, das Menüband und die Symbolleiste für den Schnellzugriff) macht in einer freigegebenen Arbeitsmappe nur Ihre eigenen Änderungen rückgängig; die Arbeit anderer wird nie zurückgenommen. Bricht die Verbindung ab, weist der Editor darauf hin und behält Ihre Änderungen; sie werden gesendet, sobald die Verbindung wieder besteht, und ein zweiter Hinweis bestätigt die Wiederverbindung.

## Was nicht geteilt wird

Diagramme, Bilder und andere Zeichnungen, Kommentare, Hyperlinks, bedingte Formate, Datenüberprüfung, Tabellen und Filter, Pivot-Tabellen, fixierte Bereiche und andere Ansichtseinstellungen, Seiteneinrichtung, Schutz, Dokumenteigenschaften und Berechnungseinstellungen bleiben in jedem Fenster lokal und werden von dort gespeichert. Das Einfügen und Löschen von Zeilen und Spalten wird als die Zellen übertragen, die dabei verschoben werden; eine gleichzeitige Änderung an einer verschobenen Zelle landet daher an ihrer alten Adresse. Eine Formel, die eine andere Person mit einem Blattnamen eingegeben hat, der gleichzeitig umbenannt wurde, behält den alten Namen. Die Konfliktregeln sind im Repository `ooxml` aufgeführt (`docs/collab-area.md`).

## Was der Host verantwortet

Der Editor liefert keinen Server mit. Ihre Anwendung verantwortet den Transport (die Server-URL oder den Provider), Authentifizierung und Autorisierung (Rollen sind im Client nur beratend), die Persistenz der Räume sowie den Zeitpunkt und das Fenster, aus dem die Arbeitsmappe gespeichert wird: Der Editor koordiniert gleichzeitiges Speichern nicht. Dies ist keine Excel-Koautorenschaft und erhebt keinen Anspruch auf Gleichwertigkeit damit.
