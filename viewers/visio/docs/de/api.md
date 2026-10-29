# Viewer-API {#viewer-api}

Die frameworkneutrale API hinter jedem Viewer-Paket. Jedes Framework-Paket
(`visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`,
`visio-svelte-viewer`, `visio-solid-viewer`, `visio-vanilla-viewer`) reicht
dieselben Handle-Methoden und Ereignisse weiter; die nativen Props, Ereignisse
und den Lebenszyklus jedes Adapters beschreibt `packages/bindings/README.md`.
Rendering und Bearbeitung sind Beta-Funktionen; das
[Funktionsverzeichnis](/parity) listet auf, was fehlt.

## Einhängen und Laden {#mount-and-load}

```ts
import { mountViewer } from 'visio-vanilla-viewer';
const viewer = mountViewer(container, {
	zoom: 1,
	events: { 'document-error': (error) => console.error(error) },
});
await viewer.load(file); // Blob, Uint8Array or ArrayBuffer
viewer.fit();
viewer.update({ pageIndex: 1 });
const snapshot = viewer.exportSvg(); // SVG string, dimensions, byteLength and diagnostics
const prepared = viewer.createPrintSnapshot(); // Frozen current-page artifacts, no printing
viewer.destroy(); // idempotent; later commands reject
```

Das Importieren des Moduls ist SSR-sicher; das Einhängen und die Registrierung
des Elements gibt es nur im Browser. `ViewerController` kommt ohne DOM aus,
sodass ein anderer Renderer denselben Zustand und dieselben Ereignisse
abonnieren kann. `renderPage` ist der Port des SVG-Renderers. Die
Modulzuständigkeiten stehen in der [Architektur](/de/architecture).

## Eigenschaften und Ereignisse {#properties-and-events}

Jedes Binding akzeptiert dieselben Eigenschaften und meldet dieselben Ereignisse (der Vertrag sind `ViewerProperties` und `ViewerEvents` in `ooxml-ui/visio`; die [Bindings](/bindings) listen die nativen Formen auf).

| Oberfläche    | Vertrag                                                                                                                                                                                |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Eigenschaften | `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `customization`                                                     |
| Ereignisse    | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change`, `file-command` |

- `canEdit` (boolean) ist die Bearbeitungsberechtigung. Der vollständige Viewer ist standardmäßig bearbeitbar, und `canEdit: false` macht ihn schreibgeschützt; die [Bausteine](/building-blocks) sind schreibgeschützt, bis der Host `true` übergibt. Eine Zeichnung ohne Paket-Bytes (ein als `document` übergebenes Modell, eine ältere `.vsd`) ist unabhängig von `canEdit` schreibgeschützt.
- `document-change` liefert `{ document, dirty, kind }`. `kind` ist `edit`, `undo` oder `redo` für Änderungen dieses Fensters, `remote`, wenn die Änderung eines Mitwirkenden angewendet wurde, und `save` nach `markClean()`.
- `markClean()` (auf `viewer.controller` und auf den Bausteinen) erklärt den aktuellen Inhalt für gespeichert. Ein Host, der `exportVsdx()` selbst speichert, ruft es danach auf: `edit.dirty` wird `false`, ein `document-change` der Art `save` wird ausgelöst, und spätere Bearbeitungen und Rückgängig-Schritte werden ab diesem Punkt gemessen. Bei einer Zeichnung, die nicht zurückgespeichert werden kann, bewirkt es nichts.
- `collaboration` (`VisioCollaborationOptions | null`) teilt die geöffnete Zeichnung in einem Raum, solange sie gesetzt ist. `collaboration-change` liefert einen `CollaborationState`, wenn das Teilen beginnt oder endet und sich Status, Synchronisierung oder Teilnehmer ändern. Am Element (`viewer.element` eines Handles) liest `collaborationState` den aktuellen Zustand, und `reconnectCollaboration()` und `stopCollaboration()` wirken darauf. Siehe [Zusammenarbeit](/de/collaboration).

## Formen auswählen {#select-shapes}

Jedes eingehängte bzw. native Handle stellt `selectShapes(targets)`,
`selectAll()` und `clearSelection()` bereit. Ein Ziel hat die Form
`{ id, name, pageId? }`; ein ausgelassenes `pageId` bedeutet die aktuelle Seite.
`controller.state.selectedShapes` ist ein geordnetes, unveränderliches Array,
dessen erstes Element die primäre `selectedShape` ist. Das bestehende Ereignis
`shape-select` meldet dieses primäre Element; `selection-change` meldet das
vollständige Array. Diese Zustände sind über den nativen Hook, das Composable,
das Signal oder den Store des jeweiligen Frameworks verfügbar. Vue sendet
`selection-change`; Angular stellt `selectionChange` bereit; alle sechs Adapter
akzeptieren `events['selection-change']`.

Ein Klick mit Umschalt, Strg oder Meta schaltet eine Form um. Klicks auf den
Hintergrund und Escape heben die Auswahl auf. Ausgeblendete oder gelöschte Ziele
werden entfernt, die Auswahl einer Gruppe schließt ihre Nachkommen aus, und ein
Wechsel von Seiten oder Dokumenten leert die Auswahl. „Alles auswählen“ umfasst
sichtbare Vordergrundformen. Die Panels für Text und Geometrie einer einzelnen
Form erfordern genau eine ausgewählte Form. Formatierung auf der Registerkarte
Start und Löschen wirken atomar auf mehrere ausgewählte Formen; der
Quellschutz eines beliebigen Ziels lehnt den gesamten Stapel ab.

## Eine leere Zeichnung erstellen {#create-a-blank-drawing}

Jedes eingehängte bzw. native Handle und das gemeinsame Element stellen
`createBlankDrawing(options?: CreateVsdxOptions): Promise<void>` bereit. Der
Controller bietet dieselbe Methode. Die optionalen Werte `width` und `height`
verwenden physische Zoll; die Vorgaben sind 8,5 mal 11 mit einer leeren Seite
und einem Zeichnungsmaßstab von 1:1. Datei > Neu > Leere Zeichnung und Strg+N
rufen denselben Lebenszyklus auf. Bearbeitbare Steuerelemente behalten ihr
natives Tastaturverhalten.

Eine akzeptierte Neu-Anfrage ersetzt Dokument, Auswahl und Verlauf, und das
Element nennt es `New drawing.vsdx`. Die Quellbytes können sofort exportiert
werden; `dirty` hält spätere Bearbeitungen fest. Überholte Ladevorgänge,
Abbrüche und das Entsorgen können kein verspätetes Ergebnis mehr annehmen.
Zustandsänderungen laufen über die bereits vorhandene reaktive Zustands-API
jedes Bindings.

## Zeigerbewegung und Auswahlrahmen {#pointer-movement-and-marquee-selection}

Das Ziehen einer geeigneten lokalen 2D-Form verschiebt die aktuelle Auswahl um
eine gemeinsame Translation. Eine Schwelle von vier Client-Pixeln erhält
gewöhnliche Klicks. Die Vorschau ändert nur die SVG-Darstellung; beim Loslassen
wird eine atomare Quellbearbeitung übermittelt. Escape, verlorene
Zeigererfassung, Änderungen von Quelle, Seite oder Auswahl sowie Werkzeugwechsel
brechen sie ab. Die Bewegung nutzt physische Seitenkoordinaten und rechnet im
Kern in Zeichnungseinheiten um, wobei gespeicherte Drehung und lokale Pins
erhalten bleiben. Quellschutz bleibt maßgeblich.

Das Ziehen über leeren Seitenraum wählt vollständig umschlossene sichtbare
Vordergrundformen anhand ihrer transformierten Breiten- und Höhenausdehnung aus.
Umschalt, Strg oder Meta fügen sie der aktuellen Auswahl hinzu. Die
Zeigerbewegung schließt derzeit Master, Gruppen, Ebenen, Fremdobjekte und
verklebte Verbinder aus. Größenänderungs-Griffe und Verbinder-Routing sind
getrennte Funktionen; die Größenänderungs-Griffe werden unten beschrieben, das
Routing liegt außerhalb dieses Gestenumfangs.

## Größenänderungs-Griffe und Größe & Position {#resize-handles-and-size-position}

Das gemeinsame Zeigerwerkzeug zeigt acht Größenänderungs-Griffe für eine
geeignete lokale 2D-Form. Kantengriffe ändern eine Dimension, Eckgriffe beide.
Die gegenüberliegende Kante bzw. Ecke bleibt in den gespeicherten gedrehten und
gespiegelten lokalen Achsen der Form fest. Gespeicherte außermittige lokale Pins
werden proportional skaliert, soweit der Quellnachweis das erlaubt. Während des
Ziehens zeigt nur der Auswahlrahmen die neuen Grenzen als Vorschau; beim
Loslassen werden eine Quellbearbeitung und ein Rückgängig-Schritt übermittelt.
Wird der gegenüberliegende Griff überschritten, wird die Dimension auf den
kleineren Wert aus ursprünglicher Größe und 1/16 physischem Zoll begrenzt,
ohne die Spiegelungs-Flags zu ändern. Einrasten, Gesten mit
Seitenverhältnis-Modifikator, Vorschau des vollständigen Inhalts und die
Größenänderung mehrerer Formen sind nicht implementiert.

Ansicht > Aufgabenbereiche > Größe & Position öffnet den gemeinsamen
numerischen Bereich. X/Y sind der Drehpunkt (Pin) in Zeichnungszoll vom Ursprung
unten links, wobei Y nach oben wächst. Breite und Höhe verwenden ebenfalls
Zeichnungszoll, unabhängig vom physischen Seitenmaßstab. Der Winkel wird in Grad
angegeben, positiv gegen den Uhrzeigersinn. Eingabetaste oder das Verlassen
eines Feldes übernimmt dieses Feld genau einmal; Escape stellt den gespeicherten
Wert wieder her. Änderungen von Breite und Höhe in diesem Bereich halten den
Drehpunkt fest, anders als die verankerten Zeigergriffe. Die Eingabe des exakten
aktuellen Werts lässt Quellbytes, Formeln und Verlauf unverändert. Entwürfe
überstehen Zoom und andere Bearbeitungen an derselben Quelle bzw. Auswahl; ein
Ersetzen oder eine andere Auswahl verwirft sie.

Beide Steuerelemente erfordern ein quellbasiertes VSDX und eine gewöhnliche
lokale, ebenenlose, unverklebte 2D-Form mit gespeicherten Pins. Gruppen, Master,
Fremdobjekte und Mehrfachauswahl sind nicht verfügbar. Die Eignung der Szene ist
vorläufig: Der Kern prüft bei der Übermittlung Quellgeometrie, Schutz, Formeln
und Abhängigkeiten. Eine Ablehnung zeigt einen Fehler an, ohne Quellbytes oder
Verlauf zu ändern. Abbruch, Entsorgen und neuere Absichten bezüglich Quelle,
Seite oder Auswahl verhindern, dass eine ausstehende Bearbeitung angenommen wird.

Der Kern-Helfer `visioSizePositionState(page, shapeId)` liefert den
eingefrorenen numerischen Zustand. `visioSizePositionCommand(page, shapeId, field, value)`
liefert gewöhnliche Bearbeitungen, ein leeres Array für exakte Eingaben ohne
Wirkung oder `undefined` für einen ungültigen oder nicht unterstützten Kandidaten.
Eigene Steuerelemente können diese Helfer mit der bestehenden atomaren
Bearbeitungs-API nutzen. Einheitenstrings, Formeleingabe und das Ändern der
Pin-Position werden vom numerischen Bereich nicht unterstützt.

## Diagrammtext suchen {#find-diagram-text}

Die gemeinsamen Suchsteuerelemente finden Literaltext über sichtbare Formen und
Seiten hinweg. Programmatische Hosts nutzen `controller.setSearchQuery(query)`,
`nextSearchResult()`, `previousSearchResult()` und `selectSearchResult(index)`;
die unveränderliche Anfrage und die Ergebnisse stehen in
`controller.state.search`. Eine Suchanfrage lässt die Auswahl unberührt, bis
ausdrücklich navigiert wird. Ein neues Dokument leert die Suche, und neuere
reentrante Host-Aktionen verdrängen die alte Ergebnisnavigation.

Die Suche liefert ein Ergebnis pro passender Form, mit einer begrenzten
Textvorschau. Sie verwendet den ECMAScript-Vergleich in Kleinschreibung, keine
regulären Ausdrücke und keine sprachabhängige Sortierung. Die Limits sind 256
Zeichen pro Anfrage, 500 passende Formen, 32.768 indizierte Zeichen pro Form und
2 Millionen indizierte Zeichen insgesamt. Der Status kennzeichnet Teilergebnisse
ausdrücklich. Ausgeblendete Teilbäume und unterdrückter Gruppentext sind
ausgeschlossen; Hintergrundseiten werden einmal als eigene Seiten indiziert. Die
Hervorhebung einzelner Vorkommen und die Suche in Shape-Daten sind nicht
implementiert.

## Ebenen anzeigen {#display-layers}

Nutzen Sie die gemeinsamen Ebenen-Steuerelemente oder
`setLayerVisibility(pageId, layerId, visible)` jedes eingehängten bzw. nativen
Handles, um die Anzeigesichtbarkeit zu überschreiben. `visible: null` stellt ein
gespeichertes Flag wieder her; `resetLayerVisibility(pageId?)` setzt eine
Quellseite oder das gesamte Dokument zurück. Quellseiten-IDs unterscheiden
Vordergrund- und Hintergrundebenen. Eingefrorene Überschreibungen stehen in
`controller.state.layerVisibilityOverrides`.

Ebenenänderungen synchronisieren Rendering, Auswahl und Suche, ohne das Dokument
zu ändern. Hilfslinien, NoShow und unvollständig beschriebene ältere
ausgeblendete Formen lassen sich durch eine Ebenenüberschreibung nicht
einblenden. Das barrierefreie Panel zeigt höchstens 200 eindeutige Ebenen und
meldet eine Kürzung; die API erlaubt höchstens 25.000 dokumentbezogene
Überschreibungen. SVG- und Druckartefakte verwenden weiterhin die gespeicherte
Sichtbarkeit. Ebenenfarben, Bearbeitung und gespeicherte Druckrichtlinien sind
nicht implementiert.

## Add-in-Registerkarten im Menüband {#ribbon-add-in-tabs}

Ein Host kann nach Hilfe Registerkarten hinzufügen, so wie es ein Add-in in
Visio tut. `ribbonAddIns` ist eine Eigenschaft des Elements und eine Prop jedes
Framework-Bindings, und `office-ribbon-add-in` gehört zur Ereigniszuordnung:

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() }],
			},
		],
	},
];

<VisioViewer
	ribbonAddIns={tabs}
	events={{ 'office-ribbon-add-in': ({ tab, command }) => track(tab, command) }}
/>;
```

Vue nimmt `:ribbon-add-ins` entgegen und sendet `office-ribbon-add-in`; Angular
nimmt `[ribbonAddIns]` entgegen und hat die Ausgabe `(ribbonAddIn)`; das
Vanilla-Binding nimmt `mountViewer(host, { ribbonAddIns })` und
`viewer.update({ ribbonAddIns })` entgegen. Am Element selbst setzen Sie
`element.ribbonAddIns` und hören auf das DOM-Ereignis.

Ein Befehl führt seinen `run`-Callback aus und löst dann das Ereignis aus.
Lassen Sie die Prop weg, bleiben die aktuellen Registerkarten erhalten; ein
leeres Array entfernt sie. Wird ein gleichwertiger Deskriptor erneut übergeben,
wie es ein Framework bei jedem Rendern tut, bleiben die Panels erhalten und nur
die Callbacks werden ersetzt. Der Deskriptor (`RibbonAddInTab`, aus
`ooxml-core/ribbon`, auch von `ooxml-ui` und jedem Framework-Paket exportiert)
wird mit den Editoren für Word, Excel und PowerPoint geteilt. Eine
Registerkarte darf nicht die ID einer eingebauten Registerkarte verwenden
(`home`, `insert`, `design`, `data`, `process`, `review`, `view`, `help`).

## Anpassung {#customization}

`customization` nimmt ein `VisioCustomization`-Objekt entgegen, den gemeinsamen Vertrag von `ooxml-core/embedding`, den auch die anderen Editoren verwenden. Jedes Binding reicht es als Prop bzw. Input `customization` weiter. Es wird angewendet, wenn sich die Objektidentität ändert; halten Sie es daher stabil. `null` stellt die Standardwerte wieder her.

```ts
const customization = {
	capabilities: { export: false, print: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send to chat', onSelect: ({ shapes }) => send(shapes) }],
	},
};
```

- `capabilities`: `false` entfernt jeden Weg zur Funktion (`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`): Menüband, Seiten im Datei-Register, Menüs und Tastenkürzel. `hiddenActions` bleibt unverändert; beide werden vereinigt.
- `backstage.extraPages`: Host-Seiten `{ id, label, group?, order?, render }` im Datei-Register, nach den eingebauten Seiten ihrer Gruppe. `render` ist eine Mount-Funktion, ein Custom-Element-Tag oder `{ element, properties }` und erhält einen `VisioRegionContext`.
- `contextMenu.extraCommands`: Host-Befehle `{ id, label, icon?, group?, when?, disabled?, onSelect }` oberhalb (`top`, Standard) oder unterhalb (`bottom`) der eingebauten Einträge. `onSelect` läuft nach dem Schließen des Menüs mit einem `VisioMenuContext`. Beschriftungen werden unverändert angezeigt (lokalisieren Sie sie selbst).

`edit: false` ist eine eigene Schreibsperre neben dem schreibgeschützten Schalter des Hosts und einer Zusammenarbeitsrolle `viewer`: Keine hebt eine andere auf. Capabilities sind eine UI-Richtlinie, kein Dokumentschutz. Das Ersetzen von Bereichen (Optionsdialog, Bereiche) wird für Visio noch nicht angeboten. Alle Produkte im Überblick: [Einbettungsleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Statischer SVG-Export {#static-svg-export}

`exportPageSvg(model, pageIndex, { maxBytes })` und das `exportSvg()` jedes
eingehängten bzw. nativen Handles liefern einen Snapshot der aktuellen Seite,
ohne Dokument, Auswahl oder Viewport zu ändern. Der Export umfasst unterstützte
Hintergründe, eingebettete Rasterressourcen und sichtbare Kompatibilitätshinweise
in Beschreibung und Metadaten des SVG. Der SVG-Export hat keine entfernten
Assets und keine automatischen Downloads. Die Demo lädt erst nach einem
ausdrücklichen Klick auf die Schaltfläche herunter.

Die UTF-8-Obergrenze beträgt 16 MiB; Aufrufer können sie senken. Konservative
Vorabprüfungen können Inhalte ablehnen, deren spätere Serialisierung kleiner
wäre. Schriften werden nicht eingebettet, und Text und Layout bleiben
näherungsweise. Die physischen Seitenabmessungen bleiben erhalten; nachgelagerte
Rasterisierer müssen ihre Ausgabeabmessungen und Pixelfläche begrenzen, bevor
sie eine Bildfläche anlegen. PDF, Bitmap-Export und Drucklayout bleiben
unimplementiert.

Optionale sekundäre Pixeltests unter Linux laufen nach dem Build mit
`node scripts/test-svg-rasterization.mjs`, mit dem System-Python 3,
librsvg/Cairo und der nativen Canvas-Abhängigkeit des Kerns. Diese Prüfungen
sind von der Validierung im Browser oder mit Microsoft Visio getrennt.

## Artefakte für Zeichnungsseiten vorbereiten {#prepare-drawing-page-artifacts}

`createPrintSnapshot(model, { pageIndices: [0, 2] })` erhält die ausdrücklich
angegebene Seitenreihenfolge und liefert tief eingefrorene, getrennte
SVG-Datensätze mit Zeichnungsabmessungen und allen Export-Diagnosen. Das
`createPrintSnapshot()` jedes eingehängten bzw. nativen Handles erfasst dessen
aktuelle Seite. Es lehnt einen Dokumentwechsel während der Vorbereitung ab, ohne
den Ansichtszustand zu ändern. Es gibt keine Frames, Dialoge, Downloads oder
Druckerbefehle.

Der Snapshot nutzt die gespeicherte Anzeigedarstellung. Ebenen-Eigenschaften
wie Drucken und NonPrinting sowie gespeicherte Druckereinstellungen werden nicht
angewendet; Zeichnungsabmessungen sind weder Papierformat noch Druckmaßstab.
Das Ergebnis nennt diese Einschränkungen. Schutzmechanismen begrenzen die
gewählten Seiten, aggregierte Bytes, wiederholte Hintergründe, Rasterabmessungen,
Text- und Pfadarbeit sowie wiederholte Validierung des gesamten Modells.
Aufrufer können die Limits senken. Getrennte SVGs müssen isoliert bleiben, weil
sich ihre internen Ressourcen-IDs wiederholen können. Dies ist die Vorbereitung
eines künftigen Druck-Workflows, keine native Visio-Druckunterstützung.

## Suchen und Ersetzen mit Geltungsbereich {#scoped-find-and-replace}

Start > Suchen > Ersetzen und Strg+H öffnen die gemeinsamen Ersetzen-Steuerelemente.
Weitersuchen, Zurück, Ersetzen und Alle ersetzen verwenden vollständige,
literale, groß- und kleinschreibungsabhängige Vorkommen. Die Geltungsbereiche
sind Auswahl, Aktuelle Seite und Alle Seiten. Die Auswahl umfasst ausgewählte
Gruppen-Teilbäume in der Reihenfolge von Auswahl bzw. Quelle, wobei sich
überlappende Ziele dedupliziert werden. Seitenbereiche verwenden die gespeicherte
Quellreihenfolge, einschließlich ausgeblendeten Quelltexts. Die Navigation
springt am Ende an den Anfang und wählt nach Möglichkeit ein sichtbares Ziel aus.
„Groß-/Kleinschreibung beachten“ ist aktiviert und derzeit fest; ganze Wörter,
Groß-/Kleinschreibungs-Faltung und spezielle Suchcodes sind nicht implementiert.
Diese Regeln für Geltungsbereich und Durchlauf sind begrenzte Anwendungsrichtlinien
und keine Aussage über die Gleichwertigkeit mit dem nativen Suchdialog.

Jedes native Binding stellt denselben `controller` bereit: Hosts können
`captureTextReplaceToken(scope)`, `getTextReplaceOccurrences(token, query)`,
`selectTextReplaceOccurrence(token, occurrence, query)`,
`planTextReplacement(token, { query, replacement, mode, current? })` und
`applyTextReplacePlan(plan, token)` verwenden. Navigation und Anwendung liefern
ein neues opakes Token. Behalten Sie dieses Token, um beim Wechsel zwischen
Treffern den ursprünglich ausgewählten Geltungsbereich zu bewahren. Manuelle
Änderungen von Auswahl oder Seite, das Ersetzen der Quelle oder eine andere
Operation machen frühere Tokens ungültig. `cancelTextReplace(token)` bricht nur
die aktive Anwendung dieses Tokens ab; neuere, unabhängige Bearbeitungen laufen
weiter. Pläne und Tokens gehören ihrem Controller und können nicht gefälscht
oder über Instanzen hinweg wiederverwendet werden.

Das Ersetzen nutzt eine atomare Quellbearbeitung und eine Rückgängig-Kante,
auch bei seitenübergreifender Ergebnisnavigation. Die nativen Ereignisse
`document-change`, `selection-change` und `shape-select` bleiben über die
Framework-Callbacks verfügbar. „Alle ersetzen“ erhält die Auswahl; das Ersetzen
des aktuellen Vorkommens springt zum nächsten. Ein Ersetzen durch denselben Text
erhält die exakten Quellbytes und fügt dem Verlauf nichts hinzu, navigiert aber
weiterhin. Nicht unterstützter passender Rich Text, Text mit Feldern,
geschützter oder abhängiger Text lehnt den gesamten Stapel
ab. Der Entwurf zeigt den Fehler an und bleibt bearbeitbar. Das Schließen der
Steuerelemente oder ein Wechsel der Absicht bricht die ausstehende Operation ab.

## Experimentelle lokale Nur-Text-Bearbeitung {#experimental-local-plain-text-editing}

Nach `load(bytesOrBlob)` wählen Sie eine lokale Form aus und nutzen die
gemeinsamen Textsteuerelemente, oder Sie rufen `replacePlainText(pageId, shapeId, text)`
an einem beliebigen eingehängten bzw. nativen Handle auf. `undo()` und `redo()`
nutzen den begrenzten Dokumentverlauf; `cancelEdit()` bricht ausstehende Arbeit
ab. Alle sechs Adapter leiten dieselben Methoden und das Ereignis
`document-change` weiter, dessen Detail `{ document, dirty, kind }` (kind `edit`, `undo` oder `redo` sowie `remote` und `save`; siehe oben)
lautet. Prüfen Sie `controller.state.edit` auf Quellverfügbarkeit, Beschäftigt-
und Geändert-Status, Verfügbarkeit von Rückgängig/Wiederholen, Kürzung des
Verlaufs, Fehler und Diagnosen.

Quellbasierte VSDX-Dokumente, die der Viewer geladen oder erstellt hat, lassen
sich bearbeiten oder als VSDX exportieren.
Importe des binären VSD v11 nutzen den veröffentlichten ole2-Codec über
ooxml-core: Unterstützte explizite Seiten- und Formtransformationen,
Verschiebe- und Liniengeometrie sowie Nur-Text werden mit diagnostizierten
Ersatzstilen als Vorschau dargestellt. Die ursprünglichen Legacy-Bytes bleiben
privat erhalten; Bearbeitung, Rückgängig/Wiederholen und VSDX-Export sind für
Legacy-Quellen deaktiviert. Frühere Binärversionen, Gruppen und Master,
komplexe Geometrie, Schablonen, Vorlagen und makroaktivierte Formate bleiben
nicht unterstützt. Die Zuweisung eines Modells mit `document` liefert keine
bearbeitbaren Paketbytes. Die Kern-Validierung lehnt nicht unterstützten
Rich Text, Felder, signierte Pakete und makroaktivierte
Inhalte ab. XML- und Paketbearbeitungen laufen in einem isolierten Worker;
Framework-Wrapper implementieren keine Formatlogik. Nur-Text-Bearbeitungen
verweigern betroffene oder unbekannte textabhängige Formel-Caches, statt sie
veraltet zu belassen, und beachten den effektiven Schutz `LockTextEdit`. Der
logische VSDX-Text schließt den gespeicherten abschließenden Absatzmarker aus,
behält aber beabsichtigte nachgestellte LF-Zeichen und leere Absätze bei. Die
Textsemantik des Legacy-VSD bleibt unverändert. Die Bearbeitung von Rich-Text-Inhalten
bleibt nicht unterstützt.

`exportVsdx()` liefert `{ bytes, dirty, diagnostics }` für eine ausdrücklich
heruntergeladene Kopie. Es überschreibt die Quelldatei nicht, lädt sie nicht
hoch und behauptet keine native Round-Trip-Treue. Die gemeinsame Oberfläche
lädt erst nach der ausdrücklichen Aktion des Benutzers herunter. Der Verlauf
ist begrenzt und kann ältere Rückgängig-Zustände verwerfen, was
`historyTruncated` meldet. Das Ersetzen von Nur-Text bleibt auf die oben
zugelassene Teilmenge beschränkt. Für die folgenden Formatierungs- und
Geometrieoperationen gelten eigene Zulassungsregeln. Die Nachweise zum nativen
Wiederöffnen decken bestimmte aufgezeichnete Fälle ab und stehen im
[Verifikationsprotokoll](/verification); sie belegen keine allgemeine Treue.

## Experimentelle Formatierung und Anordnung {#experimental-formatting-and-arrangement}

Das Menüband Start und `applyEdits()` teilen sich die Befehle `format-text`,
`format-shape` und `reorder-shape`. Die Textformatierung umfasst gespeicherte
Schriftfamilien, Punktgröße, Fett/Kursiv/Unterstrichen/Durchgestrichen, Farbe,
Aufzählungszeichen, Einzug sowie horizontale und vertikale Ausrichtung. Die
Formformatierung umfasst Volltonfüllung sowie Linienfarbe und -stärke,
Linienmuster 0-23, Füllmuster 0-24, Hintergrundfarbe und Transparenz. Die
Formatierung erfordert gewöhnliche lokale Blattformen. Die Textformatierung der
gesamten Form aktualisiert jede effektive Zeichen- bzw. Absatzzeile und erhält
dabei Laufmarker, Text, nicht betroffene Stilbits und unterstützte
zwischengespeicherte Felder. Zeichen- und Absatzformatierung haben getrennte
Quellzulassungen; die vertikale Ausrichtung schreibt keinen der beiden
Abschnitte um. Geschützte oder mehrdeutige Zeilen und betroffene
Formelabhängigkeiten lehnen die gesamte Transaktion ab. Das Vergrößern und
Verkleinern der Schriftgröße erfordert eine einheitliche Größe; die
Größenauswahl kann eine Auswahl mit gemischten Größen festlegen. Lokale
Zellen `F="Inh"` für Zeichen und Absatz sind bearbeitbar, wenn sich ihr Vorfahr
und der zwischengespeicherte Wert bzw. die Einheiten belegen lassen. Eine
Delegation ohne Cache erfordert einen echten Vorfahren; eine ausdrückliche
Formatierungswahl erzeugt eine lokale Überschreibung, auch wenn der Wert
unverändert ist. Geerbte Schutz- und Transformationszellen der obersten Ebene
bleiben geschützt. Schriften müssen bereits in der Quellzeichnung vorhanden sein.
Die Stapelbefehle unterstützen Ganz nach vorn/hinten sowie Eine Ebene nach
vorn/hinten für eine gewöhnliche Form der obersten Ebene im Anzeigeband null.

`format-shape` akzeptiert `linePattern`, `fillPattern`, `fillBackgroundColor`
(`'#RRGGBB'`), `lineTransparency` und `fillTransparency` (Prozent von 0 bis 100).
Die Transparenz wird in Halbprozentschritten gerundet, wobei Gleichstände
aufgerundet werden. Die Füllungstransparenz setzt Vorder- und Hintergrund
gemeinsam. Muster 0 bedeutet keine Farbe; Muster 1 ist Vollton. Ein explizites
`linePattern: 1` aktiviert eine Kontur unabhängig von der Linienfarbe. Ein
explizites Füllmuster überschreibt das durch `fillColor` implizierte
Vollton-Muster; `fillColor: 'none'` mit einem Füllmuster ungleich null wird
abgelehnt. Eine aktive Verlaufstransparenz erfordert das ausdrückliche Ersetzen
der Füllung, während gewöhnliche geerbte Volltonfüllungen aus dem Design
unterstützt werden.

Das gemeinsame Linienmenü bietet „Keine Linie“ und die eingebauten Muster. Die
Einträge „Weitere Optionen“ beider Farbmenüs und das Kontextmenü der Form öffnen
den Aufgabenbereich „Form formatieren“, der jede Änderung sofort als eine
Auswahltransaktion anwendet. Gemischte Quellwerte erscheinen leer, bis sie
geändert werden. `fillColorTheme`, `lineColorTheme` und (bei `format-text`)
`fontColorTheme` benennen das Farbfeld aus den Designfarben, aus dem eine Farbe
stammt, sodass sie als Design-Formel von Visio gespeichert wird; `beginArrow`,
`endArrow`, `beginArrowSize`, `endArrowSize`, `lineCap` und `rounding` legen
Linienenden, Linienabschluss und Eckenrundung fest. Alle sechs nativen Bindings
nutzen das bestehende `applyEdits()` und die reaktiven Dokument- und
Auswahlereignisse; es wird keine frameworkspezifische Farb-Engine hinzugefügt.
`visioShapeFormattingState(shapes)` liefert aggregierte Farbwerte der Quelle für
eigene Steuerelemente. Diese Werte sind von den Deckkraft-Multiplikatoren für
Verläufe und Kacheln getrennt.

Sechs Ausrichten-Befehle richten sich an der zuerst ausgewählten Form aus. Die
Position sorgt bei drei oder mehr Formen für gleichmäßige horizontale und
vertikale Kantenabstände. Der Kern-Helfer
`visioArrangeCommands(page, selectedIds, action)` erzeugt atomare Bearbeitungen
für Hosts mit eigenen Steuerelementen. Er misst gedrehte Breiten- und Höhenboxen,
ordnet die Verteilung nach Mittelpunkten und erhält bei Gleichstand die
Auswahlreihenfolge. Gruppen, Master, Ebenen und verklebte Verbinder liegen
außerhalb seines Zulassungsbereichs. Diese Aktionen teilen sich Verlauf,
Beibehaltung der Auswahl und VSDX-Download bzw. -Wiederöffnen mit anderen
Bearbeitungen.

## Experimentelles Duplizieren {#experimental-duplication}

`duplicateSelection(): Promise<void>` am Controller und an jedem eingehängten
bzw. nativen Handle dupliziert zugelassene ausgewählte Formen der aktuellen
Seite. Das Einfügen-Menü auf der Registerkarte Start, das Kontextmenü der Form,
„Sagen Sie mir, was Sie tun möchten“ und Strg+D rufen dieselbe Methode auf.
Texteingaben behalten ihr natives Tastaturverhalten. Das Duplizieren nutzt nicht
die System-Zwischenablage.

Der Kern-Helfer `visioDuplicateCommand(page, selectedIds)` erzeugt einen
Quellbefehl `duplicate-shapes` mit neuen seitenlokalen IDs. Kopien behalten
Quell-XML, gemischten Text und Stapelreihenfolge, erhalten neue Namen ohne
kopierte eindeutige IDs und werden um 0,33 Zeichnungszoll nach rechts und unten
verschoben. Die Quellvalidierung lässt derzeit lokale, unverklebte
2D-Blattformen zu und schließt Master, Gruppen, Fremdinhalte, Ebenen, mehrdeutige
Identitäten und nicht unterstützte Abhängigkeiten aus. Pin-Schutz und betroffene
Formel-Caches werden geprüft, bevor eine Quelländerung angenommen wird.

Die entstandenen Kopien werden in einem atomaren Verlaufsschritt zur Auswahl.
Rückgängig stellt die ursprüngliche Auswahl wieder her, Wiederholen die Kopien,
sofern keine neuere Auswahl- oder Seitenaktion dies verdrängt. Fehlgeschlagene,
abgebrochene oder überholte Operationen können Quellbytes oder Auswahl nicht
teilweise ändern. Die gewöhnlichen Ereignisse `selection-change` und
`document-change` machen das akzeptierte Ergebnis dem nativen Framework-Zustand
zugänglich.

## Experimentelle Formen-Zwischenablage {#experimental-shape-clipboard}

Jedes eingehängte bzw. native Handle stellt `copySelection()`, `cutSelection()`
und `pasteSelection()` bereit, die jeweils `Promise<void>` liefern. Rufen Sie sie
aus einer Benutzergeste auf, damit der Browser den Zugriff auf die
Zwischenablage gewähren kann. Start, Kontextmenüs, „Sagen Sie mir, was Sie tun
möchten“ und Strg/Meta+C, X und V nutzen denselben gemeinsamen Adapter.
Texteingaben behalten das native Zwischenablageverhalten. Fehler der
Zwischenablage erscheinen im Viewer, ohne Quellbytes, Verlauf oder Auswahl zu
ändern. Ausschneiden schreibt in die Zwischenablage, bevor gelöscht wird; verhindert
ein Schutz das Löschen, bleiben die kopierten Formen in der Zwischenablage, und der
Viewer meldet, dass sie nicht ausgeschnitten werden konnten.

Die Nutzlast ist begrenzter `text/plain`-Text mit dem Marker
`OOXML-VISIO-SHAPES/1`. Sie enthält das erfasste Form-XML und den benötigten
Quellkontext, sodass das Bearbeiten oder Ausschneiden der Originale den
kopierten Inhalt nicht ungültig macht. Das Einfügen liest immer die aktuelle
System-Zwischenablage; es verwendet nicht stillschweigend eine interne frühere
Kopie wieder. Native Zwischenablageformate von Microsoft Visio, beliebiger Text,
Bilder und Inhalte einfügen sind nicht unterstützt. Eingefügte Formen erhalten
neue IDs und Namen sowie einen Versatz von 0,33 Zeichnungszoll nach rechts und
unten. Wiederholtes Einfügen nutzt denselben Versatz; die native Kaskade beim
Einfügen und die seitenübergreifende Platzierung sind nicht belegt.

Der erste unterstützte Umfang sind gewöhnliche lokale, unverklebte
2D-Blattformen mit übereinstimmenden Schrift-, Stil- und Design-Definitionen
des Dokuments sowie übereinstimmenden Seiteneinstellungen, Ordnungszahl,
Hintergrundkontext und Zeichnungsverhältnis. Verweise müssen innerhalb der
kopierten Auswahl bleiben. Gruppen, Master, Ebenen, Fremdinhalte, Fragmente mit
Beziehungen, nicht unterstützte Formeln und der Ressourcenimport bleiben
abgelehnt. Der Kern validiert die gesamte Nutzlast vor einer atomaren
Quellbearbeitung. Das Einfügen wählt seine Kopien aus und teilt sich mit
Duplizieren die Regeln für begrenzten Verlauf und Auswahlabsicht.

`controller.state.clipboard` stellt den unveränderlichen Zustand `ready`,
`preparing` und `error` über das bestehende reaktive Abonnement des jeweiligen
Frameworks bereit. Die Quellerfassung läuft in einem abbrechbaren Worker und
wird durch Änderungen von Auswahl, Seite, Sichtbarkeit oder Dokument ungültig.
Hosts mit eigenem Zwischenablage-Transport können am Controller
`captureClipboardToken()`, `prepareClipboardSelection(token)`,
`getPreparedClipboard(token)`, `cutPreparedSelection(token)` und
`pasteClipboardText(text, token)` verwenden. Tokens lehnen veraltete Aktionen
ab. Der DOM-freie Kern exportiert `captureVisioClipboard`,
`serializeVisioClipboard`, `deserializeVisioClipboard`, die Bearbeitung
`paste-shapes` und den Befehlsbauer `visioPasteCommand(page, clipboard, offset?)`.

## Experimentelle Geometriebearbeitung {#experimental-geometry-editing}

Die gemeinsamen Geometrie-Steuerelemente bieten Rechteck erstellen, Verschieben,
Größe ändern und sicheres Löschen. Jedes native bzw. eingehängte Framework-Handle
leitet außerdem `applyEdits(edits: readonly VisioEdit[])` für einen atomaren
Stapel über denselben Worker, Verlauf und Abbruchpfad weiter.

Start > Textwerkzeug und Einfügen > Textfeld nutzen denselben gemeinsamen
Entwurf eines Textfelds fester Größe. Ziehen Sie über leeren Seitenraum, geben
Sie Nur-Text ein und nutzen Sie dann „Textfeld hinzufügen“ oder Strg/Meta+Eingabe;
Escape bricht ab. Leere Entwürfe erzeugen keine Quellformen; beabsichtigte
Leerzeichen und Zeilenumbrüche bleiben erhalten. Ein Klick mit dem Textwerkzeug
auf eine bestehende Form wählt sie aus und blendet die vorhandenen
Textbearbeitungs-Steuerelemente ein. Das Zeichnen von Rechtecken, Ellipsen und
Linien nutzt denselben abgesicherten Erstellungs-Lebenszyklus. Ihre interne
Controller-Transaktion erfasst vor dem Zeichnen Quelle, Seite, geordnete
Auswahl, Zoom und Sichtbarkeitsabsicht, prüft diese Absicht vor der
Quellannahme und wählt erstellte Formen im selben Rückgängig-Schritt aus. Eine
neuere Absicht oder das Entsorgen verdrängt eine ausstehende Erstellung. Interne
Erstellungs-Tokens sind keine zusätzliche native Binding-API.

Alle sechs öffentlichen Handles akzeptieren den neuen gewöhnlichen Befehl über
`applyEdits()`:

```ts
await handle.applyEdits([
	{
		type: 'create-text-box',
		pageId: '0',
		shapeId: '10',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'Text with an intentional trailing blank paragraph\n',
	},
]);
```

Text ist erforderlich. Der gespeicherte Standard-Textstil liefert die anfängliche
Zeichenformatierung. Explizites `FillPattern=0` und `LinePattern=0` machen das
neue Feld farblos; das Ändern allein der Linienfarbe aktiviert keine Kontur.
Das Erstellen von Entwürfen fester Größe belegt weder native Standardwerte des
GUI-Textwerkzeugs noch automatische Textgrößenanpassung oder die Bearbeitung von
Rich-Text-Inhalten. Das generische `applyEdits()` behält sein bestehendes
Auswahlverhalten; das gemeinsame Zeichenwerkzeug verantwortet die Auswahl
erstellter Formen. Native Nachweise decken 24 gespeicherte API-Fälle mit festen
Feldern bei drei Zeichnungsmaßstäben ab, unter Verwendung der Standard- und
benutzerdefinierten Stile Text Only/Arial 18 pt, wobei die Kern-Ausgaben in
Visio wieder geöffnet wurden. Diese Fälle mit gespeicherten Stilen messen keine
interaktiven GUI-Standardwerte; siehe das [Verifikationsprotokoll](/verification).

Einfügen > Neue Seite läuft über den bestehenden quellbasierten Befehl zum
Einfügen von Seiten und wählt dann die neue Seite aus. Es teilt sich Einfügeverhalten,
Rückgängig/Wiederholen und Quellzulassung mit der Seitenleiste; dies fügt keine
vorlagenbasierte Seitenerstellung hinzu.

```ts
await handle.applyEdits([
	{
		type: 'create-rectangle',
		pageId: '0',
		shapeId: '9',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'New shape',
	},
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 6, y: 7 },
]);
```

Koordinaten sind Positionen des Drehpunkts in Zeichnungszoll, Ursprung unten
links, nach oben positiv. Die Größenänderung hält den Pin fest, sofern der
Befehl nicht `anchor: { x, y }` enthält. Jede Anker-Koordinate ist `0`, `0.5`
oder `1` und bezeichnet den festen Punkt in den normalisierten, nach oben
gerichteten lokalen Breiten- und Höhengrenzen der Form. Zum Beispiel hält
`{ x: 0, y: 0 }` die lokale linke untere Ecke fest. Die verankerte Größenänderung
hat einen eigenen Quellnachweis für lokale Pins, betroffene Caches und feste
Dreh- bzw. Spiegelungswerte. IDs sind explizit. Bestehende zugelassene lokale
2D-Formen der obersten Ebene können bearbeitet werden, auch solche, die von
anderen Erzeugern importiert wurden. Der Kern wertet unterstützte betroffene
numerische ShapeSheet-Abhängigkeiten aus und erhält unberührte ZIP-Nutzlasten.
Der Viewer verändert XML nie selbst.

Dies ist eine schmale zugelassene Teilmenge. Master, Gruppen und Fremdformen,
Verbinder und Klebungen, unsicherer Schutz bzw. unsichere Umleitung, nicht
unterstützte betroffene Formeln, mehrdeutige Paketabhängigkeiten und das Löschen
referenzierter Formen schlagen mit einem sichtbaren Fehler fehl. Relative
Liniengeometrie wird skaliert; absolute Liniengeometrie benötigt eine
unterstützte Dimensionsabhängigkeit. Den Befehlsvertrag, numerische Limits,
Ausschlussgründe und nächste Erweiterungen beschreibt die
[`src/core/visio/README.md`](https://github.com/ChristopherVR/ooxml/blob/main/src/core/visio/README.md)
des Kerns. Das [Funktionsverzeichnis](/parity) hält die aktuell zugelassene
Menge fest, die inzwischen auch begrenzte primitive Verschiebungen von
Master-Instanzen mit expliziten lokalen Pins umfasst.

Tests mit generierten bzw. importierten einfachen Formen und ein lokaler
Chrome-Workflow bestehen. Die 19 akzeptierten Diagramme des öffentlichen Korpus
lassen derzeit null Geometriebearbeitungen zu, weil sich der Abhängigkeitsbereich
außerhalb der Seite noch nicht als unabhängig nachweisen lässt. Bei dreizehn
fehlt außerdem eine geeignete lokale Form. Die praktische Bearbeitungsabdeckung
für diesen Korpus bleibt blockiert; die nächsten Arbeiten sind ein
Abhängigkeitsgraph für Paket, Master und Design mit Bereichsbezug und die
Bearbeitung von Master-Instanzen, gefolgt vom Routing verklebter Endpunkte. Das
native Wiederöffnen wurde inzwischen für die konkreten Aufzeichnungen zu
Formatierung und Anordnung im [Verifikationsprotokoll](/verification) verifiziert;
die allgemeine Treue bleibt unverifiziert.
