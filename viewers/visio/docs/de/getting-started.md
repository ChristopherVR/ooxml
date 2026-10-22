# Erste Schritte {#getting-started}

Öffnen Sie ein Diagramm lokal, verstehen Sie die aktuellen Grenzen und binden Sie denselben Viewer in Ihre Anwendung ein.

::: warning Erste Viewer-Unterstützung, keine Visio-Zertifizierung
Dies ist eine öffentliche Beta. Werten Sie ein erfolgreich geöffnetes Dokument nicht als Beleg dafür, dass jede Form, jede Formel oder jedes Formatmerkmal originalgetreu dargestellt wird. Siehe das [Funktionsverzeichnis](/parity).
:::

## Das lokale Playground ausprobieren {#try-the-local-playground}

1. [Öffnen Sie das Playground](/demo/){target="_self"}, um ein originales Beispieldokument zu sehen.
2. Wählen Sie über die Dateiauswahl oder den Ablagebereich eine lokale `.vsdx`-Datei.
3. Wechseln Sie die Seiten, passen Sie den Zoom an, lassen Sie die Seite einpassen und wählen Sie eine Form aus.
4. Lesen Sie die Kompatibilitätshinweise. Vergleichen Sie wichtige Dateien mit ihrer Darstellung in Microsoft Visio.

Es wird kein Dokument hochgeladen. Die Steuerelemente unterstützen Suche, Anzeigeebenen, experimentelle Nur-Text-Bearbeitung, begrenztes Rückgängig/Wiederholen, SVG-Export und ausdrückliche Downloads von VSDX-Kopien. Druck-Snapshots sind nur eine API-Vorbereitung. Datei > Exportieren kann ein PDF aus Bildseiten erzeugen, aber vektorbasiertes oder textauswählbares PDF, allgemeines Zeichnen und die meisten Funktionen des älteren binären `.vsd`-Formats liegen außerhalb dieses Builds.

Die Builds desselben Playgrounds für die anderen Frameworks sind auf [der Seite der Demos](/demos) aufgeführt.

## Ein Framework-Paket installieren {#install-a-framework-package}

Installieren Sie ein Viewer-Paket für Ihr Framework. Jedes enthält den gemeinsamen Viewer, seinen Adapter, die Worker und die Dokument-API.

::: code-group

```bash [React]
npm install visio-react-viewer react
```

```bash [Vue]
npm install visio-vue-viewer vue
```

```bash [Angular]
npm install visio-angular-viewer
```

```bash [Svelte]
npm install visio-svelte-viewer svelte
```

```bash [Solid]
npm install visio-solid-viewer solid-js
```

```bash [Vanilla JS]
npm install visio-vanilla-viewer
```

:::

| Framework | Paket                  | Export                 |
| --------- | ---------------------- | ---------------------- |
| Vanilla   | `visio-vanilla-viewer` | `mountViewer`          |
| React     | `visio-react-viewer`   | `VisioViewer`          |
| Vue 3     | `visio-vue-viewer`     | `VisioViewer`          |
| Angular   | `visio-angular-viewer` | `VisioViewerComponent` |
| Svelte 5  | `visio-svelte-viewer`  | default `VisioViewer`  |
| Solid     | `visio-solid-viewer`   | `VisioViewer`          |

Alle sechs Adapter reichen denselben Controller, dieselben Eigenschaften, Ereignisse und dasselbe imperative Handle über den gemeinsamen Viewer weiter. Siehe die [Framework-Anleitungen](/frameworks/react) und den [Binding-Vertrag](/bindings).

## Einen Viewer einhängen {#mount-a-viewer}

Der Container muss eine Höhe haben. Ein Browser-`File`, `Blob`, `Uint8Array` oder `ArrayBuffer` kann an `load` übergeben werden.

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	events: { 'shape-select': (shape) => console.log(shape) },
});

await viewer.load(file); // A local File or bytes
viewer.fit();

// When your view unmounts:
viewer.destroy();
```

## Den Entwicklungs-Workspace ausführen {#run-the-development-workspace}

Der Viewer gehört zum `ooxml`-Monorepo, einem Bun-Workspace. Installieren Sie einmal im Repository-Stamm mit `bun install`; `ooxml-core` (`src/core/visio`) und `ooxml-ui` (`src/ui/src/visio`) werden zu den Paketen im selben Checkout aufgelöst, ein Verknüpfungsschritt ist nicht nötig.

```bash
bun install
cd viewers/visio
bun run check
bun run dev
```

Verwenden Sie die von Vite ausgegebene lokale URL. Dies startet einen Entwicklungsserver; die Website wird dadurch nicht veröffentlicht.

## Experimentelle Nur-Text-Bearbeitung {#editing}

Laden Sie ein lokales Quellpaket, wählen Sie eine unterstützte Form aus und nutzen Sie deren Textsteuerelemente. `replacePlainText(pageId, shapeId, text)`, `undo()` und `redo()` stellen dieselben Operationen über jedes Handle bereit. Ein mit `document` zugewiesenes Modell hat keine Paketbytes und bleibt schreibgeschützt.

Die Bearbeitung lehnt Text mit Masterverknüpfung, Rich Text, Felder, signierte Pakete und Makroinhalte ab. Der Verlauf ist begrenzt und kann alte Rückgängig-Zustände verwerfen. Gespeicherte Formel-Caches werden nicht neu berechnet. Allgemeine Zeichen- und Stilbearbeitung sind nicht implementiert.

`exportVsdx()` liefert die ursprünglichen oder bearbeiteten Paketbytes samt Diagnosen für einen ausdrücklichen Kopie-Download. Die Eingabedatei wird nie überschrieben. Das Wiederöffnen in nativem Visio und eine verlustfreie Round-Trip-Treue sind weiterhin nicht verifiziert.

## Ein statisches SVG exportieren {#export}

`viewer.exportSvg()` liefert den SVG-String der aktuellen Seite, die Abmessungen, die Bytelänge und Kompatibilitätsdiagnosen. Auswahl, Zoom und das Quelldokument bleiben unverändert. Die Schaltfläche „Export SVG“ des Playgrounds löst einen ausdrücklichen lokalen Download aus.

Unterstützte Hintergrundinhalte und validierte Rasterbilder werden eingebettet. Schriften werden nicht eingebettet; Umbruch und andere Näherungen bleiben in der SVG-Beschreibung und den Metadaten sichtbar. Der Export ist auf 16 MiB begrenzt, mit konservativen frühen Ressourcenprüfungen. Aufrufer können das Limit mit `{ maxBytes }` senken. Dies ist keine bearbeitbare VSDX-Ausgabe. Nachgelagerte Rasterisierer müssen Abmessungen und Pixelfläche unabhängig begrenzen.

## Seiten-Snapshots vorbereiten {#prepare-page-snapshots}

`createPrintSnapshot()` liefert über die API unveränderliche SVG-Seitenartefakte und Diagnosen. Die Zeichnungsabmessungen bleiben erhalten, es werden jedoch keine nativen Druckereinstellungen angewendet, kein Druckdialog geöffnet und kein PDF-Export implementiert. Damit werden Daten für einen künftigen Druck-Workflow vorbereitet.

## Eigenschaften und Ereignisse {#properties-and-events}

| Oberfläche    | Aktueller Vertrag                                                                                                                                                      |
| ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Eigenschaften | `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `ribbonAddIns`, `collaboration`                                                                             |
| Ereignisse    | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change` |
| Lebenszyklus  | `update`, `load`, `fit`, `exportSvg`, `createPrintSnapshot`, `replacePlainText`, `undo`, `redo`, `cancelEdit`, `exportVsdx`, `destroy`                                 |

`canEdit` ist die Bearbeitungsberechtigung (der Viewer ist standardmäßig bearbeitbar; `false` macht ihn schreibgeschützt), `collaboration` teilt die geöffnete Zeichnung in einem Raum, und `document-change` meldet `kind` als `edit`, `undo`, `redo`, `remote` oder `save`. Nachdem der Host die Bytes aus `exportVsdx()` gespeichert hat, löscht `controller.markClean()` das Änderungskennzeichen. Eigene Oberflächen aus den Teilen des Viewers bauen Sie mit den [Bausteinen](/building-blocks); zum Teilen einer Zeichnung siehe [Zusammenarbeit](/de/collaboration).

Hängen Sie den Viewer ein, nachdem der Host existiert, und rufen Sie `destroy()` auf, wenn die Framework-Ansicht ausgehängt wird. Mehrfaches Zerstören ist unbedenklich. Teilaktualisierungen lassen ausgelassene Eigenschaften unverändert. Die vollständige Referenz ist die [Viewer-API](/de/api).

## Verifikation vor dem Einsatz {#verification-before-adoption}

Führen Sie `bun run check` und `bun run test:browser` in `viewers/visio` aus. Die Parser-Tests des Kerns liegen bei der Engine in `src/core/visio` (ausführen mit `bun run test` im Repository-Stamm). Diese Prüfungen belegen nur das, was sie ausdrücklich prüfen; sie belegen keine visuelle Übereinstimmung mit Microsoft Visio.
