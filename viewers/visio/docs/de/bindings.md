# Framework-Bindings

Sechs schlanke native Integrationen verwenden ein gemeinsames Browser-Element und eine einzige `mountViewer`-Implementierung. Jede wird als eigenes, in sich geschlossenes Paket veröffentlicht; keines parst Dokumente, zeichnet Diagramme oder besitzt einen eigenen Renderer.

Jedes native Handle stellt `createBlankDrawing(options?: CreateVsdxOptions)` bereit.
Der gemeinsame Controller erstellt die neue Quelle und übernimmt sie über seinen vorhandenen Ladevorgang; Zustands-Hooks und Dokumentereignisse aktualisieren sich wie gewohnt. Zeigerbewegung und Rahmenauswahl laufen für alle Bindings im selben gemeinsamen Web Component.

| Framework | Paket                  | Export                                       |
| --------- | ---------------------- | -------------------------------------------- |
| Vanilla   | `visio-vanilla-viewer` | `mountViewer`                                |
| React     | `visio-react-viewer`   | `VisioViewer`                                |
| Vue 3     | `visio-vue-viewer`     | `VisioViewer`                                |
| Angular   | `visio-angular-viewer` | `VisioViewerComponent` (`visio-viewer-host`) |
| Svelte 5  | `visio-svelte-viewer`  | Standard-Export `VisioViewer`                |
| Solid     | `visio-solid-viewer`   | `VisioViewer`                                |

Importieren Sie nur das Paket des benötigten Frameworks, um nicht unbeteiligte Framework-Laufzeiten zu laden. Die Quelle der Adapter ist `packages/bindings` im Repository; dessen README nennt die Prüfungen, die sie absichern.

Die Verdrahtung ist gemeinsam. Der Vertrag (`VIEWER_CONTRACT`, die Schlüssel für Eigenschaften und Ereignisse sowie die Member des Handles) und die revisionsbewusste Weiterleitung von Patches in `mountFrameworkViewer` liegen in `packages/bindings/src/common.ts`. Die generische Binding-Maschinerie (das verzögerte Handle, die Auswahl der Props sowie die React- und Vue-Komponentenfabriken) ist das `ooxml-ui/bindings`-Kit. React und Vue entstehen aus diesen Fabriken; Angular installiert seine Handle-Member aus `VIEWER_CONTRACT`; Svelte ist eine Komponente; Solid ist eine JSX-Komponente, die das gemeinsame Element direkt einbindet. `packages/bindings/tests/contract-parity.test.ts` prüft jeden Adapter gegen den Vertrag: dieselben Eigenschaften, jedes Ereignis über die Callback-Map und die native Ausgabe sowie jeden Member des Handles.

## Gemeinsamer Vertrag {#shared-contract}

Alle Adapter akzeptieren dieselben Eigenschaften:

- `document`: ein `VisioDocument` oder `null`, um es zu leeren
- `pageIndex`: nullbasierter Seitenindex
- `zoom`: Skalierung, wobei `1` für 100 % steht
- `showToolbar`: Sichtbarkeit der Symbolleiste
- `canEdit`: Bearbeitungsberechtigung. Der vollständige Viewer ist standardmäßig bearbeitbar; `false` macht ihn schreibgeschützt (Auswahl, Suche, Zoom und Export funktionieren weiterhin). Die Bausteine sind dagegen standardmäßig schreibgeschützt, siehe [Bausteine](/de/building-blocks)
- `readOnly`: eine eigene Einschränkung. `true` macht die Zeichnung schreibgeschützt und hat Vorrang vor `canEdit`; `false` oder ausgelassen hebt nur diese Einschränkung auf
- `ribbonAddIns`: Tabs, die der Host nach „Hilfe“ hinzufügt (siehe [Ribbon-Add-in-Tabs](/de/api#ribbon-add-in-tabs))
- `collaboration`: teilt die geöffnete Zeichnung, solange der Wert gesetzt ist, in einem Raum; `null` beendet die Freigabe (siehe [Zusammenarbeit](/de/collaboration))
- `storage` und `customization`: vom Host verwaltetes Speichern und Öffnen sowie das gemeinsame Anpassungsobjekt (siehe [Anpassung](/de/api#customization) und [Vom Host verwalteter Speicher](/api#host-owned-storage-and-file-commands))
- `fileProviders`: Cloud-Speicherorte für Datei > Öffnen, etwa OneDrive (siehe [OneDrive](/de/onedrive))
- `events`: typisierte Callback-Map mit wörtlichen Ereignisnamen

Die Ereignis-Map enthält derzeit `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change` und `file-command`. Sie wird aus dem kanonischen Vertrag abgeleitet. Alle Adapter leiten jedes Ereignis weiter; keiner parst Dokumente oder besitzt einen eigenen Renderer.

Aktualisierungen von Eigenschaften sind partiell. Native Framework-Snapshots werden verglichen, damit nicht zusammenhängende Neurenderings ein imperatives Laden oder einen vom Benutzer geänderten Ausschnitt nicht rückgängig machen. Ausgelassene oder `undefined`-Eigenschaften behalten den aktuellen Viewer-Zustand bei. `null` leert das Dokument ausdrücklich. Das Entfernen der Prop `events` löscht die Handler der Callback-Map. Für ein externes Ersetzen geben Sie ein neues Dokumentobjekt an; direkte Mutation ist kein unterstützter Mechanismus für Reaktivität. Seiten- und Zoom-Interaktionen können den internen Viewer-Zustand ändern; die Props fordern Änderungen erst an, wenn ihr Framework sie aktualisiert.

`document-change` liefert `{ document, dirty, kind }`, wobei `kind` `edit`, `undo`, `redo`, `remote` (die Änderung eines Mitbearbeiters) oder `save` (nach `markClean()`) ist. Das imperative `ViewerHandle` stellt `element`, `controller`, `load(bytesOrBlob)` und `fit()` bereit. Zu den Bearbeitungsmethoden gehören `applyEdits(edits)`, `replacePlainText(pageId, shapeId, text)`, `duplicateSelection()`, `copySelection()`, `cutSelection()`, `pasteSelection()`, `undo()`, `redo()`, `cancelEdit()` und `exportVsdx()`. Nachdem der Host die exportierten Bytes gespeichert hat, ruft er `controller.markClean()` auf. Die Zusammenarbeit wird über die Eigenschaft gesteuert; `element.collaborationState`, `element.reconnectCollaboration()` und `element.stopCollaboration()` sind über das `element` des Handles erreichbar. Der Zustand der Zwischenablage ist über jedes native reaktive Abonnement unter `controller.state.clipboard` verfügbar; Zwischenablage-Methoden benötigen eine Benutzergeste im Browser und unterstützen das begrenzte Format für Formen des Viewers. Sie leiten an dieselbe quellgestützte, im Worker isolierte Implementierung weiter. `ViewerEditState` und `VsdxExportResult` sind gemeinsam exportierte Typen; der Zustand ist unter `controller.state.edit` verfügbar. Angular stellt den nativen Output `documentChange` bereit; jeder Adapter unterstützt das Callback-Map-Ereignis `document-change`. Lesezugriffe und Operationen über ein Framework-Handle schlagen vor dem Mounten oder nach dem Unmounten mit einer klaren Fehlermeldung fehl. Ein zurückgehaltener Controller oder ein Element wird ebenfalls entsorgt, wenn sein Besitzer ausgehängt wird. Vanilla gibt das vollständige gemeinsame Binding mit `update()` und einem idempotenten `destroy()` zurück.

Der Editor selbst wird auf dem Client gemountet. Separate Node-Tests prüfen, dass alle sechs Einstiegspunkte ohne Browser-Globals importiert werden; React, Vue, Svelte und Solid rendern auf dem Server außerdem leere native Hosts. React und Vue haben Tests, die ihr tatsächliches Server-Markup hydratisieren, bevor das gemeinsame Element gemountet wird. Server-Rendering für Angular sowie Hydration für Solid und Svelte sind nicht verifiziert, ebenso das dokumentübergreifende Mounten. Diese Prüfungen verwenden die installierten Versionen und bedeuten keine vollständige Integrationsabdeckung für SSR-Frameworks.

## Native Verwendung {#native-usage}

React verwendet eine weitergeleitete Ref und gewöhnliche Props:

```tsx
const ref = useRef<ViewerHandle>(null);
<VisioViewer ref={ref} document={diagram} events={{ 'document-error': reportError }} />;
```

Vue stellt dasselbe Handle an der Komponenten-Ref bereit und sendet außerdem die nativen Ereignisse in kebab-case:

```vue
<VisioViewer ref="viewer" :document="diagram" @document-error="reportError" />
```

Angular stellt native Inputs und Outputs bereit. Fügen Sie die eigenständige Komponente den Imports der Anwendung hinzu:

```html
<visio-viewer-host [document]="diagram" (documentError)="reportError($event)" />
```

Die Angular-Outputs sind `documentLoad`, `documentChange`, `documentError`, `pageChange`, `zoomChange`, `shapeSelect`, `selectionChange`, `ribbonAddIn`, `collaborationChange` und `fileCommand`. `fileCommand` meldet nur den Befehl; um ihn abzubrechen, gibt ein Angular-Host `false` aus dem Callback `events['file-command']` zurück. Die vollständige Callback-Map `events` wird ebenfalls als Input akzeptiert. Vue und Angular rufen den Handler der Callback-Map auf und senden danach ihren nativen Output; verwenden Sie eine der beiden APIs, es sei denn, beide Benachrichtigungen sind gewünscht.

Verwenden Sie bei großen, unveränderlichen Dokumentmodellen Vue `shallowRef`/`markRaw` und Svelte `$state.raw`, statt ein neu erzeugtes Optionsobjekt bei jeder Aktualisierung tief zu proxyen. Ein Dokument in einen neuen reaktiven Proxy einzuwickeln ändert seine Identität und fordert ein externes Ersetzen an. Direkte Mutation wird nicht unterstützt.

Svelte verwendet native Callback-Props über `events`. Seine Komponente exportiert `load`, `fit` und `getHandle`, damit das vollständig typisierte Handle zugänglich ist, ohne das private DOM offenzulegen:

```svelte
<VisioViewer bind:this={viewer} document={diagram}
  events={{ 'document-error': reportError }} />
```

Solid liefert beim Mounten einen Ref-Callback mit dem Handle; beim Aushängen oder beim Ersetzen der Ref ist der Wert `undefined`:

```tsx
<VisioViewer
	document={diagram()}
	viewerRef={setViewer}
	events={{ 'document-error': reportError }}
/>
```

Vanilla besitzt nur das Kindelement, das es selbst erzeugt:

```ts
const viewer = mountViewer(host, { document: diagram, zoom: 1 });
viewer.update({ events: { 'document-error': reportError } });
// Later:
viewer.destroy();
```

Style-Hooks sind in jedem Framework nativ: React `className`/`style`, Vue vererbte Attribute, Angular Host-Attribute, Svelte `class`/`style` und Solid `class`/`style`. Die Anwendung muss dem Host eine nutzbare Höhe geben. Die gemeinsame Gestaltung des Viewers und die Barrierefreiheit liegen im Custom Element.
