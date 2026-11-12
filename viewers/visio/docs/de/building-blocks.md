# Bausteine

`visio-react-viewer` exportiert `useViewerBuildingBlocks`, `DiagramCanvas` sowie die optionalen `Toolbar`, `PageNavigation`, `Inspector` und `ShapesPane`. Diese verwenden den vorhandenen ViewerController, den bestehenden SVG-Renderer, den Auswahl-Inspektor und das Formenfenster wieder. Jedes andere Framework bietet dieselben Bausteine (siehe unten), auch Vanilla JavaScript.

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error, markClean } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
		onOpenSettings: () => openMySettingsDialog(),
	});
	return (
		<>
			<PageNavigation {...toolbarProps} />
			<DiagramCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

`content` akzeptiert dieselbe Byte- oder Blob-Quelle wie der Viewer. `handle` erhält den vorhandenen ViewerController; für die Persistenz im Host verwenden Sie `load`, `applyEdits`, `exportVsdx` und dessen reaktives `state.edit.dirty`. `onDocumentChange`, `onDirtyChange` und `onDocumentError` sind optional. Geben Sie dem Container der Canvas eine Höhe und lassen Sie ihn während des Ladens gemountet.

Das flache `toolbarProps` ist der gesamte Vertrag, den eine eigene Oberfläche benötigt; jedes Feld stammt aus dem vorhandenen Controller, sodass nichts doppelt gespeichert wird:

| Feld                                                                                                                                | Bedeutung                                                                                                                                                                                                                                                                                                                                               |
| ----------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `mode`, `canEdit`                                                                                                                   | `mode` ist nur dann `'edit'`, wenn der Host `canEdit` gewährt **und** die geöffnete Zeichnung eine quellgestützte `.vsdx` ist (eine Legacy-`.vsd` oder eine reine Modellzeichnung meldet `'view'`). `canEdit` ist nur dann wahr, wenn der Host Bearbeitung erlaubt, die Zeichnung quellgestützt ist und gerade kein Laden oder keine Bearbeitung läuft. |
| `pageIndex`, `pages`, `zoom`, `selection`                                                                                           | Aktuelle Seite, Seitenliste, Zoomfaktor und ausgewählte Formen.                                                                                                                                                                                                                                                                                         |
| `canUndo`, `canRedo`, `canDuplicate`, `canGroup`, `canUngroup`, `canDelete`, `canCut`, `canCopy`, `canPaste`                        | Aktivierungsflags. Bearbeitungsaktionen sind bei Schreibschutz `false`; `canCopy` benötigt eine quellgestützte Zeichnung und eine Auswahl auf der aktuellen Seite, aber keine Bearbeitungsberechtigung.                                                                                                                                                 |
| `canZoomIn`, `canZoomOut`                                                                                                           | Zoomgrenzen (10 % bis 800 %).                                                                                                                                                                                                                                                                                                                           |
| `search`                                                                                                                            | `{ query, resultCount, activeIndex, truncated }` aus der Suche des Controllers. `activeIndex` ist -1, bis der Host navigiert.                                                                                                                                                                                                                           |
| `tool`, `tools`, `canDraw`                                                                                                          | Aktives Werkzeug der Canvas (`'pointer'` oder eines von `rectangle`, `ellipse`, `line`, `freeform`, `arc`, `pencil`, `text`, `connector`), die Liste mit lokalisierten Bezeichnungen und Aktivmarkierungen sowie die Angabe, ob Zeichnen erlaubt ist.                                                                                                   |
| `labels`                                                                                                                            | Typisierte, lokalisierte Bezeichnungen (`DiagramToolbarLabels`), ein Eintrag für jedes Steuerelement, in `en`, `de`, `es`, `fr` und `zh-CN`.                                                                                                                                                                                                            |
| `onPageChange`, `onZoomChange`, `onZoomIn`, `onZoomOut`, `onFit`                                                                    | Anzeigeaktionen, immer verfügbar.                                                                                                                                                                                                                                                                                                                       |
| `onUndo`, `onRedo`, `onDuplicate`, `onGroup`, `onUngroup`, `onDelete`, `onCut`, `onCopy`, `onPaste`, `onToolChange`, `onTextChange` | Bearbeitungsaktionen. Sie behalten die Prüfungen des Controllers für Berechtigung, Laden und Quellformat bei. Zwischenablage-Aktionen verwenden die gemountete Canvas und die Browser-Zwischenablage und benötigen daher eine Benutzergeste.                                                                                                            |
| `onSearchChange`, `onNextSearchResult`, `onPreviousSearchResult`                                                                    | Suche, immer verfügbar.                                                                                                                                                                                                                                                                                                                                 |

`Inspector` verwendet `inspectorProps`, und `ShapesPane` verwendet `shapesProps`. Die Canvas selbst mountet keine Titelleiste, keine Multifunktionsleiste, keine Seitensteuerung, keinen Formen- oder Schablonenbereich, keinen Inspektor, keine Datei-Oberfläche, keine Wiederherstellung und keine Einstellungs- oder Freigabedialoge. Die vorhandenen Grenzen für VSDX-Bearbeitung und -Schutz gelten weiterhin.

## Berechtigung und Modus {#permission-and-mode}

`canEdit` folgt den PowerPoint-Bausteinen: Die Zeichnung ist **schreibgeschützt, bis der Host `canEdit: true` übergibt**. Die Berechtigung wird beim Erstellen des Controllers angewendet, bevor die Canvas gemountet wird, sodass nichts kurz bearbeitbar ist (der vollständige Viewer behält seinen bearbeitbaren Standard). Ein Widerruf bricht ausstehende Änderungen ab, entfernt das Zeichenwerkzeug und die Griffe zum Ändern der Größe, zum Drehen und für die Endpunkte. Seitennavigation, Auswahl, Suche, Scrollen und Schwenken, Zoom, Einpassen, Kopieren und Serialisierung bleiben verfügbar. Die Unterstützung der Bearbeitung je Quellformat ist von der Berechtigung des Hosts getrennt; deshalb prüft `mode` auch sie.

## Was die Canvas kann {#what-the-canvas-can-do}

Die eigenständige Canvas verwendet die Gesten des vollständigen Viewers (eine gemeinsame Verdrahtung, keine Kopie): Verschieben, Text direkt im Feld (Doppelklick, F2 oder Eingabe auf einer ausgewählten Form), Griffe zum Ändern der Größe und zum Drehen, Linienendpunkte, die Zeichenwerkzeuge, das Textwerkzeug und das Verbindungswerkzeug, das über `onToolChange` gewählt wird. Das Textwerkzeug zieht ein Feld auf und öffnet einen einfachen Textentwurf (Ctrl+Enter fügt das Textfeld hinzu, Escape bricht ab). Das Verbindungswerkzeug zieht von Form zu Form oder von Verbindungspunkt zu Verbindungspunkt und verankert die Enden genau wie der vollständige Viewer, mit einer rechtwinkligen Führung. Delete, Ctrl+Z, Ctrl+Y und Escape (Zeichenwerkzeug verlassen) funktionieren auf der fokussierten Canvas, und die Ereignisse Kopieren, Ausschneiden und Einfügen des Browsers werden berücksichtigt.
Jede dieser Aktionen ist durch `canEdit` begrenzt, genau wie beim vollständigen Viewer. Abgelehnte Bearbeitungen werden in einer höflichen Live-Region innerhalb der Canvas angekündigt.

## Optionale Bibliotheksoberfläche {#optional-library-ui}

`Toolbar` rendert jedes Steuerelement, das `toolbarProps` unterstützt: Verlauf, Zoom und Einpassen, Duplizieren, Gruppieren, Gruppierung aufheben und Löschen, Ausschneiden, Kopieren und Einfügen, die Zeichenwerkzeuge, die Suche und die unten genannten Schaltflächen für Host-Dialoge, jeweils mit Aktivierungszustand. Sie ist optional: Übergeben Sie `include` (`'file' | 'history' | 'zoom' | 'selection' | 'clipboard' | 'tools' | 'search'`), um nur einige Gruppen anzuzeigen, oder lassen Sie sie weg und verbinden Sie `toolbarProps` mit Ihren eigenen Steuerelementen. `PageNavigation` ist die Seitenauswahl.

`ShapesPane` ist das Formenfenster (Schablonen): Formen suchen, die Dokumentschablone der Zeichnung und die integrierten Schablonen. Ein Klick auf eine Vorlage oder das Ablegen auf der gemounteten Canvas durchläuft denselben Erstellungsablauf wie der vollständige Viewer und unterliegt damit derselben Berechtigung. Mounten Sie die Canvas auf demselben Controller; ohne Canvas werden Vorlagen in der Seitenmitte hinzugefügt.

## Vom Host verwaltete Dialoge {#host-owned-dialogs}

Wie die `onOpen...`-Callbacks in PowerPoint gehören Dialoge dem Host; die Bibliothek mountet im Oberflächenmodus nie einen. Übergeben Sie `onOpenFile`, `onOpenPageSetup`, `onOpenSettings` oder `onOpenShareDialog` an den Hook oder an die Fabrik. Sie kommen über `toolbarProps` mit `canOpenFile`, `canOpenPageSetup`, `canOpenSettings` und `canOpenShareDialog` zurück. Ein Callback, den Sie weglassen, fehlt, und sein Flag ist `false` (nicht aktiviert). Die Bibliotheks-`Toolbar` zeigt nur für die Callbacks, die Sie übergeben haben, eine Schaltfläche.

## Persistenz und Änderungsverfolgung {#persistence-and-dirty-tracking}

`controller.exportVsdx()` liefert die Bytes und ändert `dirty` nie. Nachdem der Host sie gespeichert hat, ruft er `markClean()` auf (am Hook-Ergebnis, am Controller oder am Handle). Das leert `state.edit.dirty`, löst `onDirtyChange(false)` und ein `document-change`-Ereignis der Art `save` aus. Spätere Änderungen und Schritte zum Rückgängigmachen werden ab diesem gespeicherten Stand gezählt.

## Speichern während der Text einer Form bearbeitet wird {#saving-while-editing-text}

Drückt der Benutzer ein vom Host verwaltetes Speichern, während er in einer Form tippt, ist der Entwurf noch nicht im Dokument. Rufen Sie zuerst `await commitTextEdit()` auf (am Hook-Ergebnis, an der Session, an den Vanilla-Blöcken, am Angular-Dienst oder am Controller/Handle) und danach `exportVsdx()`. Der Aufruf endet erst, wenn die Änderung des Entwurfs abgeschlossen ist, auch wenn Blur oder Esc sie bereits gestartet haben, und liefert ein `TextEditOutcome`:

- `committed`: Der Entwurf wurde als eine rückgängig machbare Änderung angewendet; `text` ist der angewendete Text.
- `unchanged` (`reason`: `no-draft` oder `same-text`): Nichts anzuwenden.
- `refused` (`reason`: `read-only` oder `rejected`): Der Editor hat abgelehnt; `message` nennt den Grund, `text` behält den Entwurf zur Wiederherstellung.
- `cancelled` (`reason`: `stale-document`, `superseded` oder `destroyed`): Das Dokument des Entwurfs ist weg oder die Änderung wurde ersetzt; `text` ist der Entwurf.

Jedes Ergebnis enthält `sourceGeneration` und `documentGeneration`, sowie `pageId` und `shapeId`, wenn es einen Entwurf gab. Ein Entwurf gehört zu dem Dokument, in dem er getippt wurde: Wird eine andere Datei geladen oder die Canvas entfernt, verfällt er und wird nie auf ein anderes Dokument angewendet. Hat ein Ladevorgang einen ungespeicherten Entwurf verworfen, gibt der nächste Aufruf von `commitTextEdit()` ihn einmalig als `cancelled` / `stale-document` mit `text`, `pageId`, `shapeId` und der alten `sourceGeneration` zurück, damit ein Speichern, das mit einem Dokumentwechsel zusammenfällt, das Getippte wiederherstellen kann. Validierung und Rückgängig/Wiederholen laufen über den normalen Änderungsweg.

## Pakete {#packages}

Jedes Framework-Paket exportiert `ooxml-ui/visio` erneut, sodass `mountDiagramCanvas(host, controller)`, `mountDiagramInspector(host, controller)`, `mountDiagramShapes(host, controller, { locale? })`, `diagramToolbarProps(controller, fit, locale?, hostCallbacks?)` und `fitDiagramCanvas` aus jedem von ihnen verfügbar sind. Abonnieren Sie über den vorhandenen Controller und rufen Sie die zurückgegebenen Disposer beim Aushängen auf. Es wird kein zweiter Zustandsspeicher eingeführt.

Die Viewer-Pakete und `ooxml-ui` werden gemeinsam veröffentlicht. Jedes Viewer-Paket hängt von der `ooxml-ui`-Version ab, mit der es veröffentlicht wurde, sodass die Kompositions-API immer vorhanden ist. Ein Host, der Versionen von Hand mischt, erhält von einem älteren `ooxml-ui` einen expliziten Fehler, statt dass versehentlich Anwendungsoberfläche gemountet wird.

## Zusammenarbeit {#collaboration}

Die Bausteine nehmen keine Option `collaboration` entgegen: Das Teilen einer Zeichnung in einem Raum ist eine Eigenschaft des vollständigen `<visio-viewer>` (siehe [Zusammenarbeit](/de/collaboration)). Eine eigene Oberfläche kann über `onOpenShareDialog` einen eigenen Freigabeeintrag anbieten.

## Native Framework-Komposition {#native-framework-composition}

Die framework-spezifische Komposition liegt in jedem Framework-Paket. Alle Adapter besitzen einen `ViewerController`, stellen dessen Zustand über die Reaktivität des Frameworks bereit, brechen ersetzte Ladevorgänge ab und beenden Abonnements und Canvas-Handler zusammen mit ihrem Besitzer. Eine unabhängig entfernte Canvas kann am selben Controller erneut gemountet werden, ohne die Quelle neu zu laden.

### Vue {#vue}

```vue
<script setup lang="ts">
import {
	DiagramCanvas,
	Toolbar,
	PageNavigation,
	Inspector,
	useViewerBuildingBlocks,
} from 'visio-vue-viewer';
const props = defineProps<{ content?: Uint8Array; canEdit?: boolean }>();
const blocks = useViewerBuildingBlocks(() => ({ content: props.content, canEdit: props.canEdit }));
const { toolbarProps, loading, error, mode } = blocks;
defineExpose({ controller: blocks.controller });
</script>
<DiagramCanvas v-bind="blocks.canvasProps" />
<Toolbar :toolbar-props="toolbarProps" />
<PageNavigation :toolbar-props="toolbarProps" />
<Inspector v-bind="blocks.inspectorProps" />
```

`toolbarProps`, `state`, `loading`, `error` und `mode` sind Refs; Vue-Templates lösen sie automatisch auf. Übergeben Sie einen Options-Getter oder eine reaktive Ref, um Änderungen an Quelle, Berechtigung und Sprache zu verfolgen.

### Solid {#solid}

```tsx
import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-solid-viewer';
function EmbeddedDiagram(props: { content?: Uint8Array; canEdit?: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		content: props.content,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<DiagramCanvas {...blocks.canvasProps} />
			<Toolbar {...blocks.toolbarProps()} />
		</>
	);
}
```

`toolbarProps`, `state`, `loading`, `error` und `mode` sind Accessoren. Der Controller ist als `blocks.controller` verfügbar; die Entsorgung folgt dem Besitzer in Solid. `PageNavigation` akzeptiert denselben flachen Toolbar-Deskriptor, `Inspector` akzeptiert `inspectorProps`, und `ShapesPane` akzeptiert `shapesProps`.

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
 import { writable } from 'svelte/store';
 import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-svelte-viewer';
 const options = writable({ canEdit: true });
 const blocks = createViewerBuildingBlocks(options);
 const toolbarProps = blocks.toolbarProps;
 // options.set({ canEdit: false }) updates permissions through the store subscription.
</script>
<DiagramCanvas {...blocks.canvasProps} />
<Toolbar toolbarProps={$toolbarProps} />
```

Das Optionsargument akzeptiert einen lesbaren Store oder ein einfaches Anfangsobjekt. Hosts mit einfachen Objekten können `blocks.update(options)` aus einem Rune-Effekt aufrufen. `toolbarProps`, `state`, `loading`, `error` und `mode` sind lesbare Stores. `PageNavigation` nimmt `toolbarProps`, `Inspector` nimmt `inspectorProps`, und `ShapesPane` nimmt `shapesProps`. Der Controller ist als `blocks.controller` verfügbar.

### Vanilla JavaScript {#vanilla-javascript}

```ts
import { createViewerBuildingBlocks } from 'visio-vanilla-viewer';
const blocks = createViewerBuildingBlocks({ canEdit: false, onOpenSettings: openSettings });
blocks.mountCanvas(document.querySelector('#canvas')!);
blocks.mountToolbar(document.querySelector('#toolbar')!, ['history', 'zoom']);
blocks.mountPageNavigation(document.querySelector('select')!);
blocks.mountInspector(document.querySelector('#inspector')!);
blocks.mountShapes(document.querySelector('#shapes')!);
blocks.update({ content: bytes, canEdit: true });
blocks.subscribe(() => render(blocks.toolbarProps, blocks.mode));
// later: blocks.markClean(); blocks.destroy();
```

Die Vanilla-Fabrik führt dieselbe Sitzung aus wie die anderen Adapter. `toolbarProps`, `state`, `loading`, `error` und `mode` sind Getter; `subscribe(listener)` meldet jede Zustands- oder Optionsänderung, und jeder Helfer `mount...` gibt seinen Disposer zurück. `destroy()` entsorgt alles, was noch gemountet ist. Die gemountete Toolbar aktualisiert ihre Elemente an Ort und Stelle, sodass der Fokus Aktualisierungen übersteht.

### Angular {#angular}

```ts
import { Component, Input, inject, type OnChanges } from '@angular/core';
import {
	ViewerBuildingBlocks,
	DiagramCanvasComponent,
	ToolbarComponent,
} from 'visio-angular-viewer';
@Component({
	selector: 'embedded-diagram',
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [DiagramCanvasComponent, ToolbarComponent],
	template: '<visio-diagram-canvas /><visio-diagram-toolbar />',
})
export class EmbeddedDiagram implements OnChanges {
	@Input() content?: Uint8Array;
	@Input() canEdit = false;
	readonly blocks = inject(ViewerBuildingBlocks);
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit });
	}
}
```

Der gescopte Dienst stellt die schreibgeschützten Signale `state`, `toolbarProps`, `loading`, `error` und `mode`, den vorhandenen Controller sowie die Event-Emitter `documentChange`, `dirtyChange` und `documentError` bereit. `DiagramCanvasComponent`, `ToolbarComponent`, `PageNavigationComponent`, `InspectorComponent` und `ShapesPaneComponent` injizieren diesen Besitzer. Platzieren Sie den Provider auf jedem Editor-Host statt anwendungsweit, um mehrere Dokumente voneinander zu isolieren. Alternativ umschließen Sie die Steuerelemente mit `<visio-building-blocks [content]="bytes" [canEdit]="allowed" (dirtyChange)="saveNeeded = $event">` über `ViewerBuildingBlocksComponent`, der Eingaben, Ausgaben und einen gescopten Provider liefert.

Alle optionalen Steuerelemente rendern native Framework-Templates. Hosts können sie weglassen und die flachen Toolbar-Aktionen aus ihrer eigenen Oberfläche aufrufen. Toolbar-Callbacks behalten die Prüfungen des Controllers für Berechtigung, Laden und Quellformat bei. Browser-Regressionstests decken alle sechs nativen Kompositionen ab, einschließlich Bearbeitung, Rückgängig und Wiederholen, Widerruf der Berechtigung, Serialisierung und erneutes Laden, Fehlerwiederherstellung und Abbau.

Für einen Host, der bereits Angular-Signale verwendet, injiziert `createViewerBuildingBlocks(() => ({ content: bytes(), canEdit: allowed() }))` den gescopten Dienst `ViewerBuildingBlocks` und verbindet den Options-Getter mit einem Angular-Effekt. Rufen Sie es im Injektionskontext der Komponente auf. Die Event-Emitter des Dienstes sind auch vor dem ersten expliziten `update` aktiv.

Wenn sich die Prop `mount` einer eigenständigen Canvas ändert, gibt sie den vorherigen Besitzer frei, bevor sie den nächsten an denselben Host anhängt. Wenn sich der Controller eines Inspektors ändert, ersetzt er ebenso sein Abonnement. Das unterstützt den Wechsel zwischen unabhängigen Editor-Sitzungen, ohne verwaiste Eingabe-Handler oder Zustandsabonnements.

## Gemeinsames Office-Design {#shared-office-theme}

Jede native optionale Toolbar, Seitenauswahl und jeder Inspektor folgt denselben CSS-Variablen `--office-*` wie die Zeichen-Canvas. Die Bausteine installieren ihr gemeinsames Stylesheet automatisch; es gibt kein framework-spezifisches Design-Objekt und keinen zweiten Dokumentzustandsspeicher. Setzen Sie das Design auf Ihrem eigenen Host-Element:

```html
<section
	data-office-theme="dark"
	style="--office-accent: #a78bfa; --office-font: Inter, sans-serif"
>
	<!-- DiagramCanvas, Toolbar, PageNavigation and Inspector -->
</section>
```

Verwenden Sie `data-office-theme="light"`, `"dark"` oder `"auto"` am Dokumentstamm oder an einem umschließenden Abschnitt. `auto` folgt der Systemeinstellung, einschließlich Laufzeitänderungen. Gescopte helle und dunkle Hosts können auf einer Seite nebeneinander bestehen. Ohne gescoptes Attribut erben die Bausteine das Root-Design und die Host-Variablen über die Shadow-Grenze der Canvas hinweg. Ein Wechsel des Designs lädt die Canvas nicht neu, setzt die Auswahl nicht zurück und ändert keine Dokumentfarben oder serialisierten Bytes.

Für eine host-spezifische Palette überschreiben Sie `--office-background`, `--office-surface`, `--office-foreground`, `--office-muted-foreground`, `--office-border`, `--office-accent`, `--office-ring` und `--office-font` auf dem umschließenden Abschnitt. Die älteren Visio-Überschreibungen `--vv-surface`, `--vv-ink`, `--vv-border`, `--vv-accent`, `--vv-focus` und `--vv-font` bleiben unterstützt und haben Vorrang vor ihren kanonischen Gegenstücken. Die Canvas behält zusätzlich `--vv-background`, `--vv-secondary` und die vorhandenen Überschreibungen der Statusleiste. Die Diagrammseite behält ihr eigenes Dokumentdesign, während die umgebenden Steuerelemente dem UI-Design folgen.
