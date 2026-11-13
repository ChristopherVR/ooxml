# Bausteine

Jedes Framework-Paket (`@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer`, `xlsx-angular-viewer`,
`xlsx-svelte-viewer`, `xlsx-solid-viewer` und `xlsx-vanilla-viewer`) exportiert einen Hook oder
eine Factory für Bausteine (`useViewerBuildingBlocks`, `createViewerBuildingBlocks`), eine
`SpreadsheetCanvas` sowie die optionalen Steuerelemente `Toolbar`, `FormulaBar` und `SheetTabs`.
Alle nutzen die vorhandene Sitzung der Arbeitsmappe, das Raster, die Formeln, die Validierung, den
Schutz und den Verlauf; es gibt keinen zweiten Zustandsspeicher. Die Form folgt dem Hook
`useViewerBuildingBlocks` des PowerPoint-Viewers.

```tsx
import {
	SpreadsheetCanvas,
	FormulaBar,
	SheetTabs,
	useViewerBuildingBlocks,
} from '@christophervr/xlsx-react-viewer';

function EmbeddedSheet({ workbook, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		workbook,
		canEdit,
		handle,
	});
	return (
		<>
			{toolbarProps && <FormulaBar {...toolbarProps} />}
			<SpreadsheetCanvas {...canvasProps} />
			{toolbarProps && <SheetTabs {...toolbarProps} />}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Der Hook liefert außerdem `mode` (`'view'` oder `'edit'`) sowie die reaktiven Werte `dirty`, `selection`
(als Text in A1-Notation), `collaboration` (wer im Raum ist und der Verbindungszustand) und `editError`
(siehe [Validierungsfehler](#validation-failures)). Rendern Sie die Arbeitsfläche während des Ladens;
`toolbarProps` ist bis zum Einbinden null. Geben Sie ihrem Container eine Höhe. Die Arbeitsfläche
bindet nur das Raster ein. Datei, Einstellungen, Freigabe, Titelleiste, Menüband, Statusleiste,
Bearbeitungsleiste und Blattregister fehlen, sofern der Host keine eigenen Steuerelemente bereitstellt.

`toolbarProps` enthält die Auswahl bzw. die aktive Zelle, den Zustand des Blatts, den Formeltext, die
Formatierung und die Zustände `enabled`. Verbinden Sie `onBold`, `onItalic`, `onUnderline`,
`onFontFamilyChange`, `onFontSizeChange`, `onNumberFormatChange`, `onAlignmentChange`, `onUndo` und
`onRedo` mit eigenen Steuerelementen. Die Formatfelder sind bei gemischten Werten null. Die Prüfung
der Formatierung ist auf 4096 Zellen begrenzt: Größere Auswahlen liefern null und
`formattingComplete=false`.

`enabled` hat je Callback ein Kennzeichen: `bold`, `italic`, `underline`, `fontFamily`, `fontSize`,
`alignment`, `numberFormat`, `undo`, `redo` sowie die Arbeitsabläufe des Hosts `openShareDialog`,
`settings`, `formatCells`, `findReplace` und `insertFunction`. `canUndo` und `canRedo` wiederholen die
Kennzeichen für Rückgängig und Wiederholen. `options` listet die Auswahlmöglichkeiten, die ein eigenes
Steuerelement für Schriftart, Schriftgrad, Ausrichtung oder Zahlenformat anbieten kann
(`fontFamilies`, `fontSizes`, `alignments`, `numberFormats`), und `labels` enthält die lokalisierten
Namen der Steuerelemente.

`onSelect`, `onSheetChange`, `onFormulaChange` und `commitEdit` liefern false zurück, wenn die
Zellvalidierung die Aktion blockiert. Callbacks für die Formatierung liefern ein `Promise<boolean>`.
Sie übernehmen eine laufende Zellbearbeitung, bevor sie handeln, und behalten die Auswahl, wenn eigene
Steuerelemente den Fokus erhalten. Der Formeltext wird vom optionalen FormulaBar mit Enter übernommen;
eine abgelehnte Eingabe setzt `aria-invalid`. Das vorhandene Handle behält Speichern, Serialisierung und
die Verfolgung von Änderungen bei. Ein Speichern wird abgelehnt, wenn die Eingabe nicht übernommen werden
kann. Berechtigungen und der Schutz des Arbeitsblatts bleiben maßgeblich.

Für jedes Framework stellt `mountEditorSurface(host, options)` dieselbe Oberfläche bereit, das
vorhandene Handle, `getToolbarProps()`, `subscribe(listener)`, `update` und `destroy`. Die Hosts
verantworten Datei- und optionale Dialogabläufe; eine verborgene Werkzeugleiste der Bibliothek ist
nicht erforderlich.

`formula`, `formulaCaret`, `editing` und `editMode` lesen die aktive Bearbeitungsbrücke des Rasters.
`onFormulaInput(text, caret?)` aktualisiert denselben Entwurf; `onFormulaChange(text)` übernimmt ihn.
Ein abgelehnter Entwurf lässt sich über `onFormulaInput` korrigieren, ohne den ungültigen Wert zuvor zu
übernehmen. `onCancelEdit()` bricht ihn ab. Es gibt keinen eigenen Zustandsspeicher für Formeln.

Die Composition-API setzt die Laufzeit `ooxml-ui` desselben Workspace-Releases voraus; die
Framework-Pakete hängen davon ab und werden zusammen mit ihr veröffentlicht. Bestehende Importe des
vollständigen Viewers bleiben mit der bisher veröffentlichten Laufzeit kompatibel. Eine ältere Laufzeit
lehnt eine Composition-Anfrage ausdrücklich ab, statt versehentlich Anwendungs-UI einzubinden.

## Callbacks des Hosts für Dialoge {#host-callbacks-for-dialogs}

Die Anwendung öffnet Dialoge für Freigabe, Einstellungen, Zellen formatieren, Suchen und Ersetzen sowie
Funktion einfügen. Eine Oberfläche bindet sie nie ein. Übergeben Sie stattdessen einen Callback, wie es
die Bausteine von PowerPoint tun (`onOpenSettings`, `onOpenShareDialog`), und öffnen Sie Ihre eigene
Oberfläche:

```ts
useViewerBuildingBlocks({
	workbook,
	onOpenShareDialog: ({ activeCell }) => openMyShareDialog(activeCell),
	onOpenFormatCells: ({ selection }) => openMyFormatPanel(selection),
});
```

Die verfügbaren Callbacks sind `onOpenShareDialog`, `onOpenSettings`, `onOpenFormatCells`,
`onOpenFindReplace` und `onOpenInsertFunction`. Jeder erhält `{ sheet, sheetName, selection,
activeCell }`. `toolbarProps` stellt die entsprechenden `onOpenShareDialog`, `onOpenSettings`,
`onOpenFormatCells`, `onOpenFindReplace` und `onOpenInsertFunction` bereit (sie übernehmen zuerst eine
laufende Bearbeitung und liefern ein `Promise<boolean>`). `enabled.openShareDialog` und die übrigen
Kennzeichen sind nur dann true, wenn der Callback übergeben wurde (die Befehle zum Formatieren und
Einfügen benötigen außerdem die Berechtigung zum Bearbeiten). `handle.share()` ruft `onOpenShareDialog`
auf; ohne Callback tut es nichts, und `enabled.openShareDialog` ist false. Ein Arbeitsablauf ohne
Callback ist nicht erreichbar und wird nie stillschweigend eingebunden. Hosts, die die mitgelieferten
Dialoge wünschen, rendern stattdessen den vollständigen `SpreadsheetEditor`.

## Kontextmenüs und Workflow-Operationen {#context-menus-and-workflow-operations}

Die Typen, Signaturen und Ablehnungsbedingungen stehen in der [Element-API-Referenz](/de/api#host-menus-and-workflows).

`onOpenFormatCells`, `onOpenFindReplace` und `onOpenInsertFunction` erhalten zusätzlich zum obigen
Kontext die Operationen, um den Ablauf abzuschließen. Dieselben Operationen stehen in
`toolbarProps.workflows`, ein dauerhaft geöffnetes Panel braucht also keinen Callback. Jede
Operation prüft beim Ausführen Schreibschutz, Blattschutz und die laufende Zellbearbeitung neu,
läuft über die Bearbeitungssitzung (Rückgängig, Wiederholen, Dirty-Status und `onWorkbookChange`
verhalten sich wie in der mitgelieferten Oberfläche) und meldet eine Ablehnung über den Rückgabewert.

- `workflows.formatCells`: `getState()` liefert `{ font, alignment, border, protection }` (jeweils
  `{ value, mixed }`: gemeinsame Eigenschaften und die Namen abweichender), `numFmt` und `fill`
  (null bei gemischten Werten) sowie `complete` (false bei zu großer Auswahl). `canApply()` und
  `apply(patch)` wenden einen Stil-Patch als einen Rückgängig-Schritt auf die Auswahl an.
- `workflows.find`: `findAll(query)`, `findNext(query, { backwards })`, `select(match)`,
  `replace(query, text)`, `replaceAll(query, text)` und `canReplace()`. Suchen funktioniert auch in
  einer schreibgeschützten Arbeitsmappe; Ersetzen wird dort abgelehnt, durch Blattschutz gesperrte
  Zellen werden übersprungen. Wie das Ersetzen in Excel selbst wenden `replace` und `replaceAll` standardmäßig keine Datenüberprüfung an: Ein Ersatzwert wird auch geschrieben, wenn die Regel der Zelle ihn bei der Eingabe ablehnen würde. Zum Prüfen übergeben Sie `{ validate: true, onRefused }` als dritten Parameter oder setzen die Editor-Option `validateReplace` (Prop `validateReplace` jeder Bindung und Eigenschaft von `<xlsx-editor>`), die auch der mitgelieferte Dialog Suchen und Ersetzen beachtet. Ein von der Regel abgelehnter Wert lässt seine Zelle unverändert und wird in `onRefused(cells)` aufgeführt (`{ sheet, row, col, ref, value, message, title? }`); der Dialog zeigt einen Hinweis mit der Anzahl. Mit `{ validate: false }` nimmt ein Aufruf die Editor-Option zurück.
- `workflows.insertFunction`: `categories()`, `search({ text, category })` (dokumentierte `name`,
  `category`, `syntax` und eine lokalisierte `description`), `canInsert()` und `insert(name, args?)`.
  Es fügt `NAME(` an der Einfügemarke des aktiven Formelentwurfs ein (und startet bei Bedarf einen)
  und liefert `{ text, caret }`, damit Ihr Eingabefeld folgen kann. Ein Klartext-Entwurf wird abgelehnt.

Für ein eigenes Kontextmenü übergeben Sie `onContextMenu(request)`. Es läuft bei einem Rechtsklick
auf eine Zelle, eine Zeilen- oder Spaltenüberschrift und (bei eingebundener Registerleiste) ein
Blattregister. Geben Sie `true` zurück, um das native Menü zu unterdrücken, und zeigen Sie
`request.items` (`id`, `label`, `command`, `enabled`, `checked`, `shortcut`, `separatorBefore`) bei
`request.x`, `request.y` an; eine Auswahl führen Sie mit `request.execute(id)` aus. Jeder andere
Rückgabewert zeigt das native Menü. Ohne Rechtsklick liefert
`toolbarProps.contextMenu('cell' | 'row' | 'column' | 'tab', sheet?)` dasselbe Modell, so erhält
ein Host mit eigenen Blattregistern das Registermenü. Eine Oberfläche registriert nur wenige
Befehle, daher stehen die meisten Einträge mit `enabled: false` in der Liste; Format Cells ist
aktiv, wenn `onOpenFormatCells` gesetzt und die Arbeitsmappe bearbeitbar ist. `enabled.contextMenu`,
`enabled.find`, `enabled.replace`, `enabled.applyFormatCells` und `enabled.insertFunctionCall`
zeigen, was jetzt funktioniert. Um das native Menü zu behalten und zu ändern, übergeben Sie stattdessen
`onContextMenuItems(request)` (läuft nur, wenn `onContextMenu` nicht `true` zurückgab) und geben
`{ hide, add }` zurück. `hide` listet native Eintrags-IDs aus `request.items`; `add` listet
Host-Einträge `{ id, label, enabled, requiresEdit, checked, separatorBefore, position,
onSelect(request) }`, die oberhalb (`position: 'top'`, Standard) oder unterhalb (`'bottom'`) der
nativen Einträge stehen. `enabled` darf eine Funktion sein. Native Einträge behalten ihren
deaktivierten Zustand und werden bei der Ausführung erneut geprüft; ein Host-Eintrag ebenfalls
(sein `enabled` und `requiresEdit` gegen eine schreibgeschützte Arbeitsmappe), sodass ein offen
gelassenes Menü keine veraltete Auswahl ausführen kann. Es ergänzt `hiddenActions` und
`contextMenu.extraCommands` in `customization`, die für die mitgelieferte Anwendung gelten.

## Validierungsfehler {#validation-failures}

Eine laufende Zellbearbeitung muss vor dem Formatieren, Navigieren oder Speichern übernommen werden. Lehnt
die Datenüberprüfung sie ab, liefern die Callbacks `false`, die Bearbeitung bleibt offen, und der Fehler
wird auf drei Wegen gemeldet:

- `onValidationError({ message, title?, ref, style? })` in den Optionen (das Element sendet ein
  Ereignis `validation-error` mit denselben Details);
- die reaktive Zeichenkette `editError` (auch `toolbarProps.editError`), die wieder null wird, sobald
  der Entwurf geändert, korrigiert oder abgebrochen wird;
- das optionale `FormulaBar` setzt `aria-invalid` und zeigt die Meldung in einem `role="alert"`-Element
  an.

### Der Warnhinweis der Datenüberprüfung {#the-data-validation-alert}

Die Bausteine zeichnen nie die Schaltflächen Wiederholen / Abbrechen des Hinweises aus dem Raster. Die drei
Hinweisarten von Excel zeigt der Host über `onOpenValidationAlert(request)` an (akzeptiert auch in
`surfaceActions` und von jedem Framework-Adapter). Die Anfrage enthält `{ ref, sheet, sheetName, value,
style, title, message }`, wobei `value` der Text ist, den der Nutzer eingegeben hat, und die Methoden,
die zu den Schaltflächen von Excel passen:

| `style`       | Schaltflächen in Excel  | Methoden                                      |
| ------------- | ----------------------- | --------------------------------------------- |
| `stop`        | Wiederholen / Abbrechen | `retry()`, `cancel()`                         |
| `warning`     | Ja / Nein / Abbrechen   | `accept()` (Ja), `retry()` (Nein), `cancel()` |
| `information` | OK / Abbrechen          | `accept()` (OK), `cancel()`                   |

`retry()` hält den Editor mit dem Entwurf offen, `cancel()` stellt den alten Wert wieder her, und
`accept()` übernimmt den Wert trotzdem, genau wie die mitgelieferten Schaltflächen. Beantworten Sie eine
Anfrage nur einmal; spätere Aufrufe werden ignoriert. Bis sie beantwortet ist, ignoriert der
Zelleneditor weitere Übernahmen, beantworten Sie sie also immer (ein Callback, der eine Ausnahme wirft,
zählt als `retry()`). `onValidationError` und `editError` melden den Fehler weiterhin.

Ohne Callback wird nichts eingebunden: Jede Art verhält sich wie Wiederholen, die Bearbeitung bleibt offen,
und der Fehler wird über `onValidationError` / `editError` gemeldet.
`toolbarProps.enabled.validationAlert` ist nur dann true, wenn der Host den Callback bereitstellt. Der
vollständige `SpreadsheetEditor` und `<xlsx-editor>` behalten ihren eingebauten Hinweis.

Eigene Steuerelemente sollten außerdem `editError` anzeigen und dem Nutzer erlauben, den Entwurf über
`onFormulaInput` zu korrigieren oder über `onCancelEdit` zu verwerfen.

## Optionale Werkzeugleiste {#optional-toolbar}

Die Bibliothek-`Toolbar` zeigt jedes Steuerelement an, das die Props unterstützen, sodass sie eine selbst
gebaute ersetzen kann: Rückgängig, Wiederholen, Schriftart, Schriftgrad, Fett, Kursiv, Unterstrichen,
linksbündig / zentriert / rechtsbündig und Zahlenformat. Schaltflächen stellen `aria-pressed` bereit
(`mixed` bei gemischter Auswahl), Auswahllisten zeigen bei gemischter Formatierung einen leeren Wert, und
alle Steuerelemente beachten `enabled`. Schaltflächen erhalten keinen Fokus, daher behält das Raster
seine Auswahl beim Formatieren. Die Werkzeugleiste zeigt die oben genannten Dialogabläufe nicht an;
verbinden Sie `onOpenShareDialog` und die übrigen Callbacks `onOpen*` mit Ihren eigenen Steuerelementen.
`toolbarModel(toolbarProps)` (exportiert von `xlsx-vanilla-viewer` und der Svelte-Laufzeit) liefert
dieselbe Liste der Steuerelemente für eigene Renderer.

## Zusammenarbeit {#collaboration}

Gemeinsames Bearbeiten in Echtzeit funktioniert auf einer Oberfläche: Übergeben Sie `collaboration`
(eine Raum-ID, eine Server-URL, einen Anbieter oder eine Sitzung) in den Optionen, dann liefert der Hook
den reaktiven Zustand `collaboration` (`active`, `roomId`, `status`, `synced`, `people`). `onUndo` und
`onRedo` machen während einer Freigabe nur die Änderungen dieses Fensters rückgängig. Verlassen oder
Entfernen beendet den Raum. Die Anwesenheitsliste der Titelleiste gehört zur Anwendung, daher rendert ein
Host `collaboration.people` selbst.

## Vanilla {#vanilla}

`xlsx-vanilla-viewer` exportiert `createViewerBuildingBlocks(options)`, das auf derselben Sitzung wie die
anderen Frameworks aufbaut:

```ts
import { createViewerBuildingBlocks } from 'xlsx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ workbook, canEdit: true });
const stops = [
	blocks.mountToolbar(toolbarHost),
	blocks.mountFormulaBar(formulaHost),
	blocks.mountSheetTabs(tabsHost),
	blocks.mountCanvas(canvasHost),
];
blocks.subscribe((state) => render(state.dirty, state.editError, state.toolbarProps));
blocks.update({ workbook, canEdit: false });
blocks.getState().toolbarProps?.onBold();
blocks.destroy();
```

Die Oberfläche ist **schreibgeschützt, bis der Host `canEdit: true` übergibt**, wie bei den Bausteinen
von PowerPoint und Visio (der vollständige `SpreadsheetEditor` bleibt standardmäßig bearbeitbar). Die
Berechtigung wird angewandt, bevor die Arbeitsfläche eingebunden wird, sodass `mode` mit `'view'` beginnt
und kein bearbeitbares Aufblitzen entsteht. `readOnly` setzt sich weiterhin gegen `canEdit: true` durch.

Jedes `mount*` liefert seine Entsorgungsfunktion zurück. `getState()` liefert denselben Schnappschuss wie
die Hooks (`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration`,
`editError`), und `handle` ist das übliche imperative Handle. Die Steuerelemente aktualisieren ihre
Elemente an Ort und Stelle, daher bleiben Fokus und Entwürfe bei Aktualisierungen erhalten.

## Native Komposition in Frameworks {#native-framework-composition}

Der Framework-Code wird aus dem Paket jedes Frameworks ausgeliefert. Alle Adapter besitzen eine vorhandene
Raster-Engine und stellen dasselbe `EditorHandle` bereit; ihre nativen reaktiven Schnappschüsse liefern
`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration` und `editError`.
Optionale native Steuerelemente erhalten eine Prop bzw. Eingabe `toolbarProps`. Eigene Steuerelemente
rufen die Aktionen desselben Schnappschusses auf.

### Vue {#vue}

Rufen Sie das Composable während `setup` auf. Übergeben Sie eine Getter-Funktion, um Änderungen an
Berechtigung, Dokument, Bytes und Callbacks zu verfolgen. Berechnete Refs werden in Templates
entpackt; das Entfernen der Komponente entsorgt die Engine.

```vue
<script setup lang="ts">
import {
	useViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-vue-viewer';
const props = defineProps<{ workbook: import('xlsx-vue-viewer').Workbook; canEdit: boolean }>();
const { canvasProps, toolbarProps, loading, error, handle } = useViewerBuildingBlocks(() => ({
	workbook: props.workbook,
	canEdit: props.canEdit,
}));
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<FormulaBar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<SpreadsheetCanvas v-bind="canvasProps" style="height: 600px" />
	<SheetTabs v-if="toolbarProps" :toolbar-props="toolbarProps" />
</template>
```

### Angular {#angular}

Stellen Sie `ViewerBuildingBlocks` pro Host bereit und injizieren Sie diese Instanz dann in den Host oder
dessen eigene Steuerelemente. `options` ist ein beschreibbares Signal; die Schnappschüsse sind berechnete
Signale. Der `DestroyRef` des Injectors übernimmt die Aufräumarbeiten. Der Dienst stellt außerdem die
Ereignis-Emitter `workbookChange`, `workbookError`, `selectionChange` und `dirtyChange` für Abonnements
des Hosts bereit, neben den Callbacks in den Optionen. Verwenden Sie `takeUntilDestroyed()`, wenn Sie in
Hosts abonnieren.

```ts
@Component({
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [SpreadsheetCanvasComponent, ToolbarComponent, FormulaBarComponent, SheetTabsComponent],
	template: `
		@if (blocks.toolbarProps(); as p) {
			<spreadsheet-toolbar [toolbarProps]="p" />
			<spreadsheet-formula-bar [toolbarProps]="p" />
			<button (click)="p.onBold()" [disabled]="!p.canEdit">Custom bold</button>
			<spreadsheet-sheet-tabs [toolbarProps]="p" />
		}
		<spreadsheet-canvas [mount]="blocks.canvasProps.mount" style="display:block;height:600px" />
	`,
})
export class WorkbookHost {
	readonly workbook = input.required<Workbook>();
	readonly canEdit = input(true);
	readonly blocks = inject(ViewerBuildingBlocks);
	constructor() {
		effect(() => this.blocks.options.set({ workbook: this.workbook(), canEdit: this.canEdit() }));
	}
}
```

Importieren Sie die Komponenten, den Dienst und `Workbook` aus `xlsx-angular-viewer` sowie `Component`,
`input`, `effect` und `inject` aus `@angular/core`. `SpreadsheetBuildingBlocks` ist der produktspezifische
Name für denselben Dienst. Sein `handle` behält Laden, Speichern und Auswahl bei.

### Solid {#solid}

Erzeugen Sie die Bausteine innerhalb eines Solid-Besitzers. Eingabe-Getter und Ausgabe-Accessoren nehmen am
reaktiven Graphen von Solid teil; die Aufräumarbeiten folgen diesem Besitzer. Komponenten akzeptieren
reaktive Getter als Props.

```tsx
import { Show } from 'solid-js';
import {
	createViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-solid-viewer';
function WorkbookHost(props: { workbook: Workbook; canEdit: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		workbook: props.workbook,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<Show when={blocks.toolbarProps()}>
				{(p) => (
					<>
						<Toolbar toolbarProps={p()} />
						<FormulaBar toolbarProps={p()} />
						<button onClick={() => p().onBold()}>Custom bold</button>
						<SheetTabs toolbarProps={p()} />
					</>
				)}
			</Show>
			<SpreadsheetCanvas {...blocks.canvasProps} />
		</>
	);
}
```

### Svelte 5 {#svelte-5}

Erzeugen Sie den Store während der Initialisierung der Komponente. `$blocks` abonniert nativ; verwenden Sie
ein `$effect`, um reaktive Optionen weiterzuleiten. Die Aufräumarbeiten der Engine folgen dem Entfernen der
Komponente. Die Quellen der Komponenten und ihre Deklarationen sind als eigene Unterpfade veröffentlicht.

```svelte
<script lang="ts">
  import { createViewerBuildingBlocks, type Workbook } from 'xlsx-svelte-viewer/runtime';
  import SpreadsheetCanvas from 'xlsx-svelte-viewer/SpreadsheetCanvas.svelte';
  import Toolbar from 'xlsx-svelte-viewer/Toolbar.svelte';
  import FormulaBar from 'xlsx-svelte-viewer/FormulaBar.svelte';
  import SheetTabs from 'xlsx-svelte-viewer/SheetTabs.svelte';
  let { workbook, canEdit = false }: {workbook: Workbook; canEdit?: boolean} = $props();
  const blocks = createViewerBuildingBlocks();
  $effect(() => blocks.update({workbook, canEdit}));
</script>
{#if $blocks.toolbarProps}
  <Toolbar toolbarProps={$blocks.toolbarProps} />
  <FormulaBar toolbarProps={$blocks.toolbarProps} />
  <button onclick={() => $blocks.toolbarProps?.onBold()}>Custom bold</button>
  <SheetTabs toolbarProps={$blocks.toolbarProps} />
{/if}
<SpreadsheetCanvas {...blocks.canvasProps} />
```

`FormulaBar` teilt sich den live Entwurf des Rasters. Enter übernimmt, und eine abgelehnte Eingabe setzt
`aria-invalid`. Wiederholen erlaubt eine Korrektur; Escape bricht ab. Widersprüchliche Formatierung,
Auswahl, Blattnavigation und Speichern bleiben blockiert, solange die Validierung einen Entwurf ablehnt.
Schreibgeschützt blockiert jede Änderung, während Auswahl, Blattnavigation und Serialisierung erhalten
bleiben.

## Gemeinsames Erscheinungsbild {#shared-appearance}

Platzieren Sie die Arbeitsfläche und die optionalen nativen Steuerelemente unter einem gemeinsamen Vorfahren
mit `data-office-theme="light|dark|auto"`. Alle fünf nativen Framework-Integrationen installieren dieselben
gemeinsamen Stile für Steuerelemente und lesen die vererbten Token `--office-*`. Passen Sie
`--office-surface`, `--office-background`, `--office-foreground`, `--office-border`, `--office-accent`,
`--office-ring`, `--office-font` und `--office-radius` in diesem Geltungsbereich an. Änderungen zur
Laufzeit erhalten Entwürfe und Fokus.

Die vorhandenen Optionen `theme` und `themeColors` der Arbeitsfläche gelten weiterhin lokal; sie
schreiben die Vorfahren des Hosts oder nebengeordnete Steuerelemente nicht um. Überschreibungen der älteren
`--xve-*` auf einem gemeinsamen Vorfahren erreichen weiterhin sowohl die Arbeitsfläche als auch die optionalen
Steuerelemente. Beispiele zum Geltungsbereich und zur Kompatibilität finden Sie unter [Themes](/de/theming).
