# Element-API

`<xlsx-editor>` (Klasse `XlsxEditorElement`) ist der eine Editor, den jedes Paket einbindet. Registrieren Sie ihn mit `defineXlsxEditor()` (von jedem Paket exportiert; die Registrierung ist idempotent und tut ohne `customElements` nichts), oder überlassen Sie das `mountEditor` und den Framework-Komponenten.

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const editor = document.createElement('xlsx-editor');
editor.fileName = 'Budget.xlsx';
document.body.append(editor);
await editor.load(bytes, 'Budget.xlsx');
```

## Attribute und Eigenschaften

| Attribut           | Eigenschaft      | Typ                                                         | Standard     |
| ------------------ | ---------------- | ----------------------------------------------------------- | ------------ |
| `locale`           | `locale`         | `en \| fr \| de \| es \| zh-CN` (jedes Tag wird zugeordnet) | `en`         |
| `read-only`        | `readOnly`       | boolean                                                     | `false`      |
| `file-name`        | `fileName`       | string                                                      | `Book1.xlsx` |
| `author-name`      | `authorName`     | string                                                      | `Author`     |
| `theme`            | `theme`          | `light \| dark \| auto`                                     | `auto`       |
| `show-toolbar`     | `showToolbar`    | boolean (`"false"` blendet aus)                             | `true`       |
| `show-formula-bar` | `showFormulaBar` | boolean (`"false"` blendet aus)                             | `true`       |

Nur als Eigenschaften:

| Eigenschaft          | Typ                                  | Hinweise                                                                                                                                                               |
| -------------------- | ------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `themeColors`        | `Partial<Record<...>>`               | Token-Überschreibungen, angewendet als benutzerdefinierte Eigenschaften `--xve-*` ([Themes](/theming)).                                                                |
| `hiddenActions`      | `string[]`                           | Menüband-Steuerelemente, Titelleistenelemente und Kontextmenüeinträge, die über ihre stabile ID ausgeblendet werden.                                                   |
| `ribbonAddIns`       | `RibbonAddInTab[]`                   | Registerkarten, die der Host nach den eingebauten ergänzt ([Bindings](/bindings#ribbon-add-in-tabs)).                                                                  |
| `customization`      | `ExcelCustomization \| null`         | Capabilities, Host-Seiten im Datei-Register und Host-Menübefehle (siehe [Anpassung](#customization)).                                                                  |
| `workbook`           | `Workbook \| null`                   | Liefert die geöffnete Arbeitsmappe oder ersetzt sie durch eine gesetzte.                                                                                               |
| `dirty`              | `boolean` (nur lesen)                | Ungespeicherte Änderungen.                                                                                                                                             |
| `activeSheet`        | `number`                             | Nullbasierter Index des sichtbaren Blatts.                                                                                                                             |
| `collaboration`      | `XlsxCollaborationOptions \| null`   | Gibt die geöffnete Arbeitsmappe in einem Raum frei, solange sie gesetzt ist ([Zusammenarbeit](/de/collaboration)).                                                     |
| `collaborationState` | `XlsxCollaborationState` (nur lesen) | Ob freigegeben wird, der Verbindungsstatus und wer sich im Raum befindet.                                                                                              |
| `googleFonts`        | `false \| { baseUrl }`               | Webfont-Ersatz für Schriften, die der Leser weder installiert noch eingebettet hat. Ohne Angabe gilt die Einstellung des Trust Centers des Lesers (standardmäßig aus). |
| `surfaceOnly`        | boolean                              | Vor dem Einhängen setzen, um die Anwendungsoberfläche wegzulassen; zusammen mit `surfaceActions` verwenden.                                                            |
| `surfaceActions`     | `SurfaceHostActions`                 | Host-Callbacks, die Dialoge ersetzen, wenn `surfaceOnly` gesetzt ist. Ein Ablauf ohne Callback wird in `getToolbarProps().enabled` als nicht aktiviert gemeldet.       |
| `loading`            | `boolean` (nur lesen)                | Ob eine Arbeitsmappe gerade geöffnet wird.                                                                                                                             |

## Anpassung {#customization}

`customization` nimmt ein `ExcelCustomization`-Objekt entgegen, den gemeinsamen Vertrag von `ooxml-core/embedding`, den auch die anderen Editoren verwenden. Jedes Binding reicht es als Prop bzw. Input `customization` weiter. Es wird angewendet, wenn sich die Objektidentität ändert; halten Sie es daher stabil. `null` stellt die Standardwerte wieder her.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send range', onSelect: ({ ref }) => send(ref) }],
	},
};
```

- `capabilities`: `false` entfernt jeden Weg zur Funktion (`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`): Menüband, Seiten im Datei-Register, Menüs und Tastenkürzel. `hiddenActions` bleibt unverändert; beide werden vereinigt.
- `backstage.extraPages`: Host-Seiten `{ id, label, group?, order?, render }` im Datei-Register, nach den eingebauten Seiten ihrer Gruppe. `render` ist eine Mount-Funktion, ein Custom-Element-Tag oder `{ element, properties }` und erhält einen `ExcelRegionContext`.
- `contextMenu.extraCommands`: Host-Befehle `{ id, label, icon?, group?, when?, disabled?, onSelect }` oberhalb (`top`, Standard) oder unterhalb (`bottom`) der eingebauten Einträge. `onSelect` läuft nach dem Schließen des Menüs mit einem `ExcelMenuContext`. Beschriftungen werden unverändert angezeigt (lokalisieren Sie sie selbst).

`edit: false` ist eine eigene Schreibsperre neben dem schreibgeschützten Schalter des Hosts und einer Zusammenarbeitsrolle `viewer`: Keine hebt eine andere auf. Capabilities sind eine UI-Richtlinie, kein Dokumentschutz. Das Ersetzen von Bereichen (Optionsdialog, Bereiche) wird für Excel noch nicht angeboten. Alle Produkte im Überblick: [Einbettungsleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Methoden

| Methode                          | Rückgabe                  | Hinweise                                                                                                                                                                                                                                                                 |
| -------------------------------- | ------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `load(bytes, fileName?)`         | `Promise<void>`           | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv`; das Format wird an den Bytes erkannt. Löst bei einem Fehler `workbook-error` aus und wird abgelehnt.                                                                                                                          |
| `newWorkbook()`                  | `void`                    | Eine leere Arbeitsmappe namens `Book1.xlsx`.                                                                                                                                                                                                                             |
| `save()`                         | `Promise<Blob>`           | `.xlsx`. Setzt `dirty` nicht zurück.                                                                                                                                                                                                                                     |
| `saveBytes(format?)`             | `Promise<Uint8Array>`     | `'xlsx'` (Standard) oder `'csv'` (das aktive Blatt, formatierte Werte).                                                                                                                                                                                                  |
| `download(name?)`                | `Promise<void>`           | Speichert, lädt herunter und ruft dann `markClean()` auf.                                                                                                                                                                                                                |
| `markClean()`                    | `void`                    |                                                                                                                                                                                                                                                                          |
| `select(ref)`                    | `void`                    | Eine A1-Zelle oder ein A1-Bereich auf dem aktiven Blatt (`'C3'`, `'A1:D20'`).                                                                                                                                                                                            |
| `getSelection()`                 | `string`                  | Die aktuelle Auswahl als A1-Bezug.                                                                                                                                                                                                                                       |
| `setActiveSheet(index)`          | `void`                    |                                                                                                                                                                                                                                                                          |
| `undo()`, `redo()`               | `void`                    | Bei Freigabe werden nur die Änderungen dieses Fensters rückgängig gemacht.                                                                                                                                                                                               |
| `startCollaboration(options)`    | `Promise<void>`           | Tritt einem Raum bei; wird abgelehnt, wenn bereits freigegeben wird oder keine Arbeitsmappe geöffnet ist.                                                                                                                                                                |
| `stopCollaboration()`            | `void`                    | Verlässt den Raum und leert `collaboration`; die Arbeitsmappe bleibt unverändert.                                                                                                                                                                                        |
| `reconnectCollaboration()`       | `void`                    | Startet die Verbindung neu und behält offline vorgenommene Änderungen.                                                                                                                                                                                                   |
| `share()`                        | `void`                    | Öffnet das Dialogfeld Freigeben.                                                                                                                                                                                                                                         |
| `focusGrid()`                    | `void`                    | Setzt den Tastaturfokus auf das Raster, sodass Eingaben die aktive Zelle bearbeiten.                                                                                                                                                                                     |
| `openCloudFile(file)`            | `Promise<void>`           | Öffnet eine Datei, die ein Host von einem `CloudFileProvider` erhalten hat, wie Datei > Öffnen > OneDrive: schreibgeschützt, wenn `file.canEdit` false ist; Speichern und Speichern unter werden an `file.storage` geleitet. Wird abgelehnt, wenn das Laden fehlschlägt. |
| `commitEdit()`                   | `boolean`                 | Übernimmt eine laufende Zellbearbeitung, bevor der Host navigiert oder speichert. `false` bedeutet, dass die Prüfung sie blockiert hat.                                                                                                                                  |
| `getToolbarProps()`              | `SpreadsheetToolbarProps` | Der headless Zustand der Symbolleiste und die `enabled`-Flags für benutzerdefinierte Steuerelemente ([Bausteine](/building-blocks)).                                                                                                                                     |
| `subscribeEditorState(listener)` | `() => void`              | Ruft `listener` bei jeder Änderung des Editorzustands auf; gibt die Funktion zum Abbestellen zurück.                                                                                                                                                                     |

## Ereignisse

Alle Ereignisse sind `CustomEvent`s, die aufsteigen (bubbling) und die Shadow-Grenze überqueren (`composed`).

| Ereignis               | `detail`                                                                                              | Hinweise                                                                                                                                             |
| ---------------------- | ----------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                                        | Nach einer Änderung, dem Laden oder einer neuen Arbeitsmappe.                                                                                        |
| `workbook-error`       | `{ error, message }`                                                                                  | Laden oder Speichern ist fehlgeschlagen.                                                                                                             |
| `workbook-warning`     | `{ warnings }`                                                                                        | Inhalt, den der Loader gelesen, aber nicht abbilden konnte.                                                                                          |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }`                                                                  | Der wirksame Schreibschutzzustand. `locked` wird durch eine Zusammenarbeitsrolle `viewer` gesetzt; `hostReadOnly` ist der eigene Schalter des Hosts. |
| `ribbon-action`        | `{ id }`                                                                                              | Ein Menübandbefehl wurde ausgeführt.                                                                                                                 |
| `ribbon-customize`     | `{ hiddenActions }`                                                                                   | Der Benutzer hat geändert, welche Menübandbefehle angezeigt werden (Datei > Optionen > Menüband anpassen).                                           |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' \| 'share' }` | **Abbrechbar**: `preventDefault()`, um den Befehl selbst zu behandeln.                                                                               |
| `selection-change`     | `{ sheet, ref, active }`                                                                              | `ref` ist der ausgewählte Bereich, `active` die aktive Zelle.                                                                                        |
| `sheet-change`         | `{ index, name }`                                                                                     |                                                                                                                                                      |
| `dirty-change`         | `{ dirty }`                                                                                           |                                                                                                                                                      |
| `office-ribbon-add-in` | `{ tab, command }`                                                                                    | Ein Befehl einer Host-Registerkarte (`ribbonAddIns`) wurde gewählt.                                                                                  |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                                          | Die Freigabe wurde gestartet oder beendet, die Verbindung hat sich geändert, oder jemand ist beigetreten oder gegangen.                              |
| `validation-error`     | `{ message, title?, ref, style? }`                                                                    | Eine laufende Zellbearbeitung ist an der Prüfung gescheitert und bleibt geöffnet, bis sie korrigiert oder abgebrochen wird.                          |

```ts
editor.addEventListener('file-command', (event) => {
	if (event.detail.command === 'save') {
		event.preventDefault();
		void editor
			.save()
			.then(uploadToMyStorage)
			.then(() => editor.markClean());
	}
});
```

## Speicher des Hosts {#host-owned-storage}

`storage` (eine Eigenschaft des Elements und eine Prop jedes Bindings) nimmt ein `DocumentStorage` aus
`ooxml-core/embedding` entgegen. Ist sie gesetzt:

- Datei > Speichern, Speichern unter und Strg+S übergeben die gespeicherte Arbeitsmappe an
  `storage.save(bytes, meta)`, statt sie herunterzuladen. `meta` ist `{ fileName, mimeType, reason }`,
  mit `reason` `'save'` oder `'saveAs'`; Quellen im Format `.xls` und `.csv` werden wie beim Download
  als `.xlsx` gespeichert. Sobald das Promise erfüllt ist, gilt die Arbeitsmappe als unverändert; wird
  es mit `{ fileName }` erfüllt, übernimmt die Titelleiste diesen Namen. Eine Ablehnung wird als
  `workbook-error` gemeldet, und die Arbeitsmappe bleibt geändert.
- Datei > Öffnen ruft `storage.open()` auf, wenn der Speicher diese Methode hat, und öffnet das
  `{ bytes, fileName }`, mit dem sie erfüllt wird (`undefined` bedeutet, dass der Benutzer abgebrochen
  hat). Ohne `open` wird die Dateiauswahl verwendet.
- Es läuft immer nur ein Speichervorgang: Speichern während eines laufenden Speichervorgangs wird
  ignoriert, nicht eingereiht. Eine Bearbeitung während des Speicherns beim Host lässt die Arbeitsmappe
  geändert, da die gespeicherten Bytes sie nicht enthalten.
- `file-command` wird weiterhin zuerst ausgelöst; ein Host, der `preventDefault()` aufruft, führt den
  Befehl selbst aus, und `storage` wird nicht aufgerufen.

Exportieren, Als CSV speichern und Drucken laden weiterhin im Browser herunter bzw. drucken dort. Excel
hat keine AutoWiederherstellung, daher werden die Wiederherstellungsmethoden von `DocumentStorage`
nicht verwendet.

```ts
editor.storage = {
	save: (bytes, { fileName }) => myDrive.put(fileName, bytes),
	open: () => myDrive.pick(), // resolves { bytes, fileName } or undefined
};
```

## Cloud-Speicherorte in Datei > Öffnen {#cloud-places-in-file-open}

`fileProviders` (eine Eigenschaft des Elements und eine Prop jedes Bindings) nimmt ein Array von
`CloudFileProvider` (`ooxml-core/embedding`) entgegen, zum Beispiel `createOneDriveFileProvider` aus
`ooxml-core/embedding/onedrive`. Leer oder weggelassen bleibt Datei > Öffnen unverändert. Mit Providern
listet Datei > Öffnen "Dieses Gerät" (Durchsuchen) und einen Speicherort pro Provider auf, der den
gemeinsamen Browser `<office-ui-cloud-files>` zeigt, beschränkt auf `.xlsx`. Seine Texte folgen dem
`locale` des Editors.

- Beim Öffnen einer Datei wird wie bei Öffnen gefragt, ob ungespeicherte Änderungen verworfen werden
  sollen; dann werden die Bytes über den normalen Ladeweg geladen, der Name der Datei übernommen und
  die Datei-Ansicht geschlossen.
- Die Arbeitsmappe speichert danach über den eigenen `storage` der Datei, nicht über den `storage` des
  Hosts: Speichern und Strg+S schreiben in dasselbe Element, Speichern unter schreibt eine Kopie, und
  die Arbeitsmappe gilt erst als unverändert, wenn der Provider erfüllt hat. Ein abgelehnter
  Speichervorgang wird als `workbook-error` gemeldet, und die Arbeitsmappe bleibt geändert. Das Öffnen,
  Erstellen oder Laden einer anderen Arbeitsmappe gibt Speichern an den `storage` des Hosts zurück.
- Eine Datei, die der Provider mit `canEdit: false` meldet, wird schreibgeschützt geöffnet (`readOnly`
  ist true, ein `readonly-change`-Ereignis wird ausgelöst). Es gilt das Strengere aus dem `readOnly`
  des Hosts und der Datei, und der eigene Wert des Hosts wird nicht geändert.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken,
		signIn,
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

## Host-eigene Menüs und Workflows {#host-menus-and-workflows}

Eine vom Host verwaltete Oberfläche (die [Bausteine](/de/building-blocks) oder `<xlsx-editor>` mit `surfaceOnly` und `surfaceActions`) hat keine eigenen Kontextmenüs und keine eigenen Dialoge für Zellen formatieren, Suchen und Ersetzen oder Funktion einfügen. Der Host zeichnet sie über die unten genannten Callbacks oder über die Operationen in `toolbarProps`. Das Verhalten beschreibt [Kontextmenüs und Workflow-Operationen](/de/building-blocks#context-menus-and-workflow-operations); dieser Abschnitt nennt die Typen und die Bedingungen, unter denen etwas abgelehnt wird.

Die Callbacks werden in `surfaceActions` gesetzt oder als Optionen an `useViewerBuildingBlocks`, `createViewerBuildingBlocks` und `mountEditorSurface` übergeben.

### Kontextmenü-Callbacks {#context-menu-callbacks}

```ts
onContextMenu?: (request: ContextMenuRequest) => boolean | void;
onContextMenuItems?: (request: ContextMenuRequest) => ContextMenuCustomization | void;
```

`onContextMenu` läuft bei einem Rechtsklick oder Umschalt+F10 auf einer Zelle, einer Zeilenüberschrift, einer Spaltenüberschrift oder einem Blattregister. Gibt er `true` zurück, wird das native Menü unterdrückt und das eigene gezeichnet; jeder andere Rückgabewert zeigt das native Menü. `onContextMenuItems` läuft nur, wenn `onContextMenu` nicht `true` zurückgegeben hat. Es filtert und erweitert das native Menü, das weiterhin selbst gezeichnet wird.

```ts
type ContextMenuKind = 'cell' | 'row' | 'column' | 'tab';

interface SurfaceActionContext {
	sheet: number;
	sheetName: string;
	selection: string; // A1-Text
	activeCell: string; // A1-Text
}

interface ContextMenuItem {
	id: string;
	label: string; // lokalisiert
	command?: string;
	enabled: boolean;
	checked?: boolean;
	shortcut?: string;
	separatorBefore?: boolean;
}

interface ContextMenuModel extends SurfaceActionContext {
	menu: ContextMenuKind;
	menuSheet: number; // das angeklickte Register bei 'tab', sonst das aktive Blatt
	readOnly: boolean;
	items: ContextMenuItem[];
	execute(id: string): Promise<boolean>;
}

interface ContextMenuRequest extends ContextMenuModel {
	x: number; // Client-Koordinaten des Rechtsklicks
	y: number;
}

interface HostContextMenuItem {
	id: string; // eindeutig unter den Host-Einträgen eines Menüs; eine leere oder wiederholte ID wird ignoriert
	label: string;
	enabled?: boolean | (() => boolean); // eine Funktion wird beim Ausführen erneut ausgewertet
	requiresEdit?: boolean; // in einer schreibgeschützten Arbeitsmappe deaktiviert und bei Ausführung abgelehnt
	checked?: boolean;
	separatorBefore?: boolean;
	position?: 'top' | 'bottom'; // Standard 'top'
	onSelect(request: ContextMenuRequest): void | Promise<void>;
}

interface ContextMenuCustomization {
	hide?: readonly string[]; // native Eintrags-IDs aus request.items
	add?: readonly HostContextMenuItem[];
}
```

`request.execute(id)` liefert `true`, wenn der Eintrag ausgeführt wurde. Es liefert `false`, wenn die laufende Zellbearbeitung die Prüfung nicht besteht (die Bearbeitung bleibt offen), wenn `id` nicht in `items` steht oder wenn der Eintrag in diesem Moment nicht ausführbar ist (schreibgeschützt, Blattschutz oder ein Befehl, der in der Surface nicht verfügbar ist). Ein Blattregister-Menü aktiviert zuerst sein Blatt. Host-Einträge aus `onContextMenuItems` führen `onSelect(request)` nach dem Schließen des Menüs aus, sofern sie bis dahin nicht deaktiviert sind.

### toolbarProps.contextMenu {#toolbarprops-contextmenu}

```ts
contextMenu(menu: ContextMenuKind, sheet?: number): ContextMenuModel;
```

Liefert dasselbe Modell ohne Rechtsklick, etwa für einen Host, der eigene Blattregister oder eine eigene Werkzeugleiste zeichnet. `sheet` ist standardmäßig das aktive Blatt und das Register, zu dem ein `'tab'`-Menü gehört. Befehle, die in der Surface nicht verfügbar sind, stehen mit `enabled: false` in der Liste. Zellen formatieren ist aktiviert, wenn `onOpenFormatCells` installiert und die Arbeitsmappe bearbeitbar ist. Einträge, die eine Aktion des mitgelieferten Blattregisters ausführen, etwa Umbenennen, stehen deaktiviert in der Liste.

### Workflow-Callbacks {#workflow-callbacks}

```ts
onOpenFormatCells?: (request: FormatCellsRequest) => void;
onOpenFindReplace?: (request: FindReplaceRequest) => void;
onOpenInsertFunction?: (request: InsertFunctionRequest) => void;

interface FormatCellsRequest extends SurfaceActionContext, FormatCellsOperations {}
interface FindReplaceRequest extends SurfaceActionContext, FindOperations {}
interface InsertFunctionRequest extends SurfaceActionContext, InsertFunctionOperations {}
```

Jede Anfrage enthält den oben genannten Kontext plus die Operationen dieses Workflows, unten beschrieben. Der Callback läuft nur, wenn der Workflow verfügbar ist: Zellen formatieren und Funktion einfügen brauchen Bearbeitungsrechte, Suchen und Ersetzen nicht. Die laufende Zellbearbeitung wird zuerst übernommen. Die `onOpen*`-Funktionen in `toolbarProps` liefern ein `Promise<boolean>`, das `false` ist, wenn der Callback fehlt, die Bearbeitungsrechte fehlen oder die laufende Bearbeitung die Prüfung nicht besteht.

### toolbarProps.workflows {#toolbarprops-workflows}

`toolbarProps.workflows` enthält dieselben Operationen ohne Callback, sodass ein dauerhaft offenes Panel keinen `onOpen*`-Callback braucht. Jede Operation liest den Schreibschutz, den Blattschutz und die laufende Zellbearbeitung beim Ausführen erneut.

#### Zellen formatieren {#format-cells-operations}

```ts
interface FormatCellsOperations {
	getState(): FormatCellsState;
	canApply(): boolean;
	apply(patch: StylePatch): boolean;
}

interface FormatAspect<T> {
	value: T; // die Eigenschaften, auf die sich alle geprüften Zellen einigen
	mixed: string[]; // die Namen der Eigenschaften, die abweichen
}

interface FormatCellsState {
	font: FormatAspect<Font>;
	alignment: FormatAspect<Alignment>;
	border: FormatAspect<Border>;
	protection: FormatAspect<Protection>;
	numFmt: string | null; // null, wenn die Auswahl Formate mischt
	fill: Fill | null; // null, wenn die Auswahl Füllungen mischt
	complete: boolean; // false, wenn die Auswahl mehr als 4096 Zellen umfasst
}
```

- `getState()` liefert die Formatierung der Auswahl. Ist `complete` `false`, sind die Listen der abweichenden Eigenschaften unvollständig.
- `canApply()` ist true, wenn die Arbeitsmappe bearbeitbar ist und der Blattschutz das Formatieren erlaubt.
- `apply(patch)` wendet einen `StylePatch` auf die Auswahl als einen Rückgängig-Schritt an. Es liefert `false` und ändert nichts, wenn `canApply()` `false` ist oder die laufende Zellbearbeitung die Prüfung nicht besteht.

#### Suchen und Ersetzen {#find-and-replace-operations}

```ts
interface FindQuery {
	text: string;
	sheet?: number;
	range?: CellRange; // nur zusammen mit sheet
	matchCase?: boolean;
	wholeCell?: boolean;
	lookIn?: 'formulas' | 'values' | 'comments';
	wildcards?: boolean; // Standard true
	order?: 'rows' | 'columns';
}

interface FindMatch {
	sheet: number;
	row: number;
	col: number;
	text: string;
}

interface FindOperations {
	canReplace(): boolean;
	findAll(query: FindQuery): FindMatch[];
	findNext(query: FindQuery, options?: { backwards?: boolean }): FindMatch | undefined;
	select(match: FindMatch): boolean;
	replace(query: FindQuery, replacement: string): boolean;
	replaceAll(query: FindQuery, replacement: string): number;
}
```

- `canReplace()` ist true, wenn die Arbeitsmappe bearbeitbar ist. Das Suchen braucht nie Bearbeitungsrechte.
- `findAll(query)` liefert jeden Treffer und funktioniert in einer schreibgeschützten Arbeitsmappe. Bei leerem `text` oder wenn der Editor nicht verbunden ist oder lädt, liefert es `[]`.
- `findNext(query, { backwards })` wählt den nächsten (oder vorherigen) Treffer nach der aktiven Zelle und beginnt bei Bedarf von vorn. Es liefert `undefined`, wenn es keinen Treffer gibt oder die laufende Zellbearbeitung die Prüfung nicht besteht.
- `select(match)` wählt einen Treffer aus `findAll`. Es liefert `false`, wenn der Editor nicht bereit ist oder das Blatt nicht existiert.
- `replace(query, replacement)` ersetzt den Treffer in der aktiven Zelle, falls es einen gibt, und wählt den nächsten Treffer. Es liefert `true` nur, wenn eine Zelle ersetzt wurde. Es liefert `false`, wenn die Arbeitsmappe schreibgeschützt ist, die laufende Bearbeitung die Prüfung nicht besteht oder die aktive Zelle eine gesperrte Zelle auf einem geschützten Blatt ist. Wie beim eigenen Ersetzen von Excel wird keine Datenüberprüfung angewendet.
- `replaceAll(query, replacement)` liefert die Anzahl geänderter Zellen. Es liefert `0`, wenn die Arbeitsmappe schreibgeschützt ist oder die laufende Bearbeitung die Prüfung nicht besteht. Durch Blattschutz gesperrte Zellen werden übersprungen. Datenüberprüfung wird nicht angewendet.

#### Funktion einfügen {#insert-function-operations}

```ts
interface FunctionDoc {
	name: string;
	category: string;
	syntax: string; // zum Beispiel SUM(number1, [number2], ...)
	description: string; // lokalisiert
}

interface InsertFunctionOperations {
	categories(): string[];
	search(options?: { text?: string; category?: 'all' | 'recent' | string }): FunctionDoc[];
	canInsert(): boolean;
	insert(name: string, args?: string): { text: string; caret: number } | undefined;
}
```

- `categories()` liefert die verschiedenen Kategorienamen, sortiert.
- `search(options)` vergleicht `text` zuerst mit dem Namen, dann mit den Wörtern der Beschreibung. `category` ist `'all'` (Standard), `'recent'` für zuletzt verwendete Funktionen oder ein Kategoriename.
- `canInsert()` ist true, wenn die Arbeitsmappe bearbeitbar ist.
- `insert(name, args?)` fügt `NAME(` bzw. `NAME(args)`, wenn `args` angegeben ist, an der Einfügemarke des offenen Formelentwurfs ein. Ist kein Entwurf offen, beginnt es mit `=NAME(`. Das Einfügen erfolgt nie vor dem führenden `=`. Es liefert danach den Entwurfstext `text` und die Marke `caret`, damit die Eingabe dort weitergehen kann. Es liefert `undefined` und ändert nichts, wenn die Arbeitsmappe schreibgeschützt ist, `name` keine dokumentierte Funktion ist oder der offene Entwurf reiner Text ist (er beginnt nicht mit `=`).

### toolbarProps.enabled-Flags {#toolbarprops-enabled}

| Flag                 | Wahr, wenn                                                                                                                  |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `contextMenu`        | `onContextMenu` ist installiert und der Editor ist verbunden.                                                               |
| `find`               | Der Editor ist verbunden, geladen und hat eine Sitzung. Keine Bearbeitungsrechte nötig.                                     |
| `replace`            | Die Arbeitsmappe ist bearbeitbar (`workflows.find.canReplace()`).                                                           |
| `applyFormatCells`   | Formatierung darf jetzt angewendet werden: bearbeitbar und der Blattschutz erlaubt es (`workflows.formatCells.canApply()`). |
| `insertFunctionCall` | Die Arbeitsmappe ist bearbeitbar (`workflows.insertFunction.canInsert()`).                                                  |
| `formatCells`        | `onOpenFormatCells` ist installiert und die Arbeitsmappe ist bearbeitbar.                                                   |
| `findReplace`        | `onOpenFindReplace` ist installiert.                                                                                        |
| `insertFunction`     | `onOpenInsertFunction` ist installiert und die Arbeitsmappe ist bearbeitbar.                                                |

Die Flags werden für jeden `toolbarProps`-Stand berechnet. Jede Operation prüft ihren Zustand beim Ausführen erneut.

### Beispiel {#host-menus-and-workflows-example}

```tsx
import { useViewerBuildingBlocks } from '@christophervr/xlsx-react-viewer';

const { toolbarProps } = useViewerBuildingBlocks({
	workbook,
	canEdit: true,
	onContextMenu(request) {
		if (request.menu !== 'cell') return; // Überschriften und Register behalten das native Menü
		showMyMenu(request.x, request.y, request.items, (id) => request.execute(id));
		return true; // natives Menü unterdrücken
	},
	onOpenFindReplace: ({ findAll, select }) => showFindPanel({ findAll, select }),
});

// Später, aus einem eigenen Panel (toolbarProps ist null, bis die Fläche eingebunden ist):
toolbarProps?.workflows.formatCells.apply({ font: { bold: true } });
```

## Shadow Parts

Das Element stellt seine Bereiche als CSS-Shadow-Parts bereit, zum Gestalten mit `xlsx-editor::part(grid)` und für Tests: `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs`, `status-bar`, `title-bar`, `backstage` und `page-layout` (Ansicht Seitenlayout).
