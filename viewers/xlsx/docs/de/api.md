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

| Eigenschaft          | Typ                                  | Hinweise                                                                                                             |
| -------------------- | ------------------------------------ | -------------------------------------------------------------------------------------------------------------------- |
| `themeColors`        | `Partial<Record<...>>`               | Token-Überschreibungen, angewendet als benutzerdefinierte Eigenschaften `--xve-*` ([Themes](/theming)).              |
| `hiddenActions`      | `string[]`                           | Menüband-Steuerelemente, Titelleistenelemente und Kontextmenüeinträge, die über ihre stabile ID ausgeblendet werden. |
| `ribbonAddIns`       | `RibbonAddInTab[]`                   | Registerkarten, die der Host nach den eingebauten ergänzt ([Bindings](/bindings#ribbon-add-in-tabs)).                |
| `customization`      | `ExcelCustomization \| null`         | Capabilities, Host-Seiten im Datei-Register und Host-Menübefehle (siehe [Anpassung](#customization)).                |
| `workbook`           | `Workbook \| null`                   | Liefert die geöffnete Arbeitsmappe oder ersetzt sie durch eine gesetzte.                                             |
| `dirty`              | `boolean` (nur lesen)                | Ungespeicherte Änderungen.                                                                                           |
| `activeSheet`        | `number`                             | Nullbasierter Index des sichtbaren Blatts.                                                                           |
| `collaboration`      | `XlsxCollaborationOptions \| null`   | Gibt die geöffnete Arbeitsmappe in einem Raum frei, solange sie gesetzt ist ([Zusammenarbeit](/de/collaboration)).   |
| `collaborationState` | `XlsxCollaborationState` (nur lesen) | Ob freigegeben wird, der Verbindungsstatus und wer sich im Raum befindet.                                            |

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

| Methode                       | Rückgabe              | Hinweise                                                                                                                                        |
| ----------------------------- | --------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`      | `Promise<void>`       | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv`; das Format wird an den Bytes erkannt. Löst bei einem Fehler `workbook-error` aus und wird abgelehnt. |
| `newWorkbook()`               | `void`                | Eine leere Arbeitsmappe namens `Book1.xlsx`.                                                                                                    |
| `save()`                      | `Promise<Blob>`       | `.xlsx`. Setzt `dirty` nicht zurück.                                                                                                            |
| `saveBytes(format?)`          | `Promise<Uint8Array>` | `'xlsx'` (Standard) oder `'csv'` (das aktive Blatt, formatierte Werte).                                                                         |
| `download(name?)`             | `Promise<void>`       | Speichert, lädt herunter und ruft dann `markClean()` auf.                                                                                       |
| `markClean()`                 | `void`                |                                                                                                                                                 |
| `select(ref)`                 | `void`                | Eine A1-Zelle oder ein A1-Bereich auf dem aktiven Blatt (`'C3'`, `'A1:D20'`).                                                                   |
| `getSelection()`              | `string`              | Die aktuelle Auswahl als A1-Bezug.                                                                                                              |
| `setActiveSheet(index)`       | `void`                |                                                                                                                                                 |
| `undo()`, `redo()`            | `void`                | Bei Freigabe werden nur die Änderungen dieses Fensters rückgängig gemacht.                                                                      |
| `startCollaboration(options)` | `Promise<void>`       | Tritt einem Raum bei; wird abgelehnt, wenn bereits freigegeben wird oder keine Arbeitsmappe geöffnet ist.                                       |
| `stopCollaboration()`         | `void`                | Verlässt den Raum und leert `collaboration`; die Arbeitsmappe bleibt unverändert.                                                               |
| `reconnectCollaboration()`    | `void`                | Startet die Verbindung neu und behält offline vorgenommene Änderungen.                                                                          |
| `share()`                     | `void`                | Öffnet Datei > Freigeben.                                                                                                                       |
| `focusGrid()`                 | `void`                | Setzt den Tastaturfokus auf das Raster, sodass Eingaben die aktive Zelle bearbeiten.                                                            |

## Ereignisse

Alle Ereignisse sind `CustomEvent`s, die aufsteigen (bubbling) und die Shadow-Grenze überqueren (`composed`).

| Ereignis               | `detail`                                                                                   | Hinweise                                                                                                                    |
| ---------------------- | ------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                             | Nach einer Änderung, dem Laden oder einer neuen Arbeitsmappe.                                                               |
| `workbook-error`       | `{ error, message }`                                                                       | Laden oder Speichern ist fehlgeschlagen.                                                                                    |
| `workbook-warning`     | `{ warnings }`                                                                             | Inhalt, den der Loader gelesen, aber nicht abbilden konnte.                                                                 |
| `readonly-change`      | `{ readOnly }`                                                                             |                                                                                                                             |
| `ribbon-action`        | `{ id }`                                                                                   | Ein Menübandbefehl wurde ausgeführt.                                                                                        |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' }` | **Abbrechbar**: `preventDefault()`, um den Befehl selbst zu behandeln.                                                      |
| `selection-change`     | `{ sheet, ref, active }`                                                                   | `ref` ist der ausgewählte Bereich, `active` die aktive Zelle.                                                               |
| `sheet-change`         | `{ index, name }`                                                                          |                                                                                                                             |
| `dirty-change`         | `{ dirty }`                                                                                |                                                                                                                             |
| `office-ribbon-add-in` | `{ tab, command }`                                                                         | Ein Befehl einer Host-Registerkarte (`ribbonAddIns`) wurde gewählt.                                                         |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                               | Die Freigabe wurde gestartet oder beendet, die Verbindung hat sich geändert, oder jemand ist beigetreten oder gegangen.     |
| `validation-error`     | `{ message, title?, ref, style? }`                                                         | Eine laufende Zellbearbeitung ist an der Prüfung gescheitert und bleibt geöffnet, bis sie korrigiert oder abgebrochen wird. |

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

## Shadow Parts

Das Element stellt seine Bereiche als CSS-Shadow-Parts bereit, zum Gestalten mit `xlsx-editor::part(grid)` und für Tests: `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs` und `status-bar`.
