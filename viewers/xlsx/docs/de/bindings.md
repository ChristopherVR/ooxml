# Framework-Bindings

Alle Bindings binden `<xlsx-editor>` über dieselbe Funktion `mountEditor` in `packages/bindings/src/index.ts` ein, sodass jedes Framework dieselbe Semantik erhält.

## Props {#props}

| Prop             | Typ                           | Standard     | Hinweise                                                                                                                                                                                  |
| ---------------- | ----------------------------- | ------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook`       | `Workbook`                    | keiner       | Wird angezeigt, wenn ein **neues Objekt** übergeben wird. Die Arbeitsmappe, die der Editor ausgibt, wird nie zurückgeschrieben (keine Schleife).                                          |
| `bytes`          | `Uint8Array \| ArrayBuffer`   | keiner       | Dateibytes; wird neu geladen, sobald ein neues Array übergeben wird.                                                                                                                      |
| `src`            | `string`                      | keiner       | URL, die geladen wird, sobald sich die Zeichenkette ändert; der Dateiname ist standardmäßig das letzte Segment der URL.                                                                   |
| `fileName`       | `string`                      | `Book1.xlsx` | Wird nur weitergegeben, wenn das übergeordnete Element ihn ändert, damit Datei > Öffnen im Editor die Arbeitsmappe umbenennen kann.                                                       |
| `readOnly`       | `boolean`                     | `false`      |                                                                                                                                                                                           |
| `locale`         | `string`                      | `en`         | `en`, `fr`, `de`, `es`, `zh-CN` oder ein Tag, das auf eine dieser Sprachen abgebildet wird (`de-DE`).                                                                                     |
| `theme`          | `'light' \| 'dark' \| 'auto'` | `auto`       | `auto` folgt dem Betriebssystem.                                                                                                                                                          |
| `authorName`     | `string`                      | `Author`     | Wird bei neuen Kommentaren erfasst.                                                                                                                                                       |
| `showToolbar`    | `boolean`                     | `true`       | Das Menüband.                                                                                                                                                                             |
| `showFormulaBar` | `boolean`                     | `true`       | Das Namensfeld und die Bearbeitungsleiste.                                                                                                                                                |
| `hiddenActions`  | `string[]`                    | `[]`         | Menübandsteuerelemente, Elemente der Titelleiste und Einträge des Kontextmenüs, die über ihre stabile ID ausgeblendet werden.                                                             |
| `ribbonAddIns`   | `RibbonAddInTab[]`            | keiner       | Registerkarten, die der Host nach den integrierten hinzufügt; siehe [Add-in-Registerkarten für das Menüband](#ribbon-add-in-tabs).                                                        |
| `themeColors`    | `Partial<Record<...>>`        | `{}`         | Überschreibungen von Token, siehe [Themes](/de/theming).                                                                                                                                  |
| `collaboration`  | `XlsxCollaborationOptions`    | keiner       | Teilt die geöffnete Arbeitsmappe in einem Raum, solange der Wert gesetzt ist; ein neues Objekt tritt erneut bei, `null` verlässt den Raum. Siehe [Zusammenarbeit](/de/collaboration).     |
| `storage`        | `DocumentStorage`             | keiner       | Vom Host verwalteter Speicher: Speichern und Speichern unter gehen an `storage.save()`, Datei > Öffnen an `storage.open()`, sofern vorhanden.                                             |
| `fileProviders`  | `CloudFileProvider[]`         | keiner       | Cloud-Speicherorte, die Datei > Öffnen neben „Dieses Gerät“ anbietet; eine Arbeitsmappe, die von dort geöffnet wurde, wird in diese Datei zurückgespeichert. Halten Sie das Array stabil. |
| `customization`  | `ExcelCustomization`          | keiner       | `capabilities`, `backstage.extraPages` und `contextMenu.extraCommands`; halten Sie das Objekt stabil (`null` stellt die Standardwerte wieder her).                                        |

## Callbacks und Ereignisse {#callbacks-and-events}

| Callback (React, Solid, Vanilla) | Vue-Ereignis           | Angular-Ausgabe       | Svelte-Prop             | Nutzlast                                     |
| -------------------------------- | ---------------------- | --------------------- | ----------------------- | -------------------------------------------- |
| `onWorkbookChange`               | `workbook-change`      | `workbookChange`      | `onworkbookchange`      | `Workbook`                                   |
| `onWorkbookError`                | `workbook-error`       | `workbookError`       | `onworkbookerror`       | `Error`                                      |
| `onSelectionChange`              | `selection-change`     | `selectionChange`     | `onselectionchange`     | `{ sheet, ref, active }`                     |
| `onDirtyChange`                  | `dirty-change`         | `dirtyChange`         | `ondirtychange`         | `boolean`                                    |
| `onReadOnlyChange`               | `readonly-change`      | `readOnlyChange`      | `onreadonlychange`      | `boolean`                                    |
| `onFileCommand`                  | `file-command`         | `fileCommand`         | `onfilecommand`         | `CustomEvent<FileCommandDetail>`, abbrechbar |
| `onWorkbookWarning`              | `workbook-warning`     | `workbookWarning`     | `onworkbookwarning`     | `string[]`                                   |
| `onRibbonAction`                 | `ribbon-action`        | `ribbonAction`        | `onribbonaction`        | `{ id }`                                     |
| `onRibbonCustomize`              | `ribbon-customize`     | `ribbonCustomize`     | `onribboncustomize`     | `string[]` (ausgeblendete Aktionen)          |
| `onSheetChange`                  | `sheet-change`         | `sheetChange`         | `onsheetchange`         | `{ index, name }`                            |
| `onRibbonAddIn`                  | `office-ribbon-add-in` | `ribbonAddIn`         | `onribbonaddin`         | `{ tab, command }`                           |
| `onCollaborationChange`          | `collaboration-change` | `collaborationChange` | `oncollaborationchange` | `XlsxCollaborationState`                     |
| `onReady`                        | `ready`                | `ready`               | `onready`               | das Element `<xlsx-editor>`                  |

Ladefehler (beschädigte, verschlüsselte oder nicht unterstützte Dateien) erreichen `onWorkbookError`; ein fehlgeschlagenes Laden von `src` ebenso. Ein `src`-Abruf, der noch läuft, wird verworfen, sobald neue `bytes` oder eine neue `workbook` ankommen oder `src` geleert oder geändert wird.

`onFileCommand` erhält das abbrechbare Ereignis `file-command` selbst und nicht dessen Detail: Rufen Sie `event.preventDefault()` auf, um einen Befehl aus dem Menü Datei (Neu, Öffnen, Speichern, Speichern unter, Exportieren, CSV exportieren, Drucken, Freigeben) zu übernehmen, und lesen Sie `event.detail.command`, um festzustellen, welcher Befehl es ist. Vue, Angular und Svelte geben dasselbe Ereignisobjekt aus bzw. übergeben es an den Callback. Jeder andere Callback erhält seine Nutzlast direkt. Siehe [Element-API](/de/api).

Jede Prop wird nur weitergegeben, wenn sich ihr Wert ändert (`hiddenActions` und `themeColors` werden nach Wert verglichen). So macht ein erneutes Rendern des übergeordneten Elements nie rückgängig, was der Nutzer im Editor geändert hat: Bearbeiten / Anzeigen, Datei > Optionen für Design oder Sprache oder Menüband anpassen. Ein kontrolliertes übergeordnetes Element hält seinen Zustand über `onReadOnlyChange` und `onRibbonCustomize` synchron (in Angular über `[(readOnly)]`).

## Add-in-Registerkarten für das Menüband {#ribbon-add-in-tabs}

Ein Host kann, wie ein Office-Add-in, eigene Registerkarten des Menübands nach denen von Excel hinzufügen.
`ribbonAddIns` nimmt einfache Daten entgegen und ist eine Prop jedes Bindings. Ein gewählter Befehl
führt seinen Callback `run` aus und wird über `onRibbonAddIn` (Vue
`office-ribbon-add-in`, Angular `ribbonAddIn`, Svelte `onribbonaddin`) mit
`{ tab, command }` gemeldet.

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [
					{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() },
					{ id: 'options', label: 'Options', size: 'small' },
					{ id: 'send', label: 'Send', items: [{ id: 'send-mail', label: 'By mail' }] },
				],
			},
		],
	},
];

<SpreadsheetEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

Befehle mit `size: 'small'` füllen Spalten zu dritt, und `items` erzeugt eine Dropdown-Liste.
Wird ein gleichwertiger Beschreibungseintrag erneut übergeben, wie ein Framework es bei jedem Rendern tut, bleiben die Bereiche erhalten und nur die Callbacks werden ersetzt; lassen Sie die Prop weg, entfernen sich die Registerkarten. Eine Registerkarte darf die ID einer integrierten Registerkarte nicht übernehmen, und ihre Befehle erreichen nie die eigene Befehlsbehandlung des Editors. Derselbe Beschreibungseintrag (`RibbonAddInTab`, aus `ooxml-core/ribbon`, von jedem Framework-Paket erneut exportiert) funktioniert in den Editoren für Word, Excel, PowerPoint und Visio. Der Editor liefert keine Add-in-Registerkarte mit.

## Handle {#handle}

Jeder Adapter stellt dasselbe Handle bereit: React über `ref`, Vue über die Template-Referenz, Angular über die Komponenteninstanz, Solid über `editorRef`, Svelte über die Exporte der Komponente und Vanilla als Rückgabewert von `mountEditor`.

| Member                          | Wirkung                                                                                                                                                          |
| ------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `element`                       | Das eingebundene `<xlsx-editor>`.                                                                                                                                |
| `load(bytes, fileName?)`        | Öffnet `.xlsx`-, `.xlsm`-, `.xls`- oder `.csv`-Bytes; lehnt bei einem Fehler ab.                                                                                 |
| `newWorkbook()`                 | Ersetzt die Arbeitsmappe durch eine leere.                                                                                                                       |
| `save()`                        | Liefert einen `.xlsx`-Blob. Setzt `dirty` nicht zurück.                                                                                                          |
| `saveBytes(format?)`            | Bytes als `'xlsx'` (Standard) oder als `'csv'` (aktives Blatt).                                                                                                  |
| `download(fileName?)`           | Speichert, lädt im Browser herunter und markiert die Arbeitsmappe als unverändert.                                                                               |
| `markClean()`                   | Setzt `dirty` zurück, nachdem Sie einen gespeicherten Blob selbst abgelegt haben.                                                                                |
| `select(ref)`, `getSelection()` | Auswahl als A1-Referenz (`'B2:D8'`).                                                                                                                             |
| `setActiveSheet(index)`         | Wechselt das Blatt.                                                                                                                                              |
| `share()`                       | Öffnet den Dialog Freigeben. Das Entfernen des Editors verlässt jeden freigegebenen Raum.                                                                        |
| `openCloudFile(file)`           | Öffnet eine Datei aus einem `CloudFileProvider`, so wie Datei > Öffnen > OneDrive es tut: mit ihrem Namen, ihrem Schreibschutz und ihrem Speicher für Speichern. |
| `dirty`                         | Ungespeicherte Änderungen. Svelte-`isDirty()` ist veraltet; lesen Sie `dirty`.                                                                                   |

Der Aufruf einer Handle-Methode, bevor der Editor eingebunden ist, löst `Editor is not mounted` aus.

## Serverseitiges Rendern {#server-rendering}

Jedes Paket lässt sich ohne DOM importieren. Die Registrierung (`defineXlsxEditor`) und das Einbinden sind Vorgänge auf dem Client.
