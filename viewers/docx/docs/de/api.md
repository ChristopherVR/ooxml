# Element-API und Ereignisse

Jeder Framework-Adapter umhüllt dasselbe Custom Element `<docx-editor>`. Die Adapter leiten diese Eigenschaften weiter und geben diese Ereignisse in der Idiomatik ihres Frameworks wieder aus (siehe [Framework-Bindings](/bindings)).

## Attribute und Eigenschaften

| Eigenschaft                  | Attribut          | Beschreibung                                                                                                                                                                                                                   |
| ---------------------------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `documentModel`              |                   | Das angezeigte `DocumentModel`. Eine Zuweisung ist ein externer Austausch, kein Bearbeitungsereignis.                                                                                                                          |
| `readOnly`                   | `read-only`       | Blockiert lokale Änderungen. Stapel der Zusammenarbeit von anderen Teilnehmern werden weiter angewendet.                                                                                                                       |
| `locale`                     | `locale`          | Oberflächensprache: `en`, `fr`, `de`, `es` oder `zh-CN`. Dokumentinhalte werden nie übersetzt.                                                                                                                                 |
| `fileName`                   | `file-name`       | Name des geöffneten Dokuments (Standard `Document1.docx`).                                                                                                                                                                     |
| `reviewAuthor`               | `review-author`   | Autor, der bei nachverfolgten Änderungen und Kommentaren erfasst wird (Standard `Author`).                                                                                                                                     |
| `theme`                      | `theme`           | `light`, `dark` oder `auto`. Siehe [Theming](/theming).                                                                                                                                                                        |
| `themeColors`                |                   | Teilweise Token-Überschreibungen, angewendet als benutzerdefinierte Eigenschaften `--dve-*`.                                                                                                                                   |
| `showThumbnails`             | `show-thumbnails` | Linke Leiste mit Seitenminiaturen (nur Drucklayout).                                                                                                                                                                           |
| `showToolbar`                | `show-toolbar`    | `false` blendet das Menüband aus; Titelleiste und Statusleiste bleiben.                                                                                                                                                        |
| `hiddenActions`              |                   | Auszublendende Menübandbefehle nach stabiler ID; siehe die Tabelle unter [Bindings](/bindings).                                                                                                                                |
| `ribbonAddIns`               |                   | Register, die der Host nach den eingebauten hinzufügt; siehe [Bindings](/bindings#ribbon-add-in-tabs).                                                                                                                         |
| `collaboration`              |                   | Deklarative Zusammenarbeit (`DocxCollaborationOptions` oder `null`); siehe [Zusammenarbeit](/de/collaboration#die-eigenschaft-collaboration).                                                                                  |
| `customization`              |                   | Capabilities, Host-Seiten im Datei-Register und Host-Menübefehle (siehe [Anpassung](#customization)).                                                                                                                          |
| `surfaceOnly`                |                   | Bindet nur das Dokument und seine Bearbeitungsfläche ein. Vor dem Verbinden des Elements wählen.                                                                                                                               |
| `surfaceActions`             |                   | Host-Dialoge für den Oberflächenmodus: `onOpenHyperlinkDialog`, `onOpenFindReplace` und `onOpenImagePicker`.                                                                                                                   |
| `googleFonts`                |                   | Webfonts für Dokumentschriften, die weder installiert noch eingebettet sind. Ohne Wert gilt Datei > Optionen > Trust Center (aus, bis es eingeschaltet wird); `false` schaltet sie aus; `{ baseUrl }` setzt den CSS2-Endpunkt. |
| `storage`                    |                   | Vom Host verwalteter Speicher für Speichern, Speichern unter und Öffnen; siehe „Speicher des Hosts“ weiter unten.                                                                                                              |
| `fileProviders`              |                   | Cloud-Orte für Datei > Öffnen; siehe „Cloud-Orte in Datei > Öffnen“ weiter unten.                                                                                                                                              |
| `loading` (lesen)            |                   | True, während ein Dokument geladen wird.                                                                                                                                                                                       |
| `collaborationState` (lesen) |                   | Der Zusammenarbeitsraum: ob er aktiv ist, seine Raum-ID, Status, Synchronisationszustand und Personen.                                                                                                                         |
| `dirty` (lesen)              |                   | True nach einer Änderung, false nach Speichern, Laden oder `markClean()`.                                                                                                                                                      |

## Anpassung {#customization}

`customization` nimmt ein `WordCustomization`-Objekt entgegen, den gemeinsamen Vertrag von `ooxml-core/embedding`, den auch die anderen Editoren verwenden. Jedes Binding reicht es als Prop bzw. Input `customization` weiter. Es wird angewendet, wenn sich die Objektidentität ändert; halten Sie es daher stabil. `null` stellt die Standardwerte wieder her.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [
			{ id: 'send', label: 'Send to chat', onSelect: ({ selectedText }) => send(selectedText) },
		],
	},
};
```

- `capabilities`: `false` entfernt jeden Weg zur Funktion (`edit`, `print`, `export`, `save`, `open`, `share`, `copy`, `comment`, `review`, `find`, `settings`): Menüband, Seiten im Datei-Register, Menüs und Tastenkürzel. `hiddenActions` bleibt unverändert; beide werden vereinigt.
- `backstage.extraPages`: Host-Seiten `{ id, label, group?, order?, render }` im Datei-Register, nach den eingebauten Seiten ihrer Gruppe. `render` ist eine Mount-Funktion, ein Custom-Element-Tag oder `{ element, properties }` und erhält einen `WordRegionContext`.
- `contextMenu.extraCommands`: Host-Befehle `{ id, label, icon?, group?, when?, disabled?, onSelect }` oberhalb (`top`, Standard) oder unterhalb (`bottom`) der eingebauten Einträge. `onSelect` läuft nach dem Schließen des Menüs mit einem `WordMenuContext`. Beschriftungen werden unverändert angezeigt (lokalisieren Sie sie selbst).

`edit: false` ist eine eigene Schreibsperre neben dem schreibgeschützten Schalter des Hosts und einer Zusammenarbeitsrolle `viewer`: Keine hebt eine andere auf. Capabilities sind eine UI-Richtlinie, kein Dokumentschutz. Das Ersetzen von Bereichen (Optionsdialog, Bereiche) wird für Word noch nicht angeboten. Alle Produkte im Überblick: [Einbettungsleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Methoden

| Methode                                                         | Beschreibung                                                                                                                                                    |
| --------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes)`                                                   | Lädt DOCX- oder ältere DOC-Bytes. Bei einem Fehler wird abgelehnt und `document-error` ausgelöst.                                                               |
| `save()`                                                        | Liefert einen `Blob` des serialisierten Dokuments.                                                                                                              |
| `saveBytes()`                                                   | Liefert das rohe `Uint8Array`.                                                                                                                                  |
| `download(fileName?)`                                           | Speichert und startet einen Browser-Download.                                                                                                                   |
| `markClean()`                                                   | Setzt das Dirty-Flag zurück, nachdem der Host das Dokument gespeichert hat.                                                                                     |
| `setLoadedDocument(session)`                                    | Bindet ein von `loadDocument` bereits geparstes `LoadedDocument` ein.                                                                                           |
| `openCloudFile(file)`                                           | Öffnet eine Datei aus einem `CloudFileProvider`, wie Datei > Öffnen > OneDrive.                                                                                 |
| `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` | Zustand und Callbacks für vom Host gebaute Werkzeugleisten-, Navigations- und Statussteuerelemente.                                                             |
| `subscribeEditorState(listener)`                                | Ruft `listener` bei Auswahl- und Formatänderungen auf, auch bei solchen, die das Dokument nicht als geändert markieren. Gibt eine Funktion zum Abmelden zurück. |
| `startCollaboration(config)`                                    | Tritt einer Sitzung zur Zusammenarbeit bei; siehe [Zusammenarbeit](/de/collaboration).                                                                          |
| `startYjsCollaboration(session, options)`                       | Tritt einem synchronisierten Yjs-Raum mit dem passenden Quellpaket bei.                                                                                         |
| `reconnectCollaboration()`, `resyncCollaboration()`             | Stellt die Yjs-Verbindung wieder her oder fordert einen Zustandsaustausch des Providers an.                                                                     |
| `getPendingCollaboration()`                                     | Der ausstehende Schrittstapel oder `null`.                                                                                                                      |
| `receiveCollaboration(batch)`                                   | Wendet einen von der Autorität akzeptierten Stapel an.                                                                                                          |
| `stopCollaboration(discardPending?)`                            | Verlässt die Sitzung; ausstehende Änderungen müssen bestätigt oder ausdrücklich verworfen werden.                                                               |
| `publishPresence`, `receivePresence`, `leavePresence`           | Flüchtiger Präsenzkanal; siehe [Zusammenarbeit](/de/collaboration#präsenz).                                                                                     |

## Ereignisse

Alle Ereignisse sind bubbling und composed und überschreiten daher die Shadow-Grenze.

| Ereignis               | Detail                               | Bedeutung                                                                                                |
| ---------------------- | ------------------------------------ | -------------------------------------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`                      | Das Dokument wurde durch Bearbeiten oder Überarbeitungsaktionen geändert.                                |
| `document-error`       | `Error`                              | Laden, Speichern, Drucken oder eine Bearbeitung ist fehlgeschlagen.                                      |
| `document-warning`     | `string`                             | Nicht schwerwiegender Hinweis.                                                                           |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }` | Der Benutzer hat in der Oberfläche zwischen Bearbeiten und Anzeigen umgeschaltet.                        |
| `file-command`         | `{ command, fileName? }`             | Abbrechbar. `command` ist `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` oder `share`. |
| `dirty-change`         | `boolean`                            | Der Zustand „ungespeicherte Änderungen“ hat sich umgekehrt.                                              |
| `page-change`          | `{ page, pageCount }`                | Seite oder Seitenanzahl im Drucklayout hat sich geändert (näherungsweise Paginierung).                   |
| `ribbon-action`        | `RibbonAction`                       | Ein Menübandsteuerelement wurde aktiviert.                                                               |
| `ribbon-customize`     | `readonly string[]`                  | Der Benutzer hat die angezeigten Menübandbefehle geändert; speichern Sie die IDs bei Bedarf.             |
| `office-ribbon-add-in` | `{ tab, command }`                   | Ein Befehl eines Host-Registers (`ribbonAddIns`) wurde gewählt.                                          |
| `presence-send`        | `PresenceMessage`                    | Die lokale Präsenz hat sich geändert; übertragen Sie sie an andere Clients.                              |
| `collaboration-send`   | `StepBatch`                          | Lokale Schritte stehen aus; übertragen Sie sie an die Autorität.                                         |
| `collaboration-change` | `DocxCollaborationState`             | Der Zusammenarbeitsstatus hat sich geändert: `{ active, roomId, status, synced, people }`.               |

Das Paket exportiert die typisierte Ereigniszuordnung (`DocxEditorEventMap`) und die Laufzeitliste `DOCX_EDITOR_EVENTS`.

Die Framework-Bindings leiten die Props `fileName`, `reviewAuthor` und `themeColors` sowie die Ereignisse `file-command`, `document-warning`, `readonly-change`, `ribbon-action` und `ribbon-customize` unter den Namen ihres Frameworks weiter (zum Beispiel `onFileCommand`, `fileCommand` und `onfilecommand`). Die Tabelle steht unter [Framework-Bindings](/bindings#binding-props-and-events).

## Shadow-Parts

Die Titelleiste (Dateiname, Speicherstatus, Schnellzugriff und „Was möchten Sie tun?“) ist der Shadow-Part `title-bar`.
Ein Host, der seine eigene Dokumentleiste zeichnet, blendet sie mit `docx-editor::part(title-bar) { display: none; }` aus;
Menüband und Statusleiste bleiben.

## Cloud-Orte in Datei > Öffnen

`fileProviders` (eine Eigenschaft des Elements und eine Prop jedes Bindings) nimmt
`CloudFileProvider`s aus `ooxml-core/embedding` entgegen, zum Beispiel `createOneDriveFileProvider`.
Datei > Öffnen listet dann „Dieses Gerät“ und einen Ort pro Provider auf. Eine von einem Ort
geöffnete Datei wird über dessen eigenen Speicher gespeichert und ist schreibgeschützt, wenn der
Provider meldet, dass sie nicht bearbeitet werden kann; siehe
[OneDrive-Integration](/onedrive#open-from-onedrive-in-file-open).

## Speicher des Hosts

`storage` (eine Eigenschaft des Elements und eine Prop jedes Bindings) nimmt einen `DocumentStorage`
aus `ooxml-core/embedding` entgegen. Ist er gesetzt:

- Datei > Speichern, Speichern unter und Strg+S übergeben die gespeicherten Bytes an
  `storage.save(bytes, meta)`, statt sie herunterzuladen. `meta` ist `{ fileName, mimeType, reason }`,
  mit `reason` `'save'` oder `'saveAs'`. Sobald das Promise erfüllt ist, gilt das Dokument als
  gespeichert; wird es mit `{ fileName }` erfüllt, übernimmt die Titelleiste diesen Namen. Eine
  Ablehnung wird als `document-error` gemeldet, und das Dokument bleibt geändert (dirty).
- Datei > Öffnen ruft `storage.open()` auf, wenn der Speicher diese Methode hat, und öffnet das
  zurückgegebene `{ bytes, fileName }` (`undefined` bedeutet, dass der Benutzer abgebrochen hat). Ohne
  `open` wird die Dateiauswahl verwendet.
- Es läuft immer nur ein Speichervorgang: Speichern während eines laufenden Speicherns wird
  ignoriert, nicht eingereiht. Eine Änderung während des Speicherns durch den Host lässt das Dokument
  geändert, da die gespeicherten Bytes sie nicht enthalten.
- `file-command` wird weiterhin zuerst ausgelöst. Ein Host, der `preventDefault()` aufruft,
  behandelt den Befehl selbst, und `storage` wird nicht aufgerufen. Speichern unter meldet `saveAs`
  und dann `save`.
- Ein Host, der Speichern entfernt hat (`customization.capabilities.save: false`), erhält weder das
  Ereignis `save` noch `storage.save`; Speichern unter (die Capability `export`) lädt dann nur herunter.

Exportieren, Als Text exportieren und Drucken verwenden weiterhin den Browser. Word hat keine
AutoWiederherstellung, daher werden die Wiederherstellungsmethoden von `DocumentStorage` nicht
verwendet.

```ts
editor.storage = {
	async save(bytes, { fileName, mimeType }) {
		await fetch(`/files/${encodeURIComponent(fileName)}`, {
			method: 'PUT',
			headers: { 'content-type': mimeType },
			body: bytes,
		});
	},
};
```

## Vorlesen {#read-aloud}

Der Building-Blocks-Hook (`useViewerBuildingBlocks`) legt fünf Felder zum Vorlesen auf `toolbarProps` ab.
Sie sind keine Eigenschaften oder Ereignisse des Elements. Sie nutzen die eigene Implementierung des Editors
(die Sprachsynthese des Browsers; nichts wird gesendet). Das Verhalten beschreibt der
[Building-Blocks-Leitfaden](/de/building-blocks#read-aloud).

| Feld              | Typ                      | Verhalten                                                                                                                                                                            |
| ----------------- | ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `canReadAloud`    | `boolean`                | Wahr, wenn ein Dokument angezeigt wird und der Browser sprechen kann. Schreibgeschützte Dokumente zählen mit.                                                                        |
| `readingAloud`    | `boolean`                | Wahr nur in dem Editor, der die Stimme besitzt. Liest ein anderer Editor, ist er falsch.                                                                                             |
| `readAloudIssue`  | `ReadAloudIssue \| null` | Das letzte Problem: `'unsupported'`, solange der Browser keine Sprachsynthese hat, sonst `'empty'` oder `'error'`. `null`, wenn keines vorliegt.                                     |
| `onReadAloud`     | `() => ReadAloudStart`   | Liest die Auswahl oder vom Cursor bis zum Ende, wenn nichts markiert ist. Startet neu, wenn bereits gelesen wird. Gibt `'started'` zurück oder den Grund, warum nichts begonnen hat. |
| `onStopReadAloud` | `() => boolean`          | Beendet die Wiedergabe dieses Editors. Gibt zurück, ob sie lief.                                                                                                                     |

Die Codes sind einfache Zeichenkettenvereinigungen. Die Bibliothek liefert nur die Codes; die Meldung
formuliert der Host.

```ts
type ReadAloudIssue = 'unsupported' | 'empty' | 'error';
type ReadAloudStart = 'started' | ReadAloudIssue;
```

`'error'` bedeutet, dass die Stimme ausgefallen ist. Ein abgebrochener oder unterbrochener Vortrag wird
nicht als Fehler gemeldet.

Ein React-Schaltknopf des Hosts, der das Vorlesen umschaltet und den leeren Fall formuliert:

```tsx
import { DocumentCanvas, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel }) {
	const { canvasProps, toolbarProps } = useViewerBuildingBlocks({ documentModel });
	return (
		<>
			{toolbarProps && toolbarProps.readAloudIssue !== 'unsupported' && (
				<>
					<button
						disabled={!toolbarProps.canReadAloud}
						aria-pressed={toolbarProps.readingAloud}
						onClick={() =>
							toolbarProps.readingAloud
								? toolbarProps.onStopReadAloud()
								: toolbarProps.onReadAloud()
						}
					>
						{toolbarProps.readingAloud ? 'Stop reading' : 'Read aloud'}
					</button>
					{toolbarProps.readAloudIssue === 'empty' && <span role="status">Nothing to read.</span>}
				</>
			)}
			<DocumentCanvas {...canvasProps} />
		</>
	);
}
```

Die Wiedergabe endet außerdem, wenn das Dokument bearbeitet oder ersetzt wird und wenn der Editor entfernt
wird. Die optionale `Toolbar` zeichnet mit denselben Feldern ihren eigenen Schalter „Read aloud“.
