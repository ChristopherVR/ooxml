# Element-API und Ereignisse

Jeder Framework-Adapter umhüllt dasselbe Custom Element `<docx-editor>`. Die Adapter leiten diese Eigenschaften weiter und geben diese Ereignisse in der Idiomatik ihres Frameworks wieder aus (siehe [Framework-Bindings](/bindings)).

## Attribute und Eigenschaften

| Eigenschaft      | Attribut          | Beschreibung                                                                                             |
| ---------------- | ----------------- | -------------------------------------------------------------------------------------------------------- |
| `documentModel`  |                   | Das angezeigte `DocumentModel`. Eine Zuweisung ist ein externer Austausch, kein Bearbeitungsereignis.    |
| `readOnly`       | `read-only`       | Blockiert lokale Änderungen. Stapel der Zusammenarbeit von anderen Teilnehmern werden weiter angewendet. |
| `locale`         | `locale`          | Oberflächensprache: `en`, `fr`, `de`, `es` oder `zh-CN`. Dokumentinhalte werden nie übersetzt.           |
| `fileName`       | `file-name`       | Name des geöffneten Dokuments (Standard `Document1.docx`).                                               |
| `reviewAuthor`   | `review-author`   | Autor, der bei nachverfolgten Änderungen und Kommentaren erfasst wird (Standard `Author`).               |
| `theme`          | `theme`           | `light`, `dark` oder `auto`. Siehe [Theming](/theming).                                                  |
| `themeColors`    |                   | Teilweise Token-Überschreibungen, angewendet als benutzerdefinierte Eigenschaften `--dve-*`.             |
| `showThumbnails` | `show-thumbnails` | Linke Leiste mit Seitenminiaturen (nur Drucklayout).                                                     |
| `showToolbar`    | `show-toolbar`    | `false` blendet das Menüband aus; Titelleiste und Statusleiste bleiben.                                  |
| `hiddenActions`  |                   | Auszublendende Menübandbefehle nach stabiler ID; siehe die Tabelle unter [Bindings](/bindings).          |
| `ribbonAddIns`   |                   | Register, die der Host nach den eingebauten hinzufügt; siehe [Bindings](/bindings#ribbon-add-in-tabs).   |
| `customization`  |                   | Capabilities, Host-Seiten im Datei-Register und Host-Menübefehle (siehe [Anpassung](#customization)).    |
| `dirty` (lesen)  |                   | True nach einer Änderung, false nach Speichern, Laden oder `markClean()`.                                |

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

| Methode                                               | Beschreibung                                                                                      |
| ----------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| `load(bytes)`                                         | Lädt DOCX- oder ältere DOC-Bytes. Bei einem Fehler wird abgelehnt und `document-error` ausgelöst. |
| `save()`                                              | Liefert einen `Blob` des serialisierten Dokuments.                                                |
| `saveBytes()`                                         | Liefert das rohe `Uint8Array`.                                                                    |
| `download(fileName?)`                                 | Speichert und startet einen Browser-Download.                                                     |
| `markClean()`                                         | Setzt das Dirty-Flag zurück, nachdem der Host das Dokument gespeichert hat.                       |
| `startCollaboration(config)`                          | Tritt einer Sitzung zur Zusammenarbeit bei; siehe [Zusammenarbeit](/de/collaboration).            |
| `startYjsCollaboration(session, options)`             | Tritt einem synchronisierten Yjs-Raum mit dem passenden Quellpaket bei.                           |
| `reconnectCollaboration()`, `resyncCollaboration()`   | Stellt die Yjs-Verbindung wieder her oder fordert einen Zustandsaustausch des Providers an.       |
| `getPendingCollaboration()`                           | Der ausstehende Schrittstapel oder `null`.                                                        |
| `receiveCollaboration(batch)`                         | Wendet einen von der Autorität akzeptierten Stapel an.                                            |
| `stopCollaboration(discardPending?)`                  | Verlässt die Sitzung; ausstehende Änderungen müssen bestätigt oder ausdrücklich verworfen werden. |
| `publishPresence`, `receivePresence`, `leavePresence` | Flüchtiger Präsenzkanal; siehe [Zusammenarbeit](/de/collaboration#präsenz).                       |

## Ereignisse

Alle Ereignisse sind bubbling und composed und überschreiten daher die Shadow-Grenze.

| Ereignis               | Detail                | Bedeutung                                                                                    |
| ---------------------- | --------------------- | -------------------------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`       | Das Dokument wurde durch Bearbeiten oder Überarbeitungsaktionen geändert.                    |
| `document-error`       | `Error`               | Laden, Speichern, Drucken oder eine Bearbeitung ist fehlgeschlagen.                          |
| `document-warning`     | `string`              | Nicht schwerwiegender Hinweis.                                                               |
| `readonly-change`      | `boolean`             | Der Benutzer hat in der Oberfläche zwischen Bearbeiten und Anzeigen umgeschaltet.            |
| `file-command`         | `{ command }`         | Abbrechbar. `command` ist `new`, `open`, `save`, `export`, `print` oder `share`.             |
| `dirty-change`         | `boolean`             | Der Zustand „ungespeicherte Änderungen“ hat sich umgekehrt.                                  |
| `page-change`          | `{ page, pageCount }` | Seite oder Seitenanzahl im Drucklayout hat sich geändert (näherungsweise Paginierung).       |
| `ribbon-action`        | `RibbonAction`        | Ein Menübandsteuerelement wurde aktiviert.                                                   |
| `ribbon-customize`     | `readonly string[]`   | Der Benutzer hat die angezeigten Menübandbefehle geändert; speichern Sie die IDs bei Bedarf. |
| `office-ribbon-add-in` | `{ tab, command }`    | Ein Befehl eines Host-Registers (`ribbonAddIns`) wurde gewählt.                              |
| `presence-send`        | `PresenceMessage`     | Die lokale Präsenz hat sich geändert; übertragen Sie sie an andere Clients.                  |
| `collaboration-send`   | `StepBatch`           | Lokale Schritte stehen aus; übertragen Sie sie an die Autorität.                             |

Das Paket exportiert die typisierte Ereigniszuordnung (`DocxEditorEventMap`) und die Laufzeitliste `DOCX_EDITOR_EVENTS`.
