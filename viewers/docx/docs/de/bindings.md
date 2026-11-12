# Framework-Bindings

Alle Bindings binden `<docx-editor>` über dieselbe Funktion `mountEditor` ein. Eine Zuweisung einer Eigenschaft ist ein externer Austausch des Dokuments und kein Bearbeitungsereignis. `document-change` liefert ein `DocumentModel` direkt, `document-error` liefert einen `Error`. Beide Ereignisse werden über die Shadow-Grenze hinweg weitergereicht.

React:

```tsx
import { WordEditor } from 'docx-react-viewer';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` stellt `element`, `load(bytes)`, `save()`, `download()`, `markClean()`, `openCloudFile()` und `dirty` bereit.

Vue:

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular:

```ts
import { WordEditorComponent } from 'docx-angular-viewer';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5:

```svelte
<script>
  import WordEditor from 'docx-svelte-viewer';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

SolidJS:

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from 'docx-solid-viewer';

function DocumentEditor() {
	const [model, setModel] = createSignal(initialDocument);
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

Solid verwendet dieselben Optionen wie der gemeinsame Adapter: `documentModel`, `readOnly`, `locale`, `onDocumentChange` und `onDocumentError`. Über `editorRef` erhalten Sie die Methoden `element`, `load` und `save` des eingebundenen Editors. Der Besitzer der Solid-Komponente räumt den Editor und seine Listener auf, sobald die Komponente entfernt wird.

Custom Element:

```ts
import { registerDocxEditor } from 'docx-vanilla-viewer';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

Der Vanilla-Mount-Adapter wird aus `docx-vanilla-viewer` exportiert und liefert die Operationen update, destroy, load, save und element. Er verhindert Rücksetzungen durch Rückkopplung, wenn ein übergeordnetes Element das ausgegebene Modell zurückgibt. Für externe Modelländerungen übergeben Sie ein neues Objekt. Ein Modell direkt zu verändern wird als Reaktivitätsmechanismus nicht unterstützt.

Framework-Pakete verwenden ihre eigene native Lebenszyklus- und Ref-API. Keine Framework-Vorlage enthält eine Werkzeugleiste, einen Seitenrenderer oder eine Implementierung von Dokumentbefehlen. SSR-Importe registrieren keine Elemente; das Einbinden ist eine clientseitige Operation. Der aktuelle Browservertrag prüft die installierten Versionen, nicht jede frühere Peer-Version.

Setzen Sie die optionale Eigenschaft `locale` an jedem Framework-Editor, um die Multifunktionsleiste und die Suchsteuerelemente zu lokalisieren (`en`, `fr`, `de`, `es` oder `zh-CN`, zum Beispiel `locale="de"`). Das betrifft nur die Oberflächentexte und übersetzt keine Dokumentinhalte.

## Fensterrahmen und Dateibefehle {#window-chrome-and-file-commands}

Der Editor zeichnet im Shadow Root einen Fensterrahmen im Stil von Word: eine Titelleiste (Speichern, Rückgängig und Wiederholen in der Schnellzugriffsleiste, der Dateiname und der Speicherstatus sowie eine Befehlssuche mit dem Text "Tell me what you want to do"), den Wechsel zwischen Bearbeiten und Anzeigen, **Kommentare** und **Freigeben** am rechten Ende der Registerzeile der Multifunktionsleiste (wie in Word 365), einen Reiter **Datei** mit einer Backstage-Ansicht (Informationen mit Hinweisen zur Kompatibilität, Neu, Öffnen, Speichern, Kopie als DOCX speichern, Drucken) und eine Statusleiste (Seite, Wörter, Kompatibilitätshinweise, Web- bzw. Drucklayout und Zoom).

Setzen Sie `element.fileName`, um das geöffnete Dokument zu benennen. Dateibefehle werden zuerst als abbrechbares Ereignis `file-command` angekündigt, dessen `detail.command` den Wert `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` oder `share` hat. Rufen Sie `preventDefault()` auf, um den Befehl in Ihrer Anwendung zu behandeln, zum Beispiel um in Ihren eigenen Speicher zu schreiben. Andernfalls verwendet der Editor seine nur im Browser verfügbare Standardlösung: einen Dateidialog, einen Download oder den Druckdialog. `share` (die Schaltfläche Freigeben) öffnet den Freigabedialog, den alle Office-Editoren verwenden: Sie starten eine Sitzung (Sitzungsname, Ihr Name und optional ein Server) oder treten einer Sitzung über einen Einladungslink oder den Sitzungsnamen bei. Danach können Sie den Einladungslink kopieren, sehen, wer in der Sitzung ist, und das Freigeben beenden. Dadurch wird die Eigenschaft `collaboration` gesetzt. Brechen Sie den Vorgang ab, um auf eigene Weise zu teilen. Wählen Sie **Anzeigen**, löst das `readonly-change` aus.

```ts
editor.fileName = 'Quarterly report.docx';
editor.addEventListener('file-command', (event) => {
	const { command } = (event as CustomEvent<{ command: string }>).detail;
	if (command === 'save') {
		event.preventDefault();
		void editor.save().then(uploadToMyStorage);
	}
});
```

## Speichern, Status „geändert“ und UI-Optionen {#saving-dirty-state-and-ui-options}

`save()` liefert ein `Blob` des serialisierten Dokuments (`saveBytes()` liefert das rohe `Uint8Array`). `download(fileName?)` speichert und startet einen Download im Browser. `dirty` und `dirty-change` melden nicht gespeicherte Änderungen. `dirty` wird bei `document-change` auf true gesetzt und nach Datei > Speichern, `download()`, dem Laden eines Dokuments oder `markClean()` auf false. `save()` allein hebt die Markierung nicht auf, weil nur der Host weiß, ob das Blob dauerhaft gespeichert wurde. Jeder Adapter stellt auf seinem Handle `save`, `download`, `markClean` und `dirty` bereit sowie `onDirtyChange` und `onPageChange` (oder das ereignisbasierte Gegenstück des Frameworks).

| Eigenschaft (Attribut)               | Standard | Wirkung                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| ------------------------------------ | -------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `showThumbnails` (`show-thumbnails`) | `false`  | Linke Leiste mit Seitenminiaturen. Miniaturen brauchen das Drucklayout; in der Entwurfsansicht weist die Leiste darauf hin, statt Seiten aufzulisten.                                                                                                                                                                                                                                                                                           |
| `showToolbar` (`show-toolbar`)       | `true`   | `show-toolbar="false"` blendet die Multifunktionsleiste aus; Titelleiste und Statusleiste bleiben.                                                                                                                                                                                                                                                                                                                                              |
| `hiddenActions`                      | `[]`     | Steuerelemente der Multifunktionsleiste, die über ihre stabile Kennung in Kebab-Case ausgeblendet werden (`RIBBON_ACTION_IDS`, zum Beispiel `'bold'`, `'insert-table'`, `'track-changes'`, `'paste'`, `'word-count'`). Leere Gruppen und Reiter werden ebenfalls ausgeblendet. Eine Farbschaltfläche und ihr Dropdown-Pfeil werden gemeinsam ausgeblendet. Unbekannte Kennungen werden ignoriert. Die Kennungen ändern sich nicht mit `locale`. |
| `ribbonAddIns`                       | keine    | Reiter, die der Host nach den integrierten Reitern hinzufügt; siehe [Reiter für Add-ins der Multifunktionsleiste](#ribbon-add-in-tabs).                                                                                                                                                                                                                                                                                                         |

**Veraltet:** Bevor es Kennungen gab, nahm `hiddenActions` die englische Beschriftung des Steuerelements entgegen (`'Bold'`, `'Insert table'`). Diese `LegacyRibbonLabel`-Zeichenketten werden noch für eine weitere Version akzeptiert: Sie werden in Kennungen umgewandelt, und das Element sendet pro Zuweisung ein `document-warning`. Das Lesen von `hiddenActions` liefert immer Kennungen. `ribbon-action`-Ereignisse haben nie Beschriftungen transportiert (ihr Detail wird über den Befehlstyp geschlüsselt), daher ändern sie sich nicht.

`page-change` (`{ page, pageCount }`) wird nur im Drucklayout ausgelöst. Die Seitenumbrüche stammen aus der eigenen Layout-Engine dieses Editors und sind eine Näherung, nicht der Seitenumbruch von Word. Die Statusleiste und die QuickInfos der Leiste weisen darauf hin.

## Reiter für Add-ins der Multifunktionsleiste {#ribbon-add-in-tabs}

Ein Host kann eigene Reiter der Multifunktionsleiste nach denen von Word hinzufügen, so wie es ein Office-Add-in tut. `ribbonAddIns` nimmt einfache Daten entgegen und ist eine Eigenschaft jedes Bindings. Ein ausgewählter Befehl führt seinen Callback `run` aus und wird über `onRibbonAddIn` gemeldet (Vue `office-ribbon-add-in`, Angular `ribbonAddIn`, Svelte `onribbonaddin`) mit `{ tab, command }`.

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

<WordEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

`size: 'small'`-Befehle füllen Spalten zu dritt, und `items` erzeugt ein Dropdown-Menü. Wenn Sie einen gleichen Beschreibungsblock erneut übergeben, wie es ein Framework bei jedem Render tut, bleiben die Bereiche erhalten und nur die Callbacks werden ersetzt. Lassen Sie die Eigenschaft weg, werden die Reiter entfernt. Ein Reiter darf nicht die Kennung eines integrierten Reiters verwenden, und seine Befehle erreichen nie die eigene Befehlsbehandlung des Editors. Derselbe Beschreibungsblock (`RibbonAddInTab`, aus `ooxml-core/ribbon`, von jedem Framework-Paket wieder exportiert) funktioniert in den Editoren für Word, Excel, PowerPoint und Visio. Der Editor enthält keinen Add-in-Reiter.

## Binding-Eigenschaften und -Ereignisse {#binding-props-and-events}

Jeder Framework-Adapter leitet dieselben Elementeigenschaften weiter und sendet dieselben Ereignisse erneut aus. Deshalb bedeuten die folgenden Namen in jedem Binding dasselbe. Die Ereignisnamen sind die des Elements (`file-command`); jedes Framework schreibt sie in seiner eigenen Form.

| Elementeigenschaft | React / Solid / Vanilla | Vue-Eigenschaft | Angular-Input  | Svelte-Eigenschaft |
| ------------------ | ----------------------- | --------------- | -------------- | ------------------ |
| `fileName`         | `fileName`              | `fileName`      | `fileName`     | `fileName`         |
| `reviewAuthor`     | `reviewAuthor`          | `reviewAuthor`  | `reviewAuthor` | `reviewAuthor`     |
| `themeColors`      | `themeColors`           | `themeColors`   | `themeColors`  | `themeColors`      |

`fileName`, `reviewAuthor` und `themeColors` werden weitergeleitet, sobald die Eigenschaft gesetzt wird oder sich ändert. So wird eine über Datei > Öffnen vorgenommene Umbenennung nicht durch ein erneutes Rendern mit dem alten Namen zurückgesetzt.

| Elementereignis    | React / Solid / Vanilla-Option | Vue-Ereignis       | Angular-Output    | Svelte-Callback     |
| ------------------ | ------------------------------ | ------------------ | ----------------- | ------------------- |
| `file-command`     | `onFileCommand(event)`         | `file-command`     | `fileCommand`     | `onfilecommand`     |
| `document-warning` | `onDocumentWarning(message)`   | `document-warning` | `documentWarning` | `ondocumentwarning` |
| `readonly-change`  | `onReadOnlyChange(readOnly)`   | `readonly-change`  | `readOnlyChange`  | `onreadonlychange`  |
| `ribbon-action`    | `onRibbonAction(detail)`       | `ribbon-action`    | `ribbonAction`    | `onribbonaction`    |
| `ribbon-customize` | `onRibbonCustomize(ids)`       | `ribbon-customize` | `ribbonCustomize` | `onribboncustomize` |

`file-command` ist das einzige abbrechbare Ereignis. Deshalb erhält sein Handler das Ereignis selbst und nicht dessen Detail. Rufen Sie `event.preventDefault()` auf, um einen Dateibefehl zu übernehmen (Speichern, Öffnen, Neu und so weiter). `event.detail.command` gibt an, welcher Befehl es ist. Die Handler von Vue, Angular und Svelte erhalten dasselbe `CustomEvent`-Objekt. Die anderen Ereignisse übergeben ihr Detail direkt.
