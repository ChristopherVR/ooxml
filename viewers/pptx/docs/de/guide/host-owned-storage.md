# Vom Host verwalteter Speicher

Für einen Microsoft-Graph-Provider für vorhandene Dateien mit einer Demo ohne eigene Anmeldung siehe
[OneDrive-Integration](/de/guide/onedrive) und [OneDrive testen](/de/guide/onedrive-testing).

Jedes Binding nimmt eine Prop `storage` entgegen (Angular: Input, Vanilla: Option). Damit übergibt der Host die Datei-Ein- und -Ausgabe des Viewers an sich selbst. Es ist der Vertrag `DocumentStorage` aus `ooxml-core/embedding`, den auch die Editoren für Word, Excel und Visio akzeptieren. Er wird von `ooxml-ui/pptx` als Typ re-exportiert.

```ts
import type { DocumentStorage } from 'ooxml-ui/pptx';

const storage: DocumentStorage = {
	// Required: File > Save and Save As.
	async save(bytes, { fileName, mimeType, reason }) {
		await uploadToMyServer(fileName, bytes, mimeType);
		// Optional: report the name the document now has.
		return { fileName };
	},
	// Optional, all three together: AutoRecover copies.
	readRecovery: (id) => myRecovery.get(id),
	writeRecovery: (entry) => myRecovery.put(entry),
	clearRecovery: (id) => myRecovery.delete(id),
	// Optional: lets Version History list the kept copies.
	listRecovery: () => myRecovery.list(),
};
```

## Speichern und Speichern unter {#save-and-save-as}

Ohne `storage` lädt der Viewer die gespeicherte Datei herunter. Mit `storage` rufen Speichern und jedes Format von Speichern unter (`.pptx`, `.ppsx`, `.pptm`, `.ppt`) stattdessen `storage.save(bytes, meta)` auf:

- `meta.fileName` besteht aus dem Stamm des Quellnamens und der Endung des geschriebenen Formats (`savedPresentationFileName`). Daher wird `report.ppt` als `report.pptx` gespeichert.
- `meta.mimeType` richtet sich nach dem Format (`.ppt` hat den OLE2-Typ, `.ppsx` und `.pptm` haben jeweils eigene Typen).
- `meta.reason` ist der ausgeführte Befehl: `'save'` für File > Save (die Titelleiste, der Eintrag Speichern im Backstage, die mobile Werkzeugleiste und `downloadPptx()` bei Vanilla und Svelte) und `'saveAs'` für jeden Eintrag von Speichern unter, einschließlich Speichern unter > PowerPoint-Präsentation (.pptx).

Ein fehlgeschlagenes Speichern (das `save()` des Hosts wurde abgelehnt oder die Serialisierung ist gescheitert) wird in allen fünf Bindings gleich gemeldet. Das Ereignis `error` des Bindings (`onError`, `@error`, `(error)`, `onerror`) erhält `Save failed: <reason>`. Der Benutzer sieht die Meldung „Save failed“ im Benachrichtigungsstapel des Viewers, und die Kopie für AutoRecover bleibt erhalten. Sie wird nur verworfen, nachdem ein Speichern erfolgreich war und Optionen > Speichern sie anfordert. Die Speichermethoden werden aufgelöst und nicht abgelehnt.

Die Bytes sind dieselben, die `getContent()` zurückgibt. Eine passwortgeschützte Präsentation wird daher verschlüsselt übergeben.

## Wiederherstellungskopien {#recovery-copies}

Die Wiederherstellung per Autosave (die Props `autosave` und `autosaveIntervalMs`, indiziert nach `filePath`) schreibt normalerweise in die eigene IndexedDB-Datenbank der Bibliothek. Wenn `storage` `readRecovery`, `writeRecovery` und `clearRecovery` bereitstellt, verwendet jeder Wiederherstellungspfad stattdessen diese Methoden. Dazu gehören der Autosave-Writer, die Wiederherstellungsabfrage beim Laden (und deren Schaltfläche Verwerfen), die Bereinigung nach einem erfolgreichen Speichern, Versionsverlauf, File > Open > Recent (Datei > Öffnen > Zuletzt verwendet), Konto > Speicher und Datenschutz sowie `restoreSessionDeck(storage)`. Die Eintrags-ID ist der Schlüssel `filePath`, und `savedAt` ist ein Zeitstempel in Millisekunden. Ein Speicher, der nur einige der drei Methoden bereitstellt, wird für die Wiederherstellung ignoriert, und IndexedDB bleibt in Gebrauch.

Zum Auflisten (Zuletzt verwendet, Speichernutzung im Konto, Versionsverlauf) wird `listRecovery` benötigt. Ohne diese Methode sind die Listen leer. Die Aufräumfunktionen unter Optionen > Speichern („Delete cached files“, „clear cache on close“ und das Bereinigen nach der „cache retention“) entfernen die Kopien des Hosts nur, wenn `listRecovery` bereitgestellt wird. Ohne sie lässt der Viewer die Kopien des Hosts unberührt, und die Aufbewahrung ist Aufgabe des Hosts.

## Kein Ereignis `file-command` {#no-file-command-event}

Die Elemente für Word, Excel und Visio kündigen vor Speichern, Öffnen und Neu ein abbrechbares Ereignis `file-command` an. Die PowerPoint-Bindings haben kein solches Ereignis. Um das Speichern zu übernehmen, übergeben Sie `storage`. Um File > Open zu übernehmen, verwenden Sie `onOpenFile` (bei Vue und Angular: der Input `onOpenFile`).
