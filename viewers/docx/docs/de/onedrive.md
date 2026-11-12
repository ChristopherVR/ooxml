# OneDrive-Integration

Eine Schritt-für-Schritt-Einrichtung (Registrierung bei Microsoft Entra, Anmeldung und ein vollständiges Beispiel) beginnt am besten mit dem [OneDrive-Schnellstart](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

Für andere Dienste oder eigene OneDrive-Abläufe verwenden Sie dieselbe Eigenschaft `storage` mit [`createStorageSession` und dem anbieterneutralen Integrationsleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md). Hosts können Authentifizierung, Übertragung, Merge-Handler, Konfliktdialog und Freigaberichtlinien ersetzen, ohne ein Word-Framework-Binding zu ändern.

Alle Word-Pakete verwenden die gemeinsame Verbindung `ooxml-core/embedding/onedrive`. Der Host löst den Freigabelink auf, stellt die Microsoft-Authentifizierung und die geprüfte Schreibberechtigung bereit und verbindet die entstandene Sitzung mit dem vorhandenen Speichervertrag des Editors.

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
editor.fileName = opened.file.name;
await editor.load(opened.bytes);
```

Das einfache `<docx-editor>` aus `ooxml-ui/docx` stellt `saveBytes()` bereit. Jedes Framework-Handle stellt `save()` bereit, das ein `Blob` zurückgibt, sowie `load(bytes)`, `markClean()` und `dirty`. Eine eigene Schaltfläche Speichern kann das Blob in Bytes umwandeln und `session.save(bytes)` aufrufen. Markieren Sie das Dokument nur dann als sauber, wenn die Speicherung gelungen ist und keine neuere Bearbeitung dazwischenkam. Das native Speichern läuft automatisch über `storage`, sofern der Host das abbrechbare Ereignis `file-command` nicht unterbindet.

## Aus OneDrive öffnen unter Datei > Öffnen {#open-from-onedrive-in-file-open}

Setzen Sie `fileProviders` (eine Eigenschaft des Elements und eine Prop jedes Bindings) auf eine Liste von `CloudFileProvider`-Objekten. Dann zeigt Datei > Öffnen Orte an: „Dieses Gerät“ (die Schaltfläche Durchsuchen) und je einen Ort pro Anbieter. Jeder Ort rendert den gemeinsamen Browser `<office-ui-cloud-files>` (Zuletzt verwendet, Meine Dateien, Suche, Öffnen über einen Freigabelink), gefiltert auf `.docx`. Seine Beschriftungen stammen aus den Word-Wörterbüchern in jeder Anzeigesprache.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDocxSnapshots } from 'ooxml-core/docx';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your Microsoft authentication
		signIn, // opens your sign-in UI and resolves with the account
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			docx: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

Eine an einem Ort gewählte Datei wird wie jedes andere Dokument geladen, übernimmt ihren Dateinamen und ist schreibgeschützt, wenn der Anbieter `canEdit: false` meldet (es gilt die strengere Einstellung von `readOnly` des Hosts und der Datei). Datei > Speichern und Speichern unter gehen auf den eigenen Speicher dieser Datei statt auf den `storage` des Hosts, bis ein anderes Dokument geladen wird. Die Titelleiste zeigt erst dann „gespeichert“ an, wenn der Upload abgeschlossen ist. Ein fehlgeschlagener Upload bleibt ungespeichert und löst `document-error` aus. Das Öffnen von einem Ort löst kein `file-command` aus (der Host hat die Anbieter gewählt). Speichern und Speichern unter tun das dagegen weiterhin; ein Host, der sie abbricht, behandelt sie selbst. Entfernt ein Host die Fähigkeit `open`, werden die Seite und die Orte mit ihr ausgeblendet.

Die Word-Demo übernimmt standardmäßig `createDemoOneDriveProvider` aus `demos/onedrive`. Melden Sie sich unter Datei > Öffnen > OneDrive an: Setzen Sie `VITE_ONEDRIVE_CLIENT_ID` (optional auch `VITE_ONEDRIVE_AUTHORITY`) in `demos/docx/demo-vanilla/.env.local`, um über MSAL Ihre eigene Entra-App-Registrierung zu verwenden. Alternativ geben Sie im Dialog der Demo eine Client-ID ein oder fügen im Feld „Access token“ ein Token aus dem Graph Explorer ein. Registrieren Sie den Ursprung der Demo (mit abschließendem Schrägstrich) als Weiterleitungs-URI für eine Single-Page-Anwendung. Die Demo bearbeitet nur Dateien auf dem eigenen Laufwerk der angemeldeten Person; andere Dateien öffnen sich schreibgeschützt.

Das gilt für `docx-react-viewer`, `docx-vue-viewer`, `docx-angular-viewer`, `docx-svelte-viewer`, `docx-solid-viewer` und `docx-vanilla-viewer`. Jedes leitet dieselben Props bzw. Optionen `storage` und `readOnly` weiter. `docx-core` ist eine Fassade des Modells und implementiert keine Microsoft-Anmeldung.

Die [Referenz zu Office und OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#word-docx) enthält die Paketübersicht, die Einrichtung der Authentifizierung, die Freigabebereiche und die umsetzbare Wiederherstellung bei Konflikten. Der [Testleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) trennt testlokale Übertragungen von echten Abnahmetests mit Microsoft. Es wird weder eine Simulationsoberfläche noch ein Dienst mitgeliefert.

Registrieren Sie `mergeDocxSnapshots` aus `ooxml-core/docx` als Callback `merge` der Sitzung, um unabhängige Text- und Eigenschaftsänderungen beim Speichern zusammenzuführen. Die Suite aktiviert dies automatisch. Mehrdeutige Strukturen oder unvereinbare Ersetzungen behalten beide Fassungen zur Wiederherstellung. Dies verbindet die Word-Kollaborationsräume nicht mit der gemeinsamen Bearbeitung von Microsoft Office. Die bestehenden [Erhaltungsgrenzen](/de/preservation) gelten weiterhin.

## Konto > Verbundene Dienste {#account-connected-services}

Dieselben `fileProviders` füllen außerdem die Seite **Konto** der Dateiansicht mit einer Liste „Verbundene Dienste“ (der gemeinsame `<office-ui-connected-services>`). Pro Anbieter gibt es eine Zeile mit dem angemeldeten Konto und den Schaltflächen Verbinden oder Trennen; sie rufen die eigenen `signIn` und `signOut` des Anbieters auf. Ein Anbieter ohne `signOut` zeigt an, dass der Host die Sitzung verwaltet. Die Liste folgt dem `subscribe` der Anbieter. Melden Sie sich also unter Datei > Öffnen an, wird sie aktualisiert, und ihre Texte stammen aus den Word-Wörterbüchern in jeder Sprache. Ohne `fileProviders` gibt es keine Seite Konto.

## Eine Cloud-Datei aus Ihrer eigenen Oberfläche öffnen {#open-a-cloud-file-from-your-own-ui}

Datei > Öffnen ist erst erreichbar, wenn ein Dokument geöffnet ist. Um eine Datei aus einer Startseite oder Ihrer eigenen Auswahl zu öffnen, holen Sie sich eine `CloudOpenedFile` von einem beliebigen `CloudFileProvider` (der gemeinsame Browser `<office-ui-cloud-files>` sendet sie im Ereignis `office-cloud-file-open`) und übergeben sie an `editor.openCloudFile(file)` auf `<docx-editor>` (jedes Binding-Handle stellt sie ebenfalls bereit). Das folgt demselben Weg wie Datei > Öffnen > OneDrive: der Dateiname, Schreibschutz, wenn `canEdit` false ist, und Speichern, das bis zum Laden eines anderen Dokuments an den eigenen Speicher der Datei geht.
