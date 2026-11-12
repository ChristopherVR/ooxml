# OneDrive-Integration

Für eine schrittweise Einrichtung (Registrierung bei Microsoft Entra, Anmeldung und ein vollständiges Beispiel) beginnen Sie mit dem [OneDrive-Schnellstart](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

## Datei > Öffnen > OneDrive {#file-open-onedrive}

Übergeben Sie einen Anbieter als `fileProviders`, dann zeigt Datei > Öffnen neben „Dieses Gerät“ einen OneDrive-Speicherort an:

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your sign-in owns the token
		signIn,
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

Der Speicherort listet zuletzt verwendete Dateien und Meine Dateien auf, durchsucht sie und öffnet einen Freigabelink. Eine ausgewählte Arbeitsmappe wird wie jede andere geladen, wird über Speichern an dasselbe Element zurückgespeichert (über Speichern unter als Kopie) und gilt erst als unverändert, nachdem der Upload abgeschlossen ist; ein fehlgeschlagener Upload lässt sie ungespeichert. Eine Datei, die `canWrite` ablehnt (anderes Laufwerk, geteilte Datei), öffnet sich schreibgeschützt. `fileProviders` ist eine Prop jedes Framework-Bindings (`fileProviders` in React, Vue, Solid, Svelte, Angular und Vanilla). Die Demo-App (`demos/xlsx`) übergibt standardmäßig den OneDrive-Anbieter der Demo; konfigurieren Sie ihn mit `VITE_ONEDRIVE_CLIENT_ID` in `demos/xlsx/demo-vanilla/.env.local` oder fügen Sie im Anmeldedialog ein Token aus dem Graph Explorer ein. Die untergeordnete Verdrahtung über `storage` weiter unten ist für Hosts gedacht, die ihre eigene Oberfläche zum Öffnen bauen.

Derselbe Speichervertrag unterstützt auch eigene Anbieter. Siehe [anbieterneutraler Speicher und Konfiguration der Zusammenführung](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md) für `createStorageSession`, Überschreibungen der Zusammenführung je Format und die Einstellungen für den Start der Suite. Verwenden Sie diese Erweiterungspunkte mit jedem Excel-Framework-Paket.

Excel nutzt die gemeinsame Verbindung `ooxml-core/embedding/onedrive`. Lösen Sie den Freigabelink im Host auf, prüfen Sie die berechtigungsspezifischen Rechte des Elements und übergeben Sie dem vorhandenen Editor die heruntergeladenen XLSX-Bytes:

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
await editor.load(opened.bytes, opened.file.name);
```

Das Element und die Framework-Handles stellen `saveBytes('xlsx')`, `load(bytes, fileName?)` und `markClean()` bereit. Ein eigenes Speichern des Hosts ruft `session.save(await editor.saveBytes('xlsx'))` auf; markieren Sie die Arbeitsmappe erst dann als unverändert, wenn das Speichern ohne neuere Änderungen erfolgreich war. Verwenden Sie den normalen Speicherweg des Editors, um eine Bearbeitung der aktiven Zelle zu übernehmen. CSV ist ein eigener Export und darf das verbundene XLSX-Element niemals ersetzen. Das native Speichern übergibt an `storage`, sofern der Host das abbrechbare Ereignis `file-command` nicht verhindert.

Derselbe Vertrag für `storage` und `readOnly` gilt für `@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer`, `xlsx-angular-viewer`, `xlsx-svelte-viewer`, `xlsx-solid-viewer` und `xlsx-vanilla-viewer` sowie für das schlichte Element `ooxml-ui/xlsx`. `@christophervr/xlsx-core` ist die Fassade des Modells und kein Authentifizierungsdienst.

Zur Authentifizierung, zur Freigabe und zur Wiederherstellung, die sowohl die lokale als auch die entfernte Version behält, siehe die [Referenz des Office-OneDrive-Pakets](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#excel-xlsx). Der [Testleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) erklärt Testfixtures ohne Zugangsdaten und die Abnahme mit einem Live-Konto; es wird weder eine Simulationsoberfläche noch ein Simulationsdienst ausgeliefert.

Registrieren Sie `mergeWorkbookSnapshots` aus `ooxml-core/xlsx/collab` als Callback `merge` der Sitzung, um unterstützte, voneinander unabhängige Änderungen an der Arbeitsmappe beim Speichern zusammenzuführen. Die Suite aktiviert dies automatisch; überlappende Zellbearbeitungen und nicht unterstützte Änderungen erfordern weiterhin eine Wiederherstellung. Dadurch werden die Sitzungen zur gemeinsamen Arbeit an Arbeitsmappen nicht mit der Mitbearbeitung in Microsoft Excel verbunden, externe Daten nicht aktualisiert und die [Funktionseinschränkungen](/de/features) nicht geändert.

## Konto > Verbundene Dienste {#account-connected-services}

Dieselben `fileProviders` füllen außerdem die Seite **Konto** der Backstage-Ansicht „Datei“ mit einer Liste „Verbundene Dienste“ (das gemeinsame `<office-ui-connected-services>`): eine Zeile je Anbieter mit dem angemeldeten Konto und den Schaltflächen Verbinden oder Trennen, die die eigenen Funktionen `signIn` und `signOut` des Anbieters aufrufen. Ein Anbieter ohne `signOut` zeigt an, dass der Host seine Sitzung verwaltet. Die Liste folgt dem `subscribe` der Anbieter, sodass eine Anmeldung über Datei > Öffnen sie aktualisiert, und ihre Texte stammen aus den Excel-Wörterbüchern in jeder Sprache. Ohne `fileProviders` gibt es keine Seite Konto.

## Eine Cloud-Datei aus der eigenen Oberfläche öffnen {#open-a-cloud-file-from-your-own-ui}

Datei > Öffnen ist erst erreichbar, wenn bereits ein Dokument geöffnet ist. Um eine Datei von einer Startseite oder aus Ihrer eigenen Auswahl zu öffnen, holen Sie sich ein `CloudOpenedFile` von einem beliebigen `CloudFileProvider` (der gemeinsame Browser `<office-ui-cloud-files>` sendet es im Ereignis `office-cloud-file-open`) und übergeben Sie es an `editor.openCloudFile(file)` auf `<xlsx-editor>` (das Handle jedes Bindings stellt es ebenfalls bereit). Dabei läuft derselbe Weg wie bei Datei > Öffnen > OneDrive: der Name der Datei, schreibgeschützt, wenn `canEdit` false ist, und Speichern wird an den eigenen Speicher der Datei geleitet, bis ein anderes Dokument geladen wird.
