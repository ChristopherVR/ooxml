# OneDrive-Integration

Eine Schritt-für-Schritt-Einrichtung (Registrierung in Microsoft Entra, Anmeldung und ein vollständiges Beispiel) finden Sie zuerst im [Schnellstart für OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

Für einen anderen Dienst oder eine eigene OneDrive-Implementierung siehe [anbieterneutrale Speicherintegrationen](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md). `createStorageSession` verbindet sich mit derselben Eigenschaft `storage` in allen Visio-Paketen, mit vom Host verwalteten Revisionsprüfungen, Merge-Handlern, Konfliktoberflächen und Freigaberichtlinien.

## Datei > Öffnen mit Cloud-Speicherorten {#file-open-with-cloud-places}

Übergeben Sie `fileProviders` (eine Eigenschaft des Elements und eine Prop jedes Bindings), um Cloud-Speicherorte zu Datei > Öffnen hinzuzufügen. Ist die Liste nicht leer, zeigt „Öffnen“ **Dieses Gerät** (die übliche Kachel „Durchsuchen“) und einen Speicherort je Anbieter. Dieser rendert für `.vsdx`-Zeichnungen den gemeinsamen Browser `<office-ui-cloud-files>`: anmelden, Zuletzt verwendet, Meine Dateien, Suche und Öffnen über einen Freigabelink. Die Texte stammen in jeder Sprache aus der Anzeigesprache von Visio.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDrawingSnapshots } from 'ooxml-core/visio/collab';

const oneDrive = createOneDriveFileProvider({
	getAccessToken, // the host's own sign-in
	signIn, // opens the host's sign-in UI and resolves with the account
	canWrite: async (file) => file.driveId === (await myDriveId()),
	mergers: {
		vsdx: ({ base, local, remote }) => mergeDrawingSnapshots(base.bytes, local, remote.bytes),
	},
});
// <VisioViewer fileProviders={[oneDrive]} /> (React); [fileProviders] in Angular; :file-providers in Vue.
```

Eine dort gewählte Zeichnung wird über den normalen Pfad für `.vsdx` geladen und übernimmt den Namen der Datei. Danach gilt:

- Sie ist schreibgeschützt, wenn es die Datei ist (`canEdit` ist false). Hat der Host den Viewer bereits schreibgeschützt gemacht (`canEdit`, `readOnly`, eine Rolle `viewer` der Zusammenarbeit oder `capabilities.edit`), bleibt sie für eine bearbeitbare Datei schreibgeschützt: Die strengere der beiden Einstellungen gewinnt.
- Datei > Speichern und Speichern unter verwenden den eigenen Speicher dieser Datei statt des Host-`storage`, bis eine andere Zeichnung geladen wird (Öffnen, Neu, `load` oder `document`). Speichern wird erst beendet, wenn der Anbieter die Bytes abgelegt hat: Die Zeichnung gilt dann als sauber, und sie bleibt mit einem `document-error` ungespeichert, wenn der Upload fehlschlägt.
- Das abbrechbare Ereignis `file-command` wird weiterhin zuerst ausgelöst (`open`, `save`); das Abbrechen von `open` lässt die aktuelle Zeichnung unverändert. Das Backstage schließt sich, sobald die Zeichnung geöffnet ist.
- Diagnosen des Exports werden wie bei jedem anderen Speichern gemeldet. Das alte `.vsd` ist kein Ziel für die Bearbeitung in der Cloud.

Die Demos übergeben standardmäßig einen OneDrive-Speicherort, angemeldet über den nur für die Demo gedachten Dialog (eine Client-ID oder ein im Feld **Zugriffstoken** eingefügtes Token aus dem Graph Explorer). Dieser Dialog ist Demo-Konfiguration, kein Produkt-UI; siehe `demos/onedrive`. Um den Dialog zu überspringen, setzen Sie `VITE_ONEDRIVE_CLIENT_ID` (und optional `VITE_ONEDRIVE_AUTHORITY`) in `demos/visio/.env.local`, führen `bun run dev` in `viewers/visio` aus und registrieren die ausgegebene Adresse (mit abschließendem Schrägstrich) als Redirect-URI einer Single-Page-Anwendung.

## Konto > Verbundene Dienste {#account-connected-services}

Dieselben `fileProviders` füllen außerdem die Seite **Konto** der Datei-Ansicht mit einer Liste „Verbundene Dienste“ (der gemeinsame `<office-ui-connected-services>`): eine Zeile je Anbieter mit dem angemeldeten Konto und den Schaltflächen Verbinden oder Trennen, die die eigenen Funktionen `signIn` und `signOut` des Anbieters aufrufen. Ein Anbieter ohne `signOut` zeigt an, dass der Host die Sitzung verwaltet. Die Liste folgt dem `subscribe` der Anbieter, sodass eine Anmeldung über Datei > Öffnen sie aktualisiert, und ihre Texte stammen in jeder Sprache aus den Visio-Wörterbüchern. Die Seite behält außerdem das lokale Profil und die Produktinformationen. Wie in Word und Excel erscheint die Seite „Konto“ nur, wenn es `fileProviders` gibt; ohne sie bleibt das Profil unter **Datei > Optionen** bearbeitbar.

## Eine Sitzung direkt verwenden {#using-a-session-directly}

Visio verwendet für native VSDX-Dateien die gemeinsame Verbindung `ooxml-core/embedding/onedrive`. Der Host verwaltet die Microsoft-Anmeldung, die Auflösung von Links und die Prüfung der Berechtigungen für das jeweilige Element. Übergeben Sie die heruntergeladenen Bytes an den vorhandenen Viewer, mit einer Einschränkung des Hosts auf Schreibschutz, wenn der Zugriff unbekannt ist:

```ts
// Set these through the element or binding's existing props/options.
const canEdit = opened.file.access === 'write';
const options = { storage: session.storage, canEdit, readOnly: !canEdit };
// After mounting with options:
await viewer.load(opened.bytes);
```

`exportVsdx()` liefert ein Objekt mit den nativen `bytes` und den Diagnosen des Schreibvorgangs. Eine eigene Schaltfläche „Speichern“ kann `await session.save((await viewer.exportVsdx()).bytes)` aufrufen. Das native Speichern von Datei verwendet denselben Speichervertrag, sofern der Host das abbrechbare Ereignis `file-command` nicht verhindert. Prüfen Sie die Diagnosen, statt verlustfreien Export zuzusagen. Binäre `.vsd`-Vorschauen sind keine nativen Ziele der Bearbeitung in OneDrive.

Dies gilt für das Element `ooxml-ui/visio` sowie für `visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`, `visio-svelte-viewer`, `visio-solid-viewer` und `visio-vanilla-viewer`. `visio-core` ist die Modell-Fassade. Die Verbindung fügt ihr keine Microsoft-Authentifizierung hinzu.

Die Referenz zu OneDrive für Office enthält die Paketverdrahtung, die Authentifizierung, die berechtigungsbewusste Freigabe und die Konfliktwiederherstellung, die beide Versionen erhält: [OneDrive-Integration für Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#visio-vsdx). Für Tests mit lokalen Transporten und Abnahmetests mit dem Live-Dienst siehe [Tests](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md). Es werden weder eine Simulationsoberfläche noch ein Dienst ausgeliefert.

Registrieren Sie `mergeDrawingSnapshots` aus `ooxml-core/visio/collab` als Callback `merge` der Sitzung, um beim Speichern unabhängige Änderungen an Formen und Eigenschaften zusammenzuführen, sofern sie unterstützt werden. Die Suite macht dies automatisch. Überlappende oder mehrdeutige Änderungen erfordern weiterhin eine Wiederherstellung. Dies ist keine gleichzeitige Bearbeitung durch Microsoft und keine Dateisperre. Die vorhandenen [Grenzen von Visio](/parity) und die Diagnosen des Exports gelten weiterhin.

## Eine Cloud-Datei aus der eigenen Oberfläche öffnen {#open-a-cloud-file-from-your-own-ui}

Datei > Öffnen ist erst erreichbar, wenn ein Dokument geöffnet ist. Um eine Datei aus einer Startseite oder einer eigenen Auswahl zu öffnen, erhalten Sie ein `CloudOpenedFile` von einem beliebigen `CloudFileProvider` (der gemeinsame Browser `<office-ui-cloud-files>` sendet es als `office-cloud-file-open`) und übergeben es an `viewer.openCloudFile(file)` auf `<visio-viewer>` (das Handle jedes Bindings stellt es ebenfalls bereit). Dies läuft denselben Pfad wie Datei > Öffnen > OneDrive: der Name der Datei, Schreibschutz, wenn `canEdit` false ist, und Speichern über den eigenen Speicher der Datei, bis ein anderes Dokument geladen wird.
