---
title: OneDrive-Integration
description: Verbinden Sie jedes PowerPoint-Viewer-Binding mit dem gemeinsamen Office-OneDrive-Speicher, der Freigabe und der Wiederherstellung bei Konflikten.
---

# OneDrive-Integration

Jedes PowerPoint-Framework-Paket akzeptiert außerdem einen eigenen Speicher-Provider über die bestehende Prop `storage`. Siehe [Provider-neutrale Integration und Konfiguration](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md) für `createStorageSession`, Überschreibungen von Authentifizierung und Transport, Merge-Handler je Format, Konfliktdialoge des Hosts und ausdrückliche Freigabeoptionen.

PowerPoint verwendet dieselbe optionale Verbindung `ooxml-core/embedding/onedrive` wie Word, Excel und Visio. Der Host stellt die Microsoft-Authentifizierung und die geprüften Berechtigungen für Elemente bereit. Es gibt keine mitgelieferte Microsoft-App-Registrierung, keinen Simulationsdienst und keine simulierte OneDrive-Demo.

Eine Schritt-für-Schritt-Einrichtung mit Screenshots (App-Registrierung, Anmeldung, ein Feld für den Freigabelink und ein vollständiges React-Beispiel) finden Sie im [Schnellstart für OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md). Jede Framework-Demo bietet OneDrive unter File > Open (Datei > Öffnen) an, damit Sie es mit Ihrem eigenen OneDrive ausprobieren können. Der Abschnitt „Try it in the PowerPoint demos“ (In den PowerPoint-Demos ausprobieren) des Schnellstarts erklärt die Einrichtung.

## File > Open: Orte mit `fileProviders` {#file-open-places-fileproviders}

Übergeben Sie `fileProviders` (eine Liste von `CloudFileProvider`, zum Beispiel von `createOneDriveFileProvider` aus `ooxml-core/embedding/onedrive`). Dann listet die Seite Öffnen **This device** (Dieses Gerät) sowie einen Ort je Provider. Die Prop heißt in jedem Binding gleich: `fileProviders` (React, Angular-Input, Svelte, Vanilla-Option) und `:file-providers` (Vue).

```tsx
<PowerPointViewer content={bytes} canEdit fileProviders={[oneDriveProvider]} />
```

Um Änderungen zusammenzuführen, die jemand anderes inzwischen gespeichert hat, geben Sie dem Provider einen PowerPoint-Merger. Ohne einen Merger öffnet eine neuere OneDrive-Version immer die Auswahl für Konflikte. Mit `merge: false` wird für jedes Format ebenso nachgefragt. `maxMergeAttempts` (Standard 3, wenn ein Merge greift) begrenzt die Wiederholungen:

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergePptxSnapshots } from 'ooxml-core/pptx';

const oneDriveProvider = createOneDriveFileProvider({
	getAccessToken,
	canWrite,
	signIn,
	mergers: {
		pptx: ({ base, local, remote }) => mergePptxSnapshots(base.bytes, local, remote.bytes),
	},
});
```

Eine Datei, die von einem Ort geöffnet wird, lädt wie jedes andere Öffnen (der Bearbeitungsverlauf wird zurückgesetzt, `onLoad` wird ausgelöst). Anschließend verwendet der Viewer den eigenen Speicher dieser Datei für Datei > Speichern, Speichern unter und die Wiederherstellungskopien, nicht die Prop `storage`. Er öffnet die Datei schreibgeschützt, wenn der Dienst keine Bearbeitung erlaubt. Ein Host mit `canEdit={false}` oder `readOnly` bleibt schreibgeschützt. Dieser Zustand bleibt bestehen, bis ein anderes Dokument geladen wird: eine neue Präsentation, ein Öffnen von diesem Gerät oder aus Zuletzt verwendet oder neuer `content` vom Host. Schlägt ein Speichern fehl, erscheint die Meldung über das fehlgeschlagene Speichern, und die Statusleiste zeigt weiterhin **Unsaved changes** (Ungespeicherte Änderungen).

## Konto > Verbundene Dienste {#account-connected-services}

Dieselbe Prop `fileProviders` füllt außerdem einen Abschnitt **Connected services** (Verbundene Dienste) unter **File > Account** (Datei > Konto), und zwar unterhalb der Karten für Profil, Speicher und Anmeldung. Jeder Provider ist eine Zeile mit seinem Konto (Name und E-Mail) und der Schaltfläche **Connect** (Verbinden) bzw. **Disconnect** (Trennen). Diese rufen die eigenen Methoden `signIn` und `signOut` des Providers auf. Ein Provider ohne `signOut` zeigt an, dass seine Sitzung von Ihrer Organisation verwaltet wird. Die Liste folgt dem `subscribe` des Providers, daher aktualisiert eine Anmeldung über File > Open die Anzeige. Ohne `fileProviders` erscheint nichts Zusätzliches, und die bestehende Anmeldekarte `accountAuth` bleibt unverändert. Alle fünf Bindings rendern das gemeinsame Element `office-ui-connected-services` (als `pptx-ui-connected-services`) mit übersetzten Beschriftungen.

Der [vollständige Integrationsleitfaden für Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) beschreibt Authentifizierung, Sitzungsspeicher, sichere Cloud-Kopien, Rollen und Geltungsbereiche der Freigabe, Wiederherstellung bei Konflikten und die Referenz für jeden Pakettyp.

```ts
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';

const drive = createOneDriveClient({
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
});
const opened = await drive.open(sharingLink);
const session = createOneDriveSession(drive, opened);
// Pass session.storage and opened.file.access === 'write' to the viewer.
```

Die Sitzung und die Berechtigungsdienste in diesem Beispiel sind Host-Integrationen, keine mitgelieferten Objekte. Binden Sie den Adapter als direkte Abhängigkeit von `ooxml-core` ein. `ooxml-ui/pptx/onedrive` bleibt nur als Kompatibilitätsschnittstelle für den früheren, PowerPoint-spezifischen Import bestehen.

## Alle fünf Bindings {#all-five-bindings}

| Binding/Paket                  | Bytes öffnen                          | Bearbeitung einschränken      | Aktuelle PPTX exportieren         |
| ------------------------------ | ------------------------------------- | ----------------------------- | --------------------------------- |
| React, `pptx-react-viewer`     | `content={opened.bytes}`              | `canEdit={canEdit}`           | `await ref.current.getContent()`  |
| Vue, `pptx-vue-viewer`         | `:content="opened.bytes"`             | `:can-edit="canEdit"`         | `await viewer.value.getContent()` |
| Angular, `pptx-angular-viewer` | `[content]="opened.bytes"`            | `[canEdit]="canEdit"`         | `await viewer.getContent()`       |
| Svelte, `pptx-svelte-viewer`   | `source={opened.bytes}`               | `editable={canEdit}`          | `await viewer.getContent()`       |
| Vanilla, `pptx-vanilla-viewer` | `await viewer.loadFile(opened.bytes)` | `viewer.setEditable(canEdit)` | `await viewer.save()`             |

`canEdit` (bei Svelte `editable`) bedeutet `opened.file.access === 'write'`. Ein unbekannter Zugriff bleibt schreibgeschützt. Jedes Binding akzeptiert `storage` als bestehende Prop, Input bzw. Option. Weisen Sie `session.storage` zu, um das native Speichern mit OneDrive zu verbinden. `onOpenFile` kann den Auswahldialog für Freigabelinks des Hosts öffnen. Übergeben Sie die Freigabeseite von Microsoft nicht an einen generischen Lader für Dokumentquellen.

Der Viewer markiert das Dokument erst als gespeichert, nachdem `storage.save` aufgelöst wurde, und nur dann, wenn während des Hochladens nichts bearbeitet wurde. Ein abgelehntes Speichern lässt das Dokument als ungespeichert. Für eine vom Host verwaltete Schaltfläche Speichern exportieren Sie die aktuelle PPTX und rufen `session.save(bytes)` auf. Lassen Sie den Editor eingehängt und bewahren Sie den Bearbeitungsverlauf. Verfolgen Sie Änderungen während des Hochladens getrennt vom hochgeladenen Snapshot. Betrachten Sie einen nativen Speichervorgang, der aufgelöst wird, nicht als Beleg für einen Erfolg. Einige Methoden melden Fehler über `onError` und den Benachrichtigungsstapel. Siehe [vom Host verwalteter Speicher](/de/guide/host-owned-storage).

## Konflikte und Freigabe {#conflicts-and-sharing}

Eine Änderung, die nur die Revision betrifft, kann einmal wiederholt werden, sofern die heruntergeladenen Cloud-Bytes weiterhin der geladenen Ausgangsbasis entsprechen. Ein echter inhaltlicher Konflikt bewahrt Ausgangsbasis, lokale Version und entfernte Version in `OneDriveConflictError`. Bieten Sie die Optionen „Save my edits as a OneDrive copy“ (Meine Änderungen als OneDrive-Kopie speichern), „Keep local copy and open latest“ (Lokale Kopie behalten und neueste öffnen), „Download my edits“ (Meine Änderungen herunterladen) und „Keep editing“ (Weiter bearbeiten) an. Überschreiben Sie nichts stillschweigend und laden Sie nicht neu. Registrieren Sie `mergePptxSnapshots` als Callback `merge` der Sitzung, um unabhängige Folienbearbeitungen automatisch zusammenzuführen, soweit sie unterstützt werden. Die Suite macht das standardmäßig. Nicht unterstützte Überschneidungen erfordern weiterhin eine Wiederherstellung. Die Sitzung sperrt die Datei nicht bis zum letzten Hochladen.

Speichern unter in eine kompatible native PPTX erzeugt eine eigene Cloud-Kopie über den Speicher der Sitzung. Exporte nach PPT, PPTM, PPSX oder in ein anderes Format dürfen das Original-PPTX nicht ersetzen. Behandeln Sie sie als ausdrücklich getrennten Exportablauf. Eine Aktion Teilen erstellt erst dann einen Link, wenn der Benutzer eine Rolle (Anzeigen oder Bearbeiten) und einen Geltungsbereich gewählt hat. Die Freigabe eines OneDrive-Links verbindet einen PowerPoint-Zusammenarbeitsraum nicht mit der gleichzeitigen Bearbeitung in Microsoft PowerPoint.

Siehe [Integrationstests](/de/guide/onedrive-testing) und die bestehenden [Formateinschränkungen](/de/guide/limitations).
