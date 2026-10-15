# Element-API und Ereignisse

Jedes Binding umhüllt dasselbe Custom Element `<teams-app>`. Die Bindings leiten diese Eigenschaften weiter und geben diese Ereignisse im Idiom ihres Frameworks wieder aus (siehe die [Framework-Anleitungen](/frameworks/react)).

## Attribute und Eigenschaften

| Eigenschaft     | Attribut        | Beschreibung                                                                                                                         |
| --------------- | --------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `workspaceId`   | `workspace-id`  | Der gemeinsame Raum: 1 bis 128 Zeichen aus Buchstaben, Ziffern, `-` und `_`.                                                         |
| `userName`      | `user-name`     | Anzeigename.                                                                                                                         |
| `userId`        | `user-id`       | Stabile ID; standardmäßig eine, die in diesem Browser gespeichert ist.                                                               |
| `config`        | `server-config` | `{ mode: 'local' \| 'server', syncUrl?, signalingUrl?, iceServers, iceTransportPolicy?, token? }`. Das Attribut nimmt JSON entgegen. |
| `uploadFile`    |                 | Wohin Anhänge gehen; standardmäßig der `/files`-Endpunkt des Servers.                                                                |
| `openers`       |                 | Handler je Office-Typ (`docx`, `xlsx`, `pptx`, `vsdx`).                                                                              |
| `embeds`        |                 | Callbacks je Dateityp, die die URL einer Sandbox-kompatiblen Viewer-Seite liefern; nützlich für PowerPoint.                          |
| `client` (read) |                 | Der Core-Client hinter dem Element (Zustand hinein, Aktionen hinaus).                                                                |

Lassen Sie `config` weg, dann bietet das Element einen Einstellungsdialog an (in diesem Browser gespeichert) oder läuft im lokalen Modus.

## Inhaltsvorschauen

Das Öffnen eines Anhangs zeigt standardmäßig eine Arbeitsbereichs-Vorschau. Word und Excel nutzen
die nativen Editor-Elemente, zunächst schreibgeschützt; Visio nutzt seinen Viewer. Markdown hat einen Renderer
für eine sichere Teilmenge; Textdateien zeigen den Text unverändert; HTML und Websites nutzen Sandbox-
Frames. `openers` und das abbrechbare Ereignis `teams-open-file` haben weiterhin Vorrang.

PowerPoint benötigt eine vom Host bereitgestellte Einbettungsseite:

```ts
const embeds = {
	pptx: ({ url }: { url: string | undefined }) =>
		`https://viewer.example.com/embed?source=${encodeURIComponent(url ?? '')}`,
};
```

Übergeben Sie `embeds` an ein beliebiges Komponenten-Binding. Die Seite muss die Datei-URL verstehen und
eine Sandbox mit Skripten/Formularen, aber ohne Same-Origin-Berechtigung unterstützen. Dateiabrufe
senden keine Zugangsdaten mit; Dateiserver auf anderen Origins müssen CORS erlauben. Ohne konfigurierten
PowerPoint-Viewer meldet der Bereich diese Einschränkung und bietet das externe Öffnen an.

Die Methode `previewContent({ attachment, url })` des Elements öffnet vom Host bereitgestellte
Inhalte lokal. Für eine Website verwenden Sie `attachment: { name: 'Project site', kind:
'other', mime: 'text/html' }`. Dadurch wird kein gemeinsamer Kanal-Tab angelegt. Die
Dateiansicht bietet außerdem **Preview website**. Websites, die das Einbetten verweigern, lassen sich
extern öffnen.

## Gemeinsame Kanal-Tabs und Excel-Kopien

Verwenden Sie **Pin as tab** neben einer hochgeladenen Datei oder **Add tab**, um eine benannte Website hinzuzufügen.
Definitionen werden geteilt; jeder Benutzer wählt seinen eigenen Tab. Der Ersteller kann ihn umbenennen
oder entfernen, über Besitzprüfungen im Client. Wird ein Tab anderswo entfernt, bleibt ein bereits geöffneter
lokaler Bereich verfügbar, bis er geschlossen wird.

Excel bietet lokales Bearbeiten und **Save copy to channel**, wenn Dateispeicher
konfiguriert ist. Jedes Speichern lädt eine eindeutig benannte Kopie hoch und postet sie im ursprünglichen
Kanal, wobei das Original erhalten bleibt. Fehlgeschlagene Speichervorgänge behalten die Änderungen; beim Navigieren wird
das Verwerfen ungespeicherter Arbeit bestätigt, und es wird auf einen laufenden Speichervorgang gewartet. Das ist kein gemeinsames Bearbeiten
und kein versioniertes Zurückschreiben. Hosts, die den Arbeitsbereich oder das Element zwangsweise ersetzen, müssen
die Bestätigung für ungespeicherte Arbeit selbst übernehmen.

Rohe Client-Aktionen:

```ts
const tab = teams.addTab('Budget', { type: 'file', attachment });
teams.addTab('Project site', { type: 'website', url: 'https://example.com' });
teams.renameTab(tab!.id, 'Budget review');
teams.removeTab(tab!.id);
await teams.saveFileCopy(channelId, file);
```

`TeamsState.tabs` enthält die Definitionen des ausgewählten Kanals;
`canUploadFiles` zeigt konfigurierten Speicher an. `OpenFileDetail.channelId` hält
den beim Öffnen erfassten Kanal fest. `previewContent` akzeptiert dieses Feld ebenfalls, wenn
ein Host Kanalkopien aktivieren möchte. Besitzprüfungen erfordern eine serverseitige
Autorisierung, bevor sie als Zugriffskontrolle gelten können.

## Dateien anlegen, hochladen und suchen

Beide Dateiansichten bieten Hochladen, das Anlegen leerer Excel-Arbeitsmappen und Suche. Uploads
und neue Arbeitsmappen gehen in den ausgewählten Kanal, der neben den Bedienelementen genannt wird.
Sie benötigen konfigurierten Speicher; Fehler können im ursprünglichen Kanal wiederholt werden.
Gleichnamige Dateien erhalten getrennte Speicher-URLs. Uploads akzeptieren bis zu 20 Dateien zu je 32 MiB;
die Oberfläche meldet Aktivität, bietet aber weder Byte-Fortschritt noch Abbruch.

```ts
await teams.uploadFiles(channelId, files);
const workbook = await teams.createWorkbook(channelId, 'Budget'); // Budget.xlsx
```

Das Anlegen postet einen Anhang, der von einer echten XLSX-Datei gestützt wird. Die Suche filtert die geladenen
Metadaten nach Dateiname, Autor, Kanal und erkanntem Typ. Sie durchsucht nicht den
Dateiinhalt. Diese Aktionen erfassen den ausdrücklichen Kanal vor dem asynchronen
Anlegen oder Hochladen, sodass die Auswahl eines anderen Kanals das Ergebnis nicht umleiten kann.

## Ereignisse

| Ereignis              | Bedeutung                                                                                                 |
| --------------------- | --------------------------------------------------------------------------------------------------------- |
| `teams-ready`         | Das Element ist bereit.                                                                                   |
| `teams-open-file`     | Ein Benutzer hat eine angehängte Datei geöffnet. Abbrechbar: `preventDefault()`, um sie selbst zu öffnen. |
| `teams-config-change` | Der Benutzer hat neue Servereinstellungen im Einstellungsdialog gespeichert.                              |

Schreibweisen je Binding: `onReady`, `onOpenFile`, `onConfigChange` (React, Solid, Svelte, Vanilla); `@ready`, `@open-file`, `@config-change` (Vue); `(ready)`, `(openFile)`, `(configChange)` (Angular).

## Der rohe Client

`createTeams(options)` (der rohe Hook in jedem Binding) liefert den Core-Client:

```ts
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
teams.on('notice', (notice) => show(notice)); // errors a user should see
```

`getState()` ist ein unveränderlicher `TeamsState`: Verbindungs-`status`, `channels` mit Zählern für Ungelesenes, `messages` und `files` des ausgewählten Kanals, `people` (Präsenz), `typing`, der Zustand des Eingabefelds, Suchergebnisse und `call`. Zu den Aktionen gehören `select`, `createChannel`, `send`, `startReply`, `startEdit`, `deleteMessage`, `toggleReaction`, `notifyTyping`, `setAvailability`, `search`, `openCall`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`, `toggleScreenShare`, `toggleHand` und `destroy`. Siehe [Architektur](/architecture), wie die Schichten zusammenspielen.
