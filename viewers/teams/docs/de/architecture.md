# Architektur

OpenTeams ist genauso aufgeteilt wie die anderen Produkte, die auf [`ooxml`](https://github.com/ChristopherVR/ooxml)
aufbauen (Word, Excel, PowerPoint): Die Logik und die visuellen Grundbausteine liegen dort, einschließlich des
Elements `<teams-app>` selbst. Dieser Ordner (`viewers/teams` des Repositorys `ooxml`) enthält die Framework-Bindings,
den Referenzserver und die Dokumentation. Die Demos liegen in `demos/teams/` und die Browser-Tests in `e2e/teams/`
im Root.

| Schicht                 | Ort                                                                             | Verantwortlich für                                                                                                     |
| ----------------------- | ------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Logik                   | `ooxml-core/teams` (`src/core/teams` im Repository `ooxml`)                     | der CRDT für den Chat, Präsenz, Anrufe, Signalisierung, Serverkonfiguration, der Store `createTeamsClient`             |
| Zusammenarbeit          | `ooxml-core/collab`                                                             | Yjs-Sitzungen und Transporte (WebSocket, `BroadcastChannel`), formatneutral                                            |
| Visuelle Grundbausteine | `ooxml-ui/teams` (`src/ui/src/teams`)                                           | Avatar, App-Leiste, Kanalliste, Chatliste, Eingabefeld, Vorbeitritt, Anrufraster und Steuerelemente                    |
| App                     | `ooxml-ui/teams` (`src/ui/src/teams/app`)                                       | `<teams-app>`: setzt die Grundbausteine über dem Store zusammen, Einstellungsdialog, rohe Store-Hilfsfunktionen        |
| Bind-Schicht            | `ooxml-ui/teams` (`src/ui/src/teams/app/bind.ts`)                               | `bindTeams`, `pickTeamsProps` und `TEAMS_CONTRACT`: die Verdrahtung von Props und Ereignissen, die die Bindings nutzen |
| Web-Komponente          | `packages/web-component` hier (privat, `teams-viewer`, wird nie veröffentlicht) | exportiert `ooxml-ui/teams` neu, damit die Bindings einen Namen importieren; wird in jedes Binding eingebettet         |
| Bindings                | `packages/{react,vue,angular,svelte,solid,vanilla}` hier                        | Lebenszyklus-Adapter plus je ein roher Hook, veröffentlicht als `openteams-<framework>-viewer`                         |
| Referenzserver          | `server/` hier                                                                  | Yjs-Synchronisation, Signalisierungsweiterleitung, Dateispeicher, veröffentlicht als `openteams-server`                |

Die eigene Beschreibung des Bereichs im Core finden Sie in
[`docs/teams-area.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md).

## Der Client: ein Zustand, einfache Aktionen {#the-client-one-state-plain-actions}

Alles, was eine Oberfläche braucht, kommt von einem Client, der mit `createTeamsClient` erzeugt wird (die Bindings
verpacken ihn als `createTeams`, das einen geschützten `localStorage` für Offline-Schnappschüsse ergänzt):

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
```

`getState()` liefert einen unveränderlichen `TeamsState`-Schnappschuss. Jede Schaltfläche ist eine Aktion des Clients.
Fehler, die ein Benutzer sehen soll, kommen über `on('notice', ...)`. Die Oberfläche greift nie direkt auf Yjs,
`RTCPeerConnection` oder den Server zu: **Zustand hinein, Aktionen hinaus**.

## `<teams-app>` {#} {#} {#} {#} {#}

Ein Lit-Element mit einer Vorlage und einer echten CSS-Datei (`teams-app.ts`, `teams-app.css`, in `src/ui/src/teams/app`).
Es entscheidet, was auf dem Bildschirm zu sehen ist, und leitet die Absichten des Benutzers an Client-Aktionen weiter.
Farben, Abstände und Rundungen stammen aus den Tokens von `ooxml-ui` (`var(--office-...)`). Alles, was ein Peer sendet,
wird als Text dargestellt (Semantik von `textContent`), nie als Markup, und der Core validiert alles erneut, was er aus
dem gemeinsamen Dokument liest.

## Bindings {#bindings}

Ein Binding ist ein Lebenszyklus-Adapter: Es erzeugt `<teams-app>`, reicht Props mit `bindTeams` weiter (einschließlich
der Klasse; `class` oder `className` je nach Framework) und gibt Ereignisse erneut aus. Das alles geschieht in
`ooxml-ui/teams` (`src/ui/src/teams/app/bind.ts`, auf dem gemeinsamen Baukasten `ooxml-ui/bindings`). `onOpenFile` ist
überall `(detail, event)`, außer bei Angular, dessen Ausgabe mit einem einzelnen Wert `{ detail, event }` sendet. Jedes
Binding stellt außerdem den rohen Hook über `createTeams` bereit:

| Binding  | Komponente                      | Roher Hook                                                |
| -------- | ------------------------------- | --------------------------------------------------------- |
| React    | `<Teams />`                     | `useTeams(options)` (`useTeamsClient`, `useTeamsState`)   |
| Vue 3    | `<Teams />`                     | `useTeams(() => options)` gibt flache Refs zurück         |
| Solid    | `Teams(props)`                  | `createTeamsClient(() => options)` gibt Signale zurück    |
| Svelte 5 | `Teams.svelte` (Default-Export) | `teamsStore(options)` aus `/runtime` (ein lesbarer Store) |
| Angular  | `<teams-workspace>`             | `TeamsService` (Signale)                                  |
| Vanilla  | `mountTeams(el, props)`         | `createTeams(options)`                                    |

Ein Fehler im Verhalten wird einmal behoben, in der Web-Komponente oder im Core. Ein Fehler im Hook oder in der
Verdrahtung wird in jedem Binding im selben Change behoben.

Zur Build-Zeit bettet `scripts/build-packages.mjs` die private Web-Komponente, einschließlich CSS, in jedes Framework-Bundle
ein und flacht ihre Deklarationen ab. So importiert ein veröffentlichtes Paket nur, was sein Manifest deklariert
(`ooxml-core`, `ooxml-ui`, `lit` und die Framework-Peer-Abhängigkeit).

## Chat, Präsenz und Anrufe {#chat-presence-and-calls}

- **Chat.** Kanäle liegen in einer `Y.Map`, Nachrichten in einer `Y.Map` je Kanal, Reaktionen in einer flachen Map,
  die nach Nachricht, Emoji und Benutzer geschlüsselt ist. So überschreiben sich gleichzeitige Reaktionen nie.
  Bearbeiten und Löschen sind auf dem Client nur dem Verfasser erlaubt (ein Löschen ist ein weicher Grabstein).
- **Präsenz** ist Yjs-Awareness: Wer online ist, Tippen, Status und Statusnachricht. Sie ist unterstützend und nicht
  authentifiziert. Peers erhalten zusätzlich das veraltete Dreiwert-Feld `availability`.
- **Angeheftete Nachrichten** sind je Kanal eine Yjs-Map; jedes Mitglied kann anheften oder lösen.
- **Anrufe** sind ein vollständiges Netz aus Peer-Verbindungen nach dem Muster „Perfect Negotiation“. Stummschalten,
  Kamera und Bildschirmfreigabe tauschen Tracks aus, ohne neu auszuhandeln. Das Netz ist bis etwa ein Dutzend Personen
  ehrlich (`maxParticipants`, Standard 12).
- **Dateien** sind Links, nie eingebettete Bytes: Sie werden über `uploadFile` oder den Endpunkt `/files` des Servers
  hochgeladen. Word-, Excel-, PowerPoint- und Visio-Dateien werden am Namen erkannt, damit ein Host sie im passenden
  Viewer öffnen kann (die Prop `openers` oder das Ereignis „Datei öffnen“).

## Transporte {#transports}

Im Modus `local` synchronisiert und signalisiert der Client über `BroadcastChannel`. So sehen sich nur Tabs desselben
Browsers. Die Demos auf GitHub Pages laufen so. Im Modus `server` verwendet der Client für die Synchronisation das
Drahtformat von y-websocket und für die Signalisierung ein JSON-Relais. Siehe [Eigenen Server mitbringen](/de/server).
