# Erste Schritte

OpenTeams ist ein quelloffener Team-Arbeitsbereich, bei dem Sie Ihren eigenen Server mitbringen: Kanäle und Chat, Präsenz,
Besprechungen mit Video und Bildschirmfreigabe über WebRTC sowie Office-Dateien, die in der Unterhaltung geteilt werden. Sie
installieren ein Binding für Ihr Framework, verweisen es auf einen Server (oder starten ganz ohne) und erhalten die
vollständige Oberfläche oder einen rohen Hook, mit dem Sie eine eigene bauen.

::: warning Eine frühe Implementierung
Dies ist nicht Microsoft Teams und steht in keiner Verbindung zu Microsoft. Es gibt **keine
Ende-zu-Ende-Verschlüsselung**: Wer den Server betreibt, kann jede Nachricht und jede Datei lesen. Lesen Sie
[Einschränkungen](/de/limitations), bevor Sie sich darauf verlassen.
:::

## 1. Ein Binding installieren

Jedes Binding ist in sich geschlossen: Das Element `<teams-app>` ist eingebunden, und
[`ooxml-core`](https://www.npmjs.com/package/ooxml-core) (die Logik) sowie
[`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui) (die visuellen Grundbausteine) werden als normale
Abhängigkeiten installiert.

::: code-group

```bash [React]
npm install openteams-react-viewer react
```

```bash [Vue]
npm install openteams-vue-viewer vue
```

```bash [Angular]
npm install openteams-angular-viewer @angular/core
```

```bash [Svelte]
npm install openteams-svelte-viewer svelte
```

```bash [Solid]
npm install openteams-solid-viewer solid-js
```

```bash [Vanilla]
npm install openteams-vanilla-viewer
```

:::

| Paket                      | Für                            | Peer-Abhängigkeit |
| -------------------------- | ------------------------------ | ----------------- |
| `openteams-react-viewer`   | React 18+                      | `react`           |
| `openteams-vue-viewer`     | Vue 3.4+                       | `vue`             |
| `openteams-angular-viewer` | Angular 17+ (Standalone)       | `@angular/core`   |
| `openteams-svelte-viewer`  | Svelte 5                       | `svelte`          |
| `openteams-solid-viewer`   | SolidJS 1.9                    | `solid-js`        |
| `openteams-vanilla-viewer` | kein Framework (`<teams-app>`) | keine             |
| `openteams-server`         | der Referenzserver (Node 22+)  | keine             |

Verwenden Sie pro Seite ein Binding: Jedes registriert dieselben Custom Elements. Die Namen lauten `openteams-*`,
weil der schlichte Name `openteams` auf npm zu einem unabhängigen Projekt gehört.

## 2. Einbinden

React:

```tsx
import { Teams } from 'openteams-react-viewer';

const config = {
	mode: 'server' as const,
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
};

export function Workspace() {
	return <Teams workspaceId="acme" userName="Ada" config={config} />;
}
```

Vanilla JavaScript:

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

// `config` as in the React example above.
const teams = mountTeams(document.getElementById('teams')!, {
	workspaceId: 'acme',
	userName: 'Ada',
	config,
	onOpenFile: (detail) => console.log('open', detail.attachment.kind, detail.url),
});
teams.update({ workspaceId: 'acme', userName: 'Ada Lovelace', config });
teams.destroy();
```

Oder als reines Markup, nachdem Sie `defineTeamsApp()` aus `openteams-vanilla-viewer` aufgerufen haben:

```html
<teams-app
	workspace-id="acme"
	user-name="Ada"
	server-config='{"mode":"server","syncUrl":"wss://teams.example.com/sync","signalingUrl":"wss://teams.example.com/signal","iceServers":[{"urls":"stun:stun.example.com:3478"}]}'
></teams-app>
```

Jedes Binding nimmt dieselben Props entgegen: `workspaceId` (der gemeinsame Raum: Buchstaben, Ziffern, `-` und `_`),
`userName`, `userId` (eine stabile ID; standardmäßig eine, die in diesem Browser gespeichert ist), `config`,
`uploadFile` (wohin Anhänge gehen), `openers` (Handler je Office-Typ) und drei Ereignisse:
bereit, Datei öffnen (rufen Sie darauf `preventDefault()` auf, um die Datei selbst zu öffnen) und Konfigurationsänderung
(`onReady`, `onOpenFile`, `onConfigChange` in React, Solid, Svelte und Vanilla; `@ready`,
`@open-file`, `@config-change` in Vue; `(ready)`, `(openFile)`, `(configChange)` in Angular). Die
[Framework-Anleitungen](/frameworks/react) behandeln jedes Binding.

Lassen Sie `config` weg, dann bietet das Element einen Einstellungsdialog an (in diesem Browser gespeichert) oder läuft
im lokalen Modus.

## 3. Festlegen, wo der Zustand liegt

| Modus    | `config`                                                        | Was passiert                                                                                                                   |
| -------- | --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `local`  | `{ mode: 'local', iceServers: [...] }`                          | Kein Server. Tabs eines Browsers teilen Chat, Präsenz und Anrufe über `BroadcastChannel`. Dateien werden nur per Name geteilt. |
| `server` | `{ mode: 'server', syncUrl, signalingUrl, iceServers, token? }` | Ihr Server synchronisiert das gemeinsame Dokument, leitet die Anruf-Signalisierung weiter und speichert (optional) Dateien.    |

Der lokale Modus ist der, den die [Live-Demos](/demos) verwenden. Sobald mehr als ein Browser beteiligt ist, betreiben Sie einen
Server:

```bash
npx openteams-server                    # the reference server on 127.0.0.1:8787
```

Verwenden Sie dann `ws://127.0.0.1:8787/sync` und `ws://127.0.0.1:8787/signal` (oder `wss://` hinter TLS).
Die [Server-Anleitung](/de/server) behandelt das Protokoll, die Umgebungsvariablen und das, was der Referenzserver
nicht kann; [Bereitstellung](/de/deploy) behandelt TLS, Token, TURN und Docker.

## 4. Oder eine eigene Oberfläche bauen

Jedes Binding stellt außerdem einen rohen Hook über dem Core-Client bereit: Zustand hinein, einfache Aktionen hinaus.

```tsx
import { useTeams } from 'openteams-react-viewer';

const { client, state } = useTeams({
	workspaceId: 'acme',
	user: { id: 'u1', name: 'Ada' },
	config,
});
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

`state` ist ein unveränderlicher `TeamsState`-Schnappschuss (Verbindungsstatus, Kanäle mit ungelesenen Zählern,
Nachrichten und Dateien des ausgewählten Kanals, Personen, Tippanzeigen, Suchergebnisse, der Anruf). Zu den Aktionen
gehören `select`, `createChannel`, `send`, `startReply`, `startEdit`, `deleteMessage`,
`toggleReaction`, `search`, `openCall`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`,
`toggleScreenShare` und `toggleHand`. Die React-Demo zeigt beide Stile nebeneinander.

## Aus dem Quellcode starten

Erfordert Node 22+ und Bun. Wenn das Repository [`ooxml`](https://github.com/ChristopherVR/ooxml)
neben diesem ausgecheckt ist (`../ooxml-core`), verwenden Demos und Tests dessen Quellcode; andernfalls
nutzen sie die veröffentlichten Pakete `ooxml-core` und `ooxml-ui`.

```bash
bun install
bun run dev        # server on :8787, vanilla demo on :5173, React demo on :5174
```

Öffnen Sie `http://127.0.0.1:5173/?name=Ada` und in einem anderen Tab `...?name=Bob`: zwei Personen in einem
Arbeitsbereich. `?local=1` läuft ohne den Server. `TEAMS_USE_DIST=1` lässt die Demos die gebauten
Pakete verwenden (zuvor `bun run build:packages`) und die veröffentlichten ooxml-Pakete statt jeglichen Quellcodes.
Kameras und Mikrofone erfordern `localhost` oder HTTPS.
