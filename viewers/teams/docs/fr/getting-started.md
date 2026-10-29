# Premiers pas

OpenTeams est un espace de travail d’équipe open source avec votre propre serveur : canaux et
discussions, présence, réunions avec vidéo et partage d’écran via WebRTC, et fichiers Office
partagés dans la conversation. Vous installez une intégration pour votre framework, vous la faites
pointer vers un serveur (ou vous démarrez sans), et vous obtenez l’interface complète, ou un hook
brut pour construire la vôtre.

::: warning Bêta
Ce n’est pas Microsoft Teams et il n’est pas affilié à Microsoft. Il n’y a **pas de chiffrement de
bout en bout** : quiconque exploite le serveur peut lire chaque message et chaque fichier. Consultez
les [Limites](/fr/limitations) avant de vous y fier.
:::

## 1. Installer une intégration

Chaque intégration est autonome : l’élément `<teams-app>` y est inclus, et
[`ooxml-core`](https://www.npmjs.com/package/ooxml-core) (la logique) ainsi que
[`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui) (les primitives visuelles) s’installent comme
des dépendances ordinaires.

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

| Paquet                     | Pour                               | Dépendance homologue |
| -------------------------- | ---------------------------------- | -------------------- |
| `openteams-react-viewer`   | React 18+                          | `react`              |
| `openteams-vue-viewer`     | Vue 3.4+                           | `vue`                |
| `openteams-angular-viewer` | Angular 17+ (standalone)           | `@angular/core`      |
| `openteams-svelte-viewer`  | Svelte 5                           | `svelte`             |
| `openteams-solid-viewer`   | SolidJS 1.9                        | `solid-js`           |
| `openteams-vanilla-viewer` | sans framework (`<teams-app>`)     | aucune               |
| `openteams-server`         | le serveur de référence (Node 22+) | aucune               |

Utilisez une seule intégration par page : chacune enregistre les mêmes éléments personnalisés. Les
noms sont en `openteams-*` parce que le nom `openteams` seul sur npm appartient à un projet sans
rapport.

## 2. La monter

React :

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

JavaScript natif :

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

Ou simplement du balisage, après avoir appelé `defineTeamsApp()` depuis `openteams-vanilla-viewer` :

```html
<teams-app
	workspace-id="acme"
	user-name="Ada"
	server-config='{"mode":"server","syncUrl":"wss://teams.example.com/sync","signalingUrl":"wss://teams.example.com/signal","iceServers":[{"urls":"stun:stun.example.com:3478"}]}'
></teams-app>
```

Toutes les intégrations acceptent les mêmes propriétés : `workspaceId` (la salle partagée : lettres,
chiffres, `-` et `_`), `userName`, `userId` (un identifiant stable ; par défaut, un identifiant
mémorisé dans ce navigateur), `config`, `uploadFile` (la destination des pièces jointes), `openers`
(des gestionnaires par type Office) et trois événements : prêt, ouverture de fichier (appelez
`preventDefault()` dessus pour ouvrir le fichier vous-même) et changement de configuration
(`onReady`, `onOpenFile`, `onConfigChange` en React, Solid, Svelte et JS natif ; `@ready`,
`@open-file`, `@config-change` en Vue ; `(ready)`, `(openFile)`, `(configChange)` en Angular).
Consultez les [guides par framework](/frameworks/react) pour chaque intégration.

Si vous omettez `config`, l’élément propose une boîte de dialogue de paramètres (mémorisée dans ce
navigateur) ou s’exécute en mode local.

## 3. Choisir où vit l’état

| Mode     | `config`                                                        | Ce qui se passe                                                                                                                                                    |
| -------- | --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `local`  | `{ mode: 'local', iceServers: [...] }`                          | Pas de serveur. Les onglets d’un même navigateur partagent discussions, présence et appels via `BroadcastChannel`. Les fichiers ne sont partagés que par leur nom. |
| `server` | `{ mode: 'server', syncUrl, signalingUrl, iceServers, token? }` | Votre serveur synchronise le document partagé, relaie la signalisation des appels et (facultativement) stocke les fichiers.                                        |

Le mode local est celui qu’utilisent les [démos en direct](/demos). Dès qu’il y a plus d’un
navigateur, lancez un serveur :

```bash
npx openteams-server --insecure-dev     # uniquement pour le développement local : 127.0.0.1:8787
```

Utilisez ensuite `ws://127.0.0.1:8787/sync` et `ws://127.0.0.1:8787/signal` (ou `wss://` derrière
TLS). Le [guide du serveur](/fr/server) couvre le protocole, les variables d’environnement et ce que
le serveur de référence ne fait pas ; le [déploiement](/fr/deploy) couvre TLS, les jetons, TURN et
Docker.

## 4. Ou construire votre propre interface

Chaque intégration expose aussi un hook brut au-dessus du client du cœur : l’état en entrée, des
actions simples en sortie.

```tsx
import { useTeams } from 'openteams-react-viewer';

const { client, state } = useTeams({
	workspaceId: 'acme',
	user: { id: 'u1', name: 'Ada' },
	config,
});
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

`state` est un instantané immuable `TeamsState` (état de la connexion, canaux avec compteurs de
non-lus, messages et fichiers du canal sélectionné, personnes, saisie en cours, résultats de
recherche, l’appel). Les actions comprennent `select`, `createChannel`, `send`, `startReply`,
`startEdit`, `deleteMessage`, `toggleReaction`, `search`, `openCall`, `joinCall`, `leaveCall`,
`toggleMic`, `toggleCamera`, `toggleScreenShare` et `toggleHand`. La démo React montre les deux
styles côte à côte.

## Exécuter depuis les sources

Nécessite Node 22+ et Bun. Ce dossier est `viewers/teams` du dépôt
[`ooxml`](https://github.com/ChristopherVR/ooxml) : exécutez `bun install` à la racine du dépôt, puis les
commandes ci-dessous ici. Les démos (`demos/teams/*` à la racine) et les tests utilisent les sources de
`ooxml-core` et `ooxml-ui` du même arbre.

```bash
bun install        # à la racine du dépôt
bun run dev        # dans viewers/teams : serveur :8787, démo vanilla :5173, démo React :5174
```

Ouvrez `http://127.0.0.1:5173/?name=Ada` puis, dans un autre onglet, `...?name=Bob` : deux
personnes dans un même espace de travail. `?local=1` s’exécute sans le serveur. `TEAMS_USE_DIST=1`
fait utiliser aux démos les paquets construits (`bun run build:packages` d’abord) plutôt que les sources.
`TEAMS_DEMOS=vanilla,vue` choisit les démos que lance `bun run dev`. Les caméras et les microphones exigent `localhost` ou HTTPS.
