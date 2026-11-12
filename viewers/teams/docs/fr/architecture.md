# Architecture

OpenTeams est découpé de la même manière que les autres produits construits sur
[`ooxml`](https://github.com/ChristopherVR/ooxml) (Word, Excel, PowerPoint) : la logique et les
primitives visuelles y vivent, y compris l’élément `<teams-app>` lui-même. Ce dossier
(`viewers/teams` du dépôt `ooxml`) contient les liaisons par framework, le serveur de référence et
la documentation ; les démos se trouvent dans `demos/teams/` et les tests de navigateur dans
`e2e/teams/`, à la racine.

| Couche               | Où                                                                 | Responsable de                                                                                                                             |
| -------------------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------ |
| Logique              | `ooxml-core/teams` (`src/core/teams` dans le dépôt ooxml)          | le CRDT de discussion, la présence, les appels, la signalisation, la configuration du serveur, le magasin `createTeamsClient`              |
| Collaboration        | `ooxml-core/collab`                                                | les sessions Yjs et les transports (WebSocket, `BroadcastChannel`), indépendants du format                                                 |
| Primitives visuelles | `ooxml-ui/teams` (`src/ui/src/teams`)                              | avatar, barre d’applications, liste des canaux, liste des discussions, zone de saisie, écran de pré-connexion, grille d’appel et commandes |
| Application          | `ooxml-ui/teams` (`src/ui/src/teams/app`)                          | `<teams-app>` : assemble les primitives sur le magasin, la boîte de dialogue des paramètres, les helpers bruts du magasin                  |
| Couche de liaison    | `ooxml-ui/teams` (`src/ui/src/teams/app/bind.ts`)                  | `bindTeams`, `pickTeamsProps` et `TEAMS_CONTRACT` : le câblage des propriétés et des événements qu’utilisent les liaisons                  |
| Composant web        | `packages/web-component` ici (`teams-viewer` privé, jamais publié) | réexporte `ooxml-ui/teams`, afin que les liaisons importent un seul nom ; intégré dans chaque liaison                                      |
| Liaisons             | `packages/{react,vue,angular,svelte,solid,vanilla}` ici            | adaptateurs de cycle de vie plus un hook brut pour chacune, publiés sous `openteams-<framework>-viewer`                                    |
| Serveur de référence | `server/` ici                                                      | synchronisation Yjs, relais de signalisation, stockage des fichiers, publié sous `openteams-server`                                        |

La description de ce domaine par le cœur se trouve dans
[`docs/teams-area.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md).

## Le client : un seul état, des actions simples {#the-client-one-state-plain-actions}

Tout ce dont une interface a besoin provient d’un seul client créé par `createTeamsClient` (les
liaisons l’enveloppent sous le nom `createTeams`, qui ajoute un `localStorage` protégé pour les
instantanés hors ligne) :

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
```

`getState()` renvoie un instantané `TeamsState` immuable. Chaque bouton est une action du client.
Les erreurs que l’utilisateur doit voir arrivent par `on('notice', ...)`. L’interface ne touche
jamais directement à Yjs, `RTCPeerConnection` ni au serveur : **l’état en entrée, les actions en
sortie**.

## `<teams-app>` {#} {#} {#} {#} {#}

Un seul élément Lit, avec un seul gabarit et un vrai fichier CSS (`teams-app.ts`, `teams-app.css`,
dans `src/ui/src/teams/app`). Il décide de ce qui est affiché et transmet les intentions de
l’utilisateur aux actions du client ; les couleurs, les espacements et les rayons proviennent des
jetons de `ooxml-ui` (`var(--office-...)`). Tout ce qu’un pair envoie est affiché comme du texte
(sémantique `textContent`), jamais comme du balisage, et le cœur revalide tout ce qu’il relit du
document partagé.

## Liaisons {#bindings}

Une liaison est un adaptateur de cycle de vie : elle crée `<teams-app>`, transmet les propriétés
avec `bindTeams` (la classe comprise ; `class` ou `className` selon le framework) et réémet les
événements, le tout dans `ooxml-ui/teams` (`src/ui/src/teams/app/bind.ts`, sur la boîte à outils
partagée `ooxml-ui/bindings`). `onOpenFile` est `(detail, event)` partout, sauf en Angular, dont la
sortie à valeur unique émet `{ detail, event }`. Chacune expose aussi le hook brut au-dessus de
`createTeams` :

| Liaison  | Composant                          | Hook brut                                                  |
| -------- | ---------------------------------- | ---------------------------------------------------------- |
| React    | `<Teams />`                        | `useTeams(options)` (`useTeamsClient`, `useTeamsState`)    |
| Vue 3    | `<Teams />`                        | `useTeams(() => options)` renvoie des refs superficielles  |
| Solid    | `Teams(props)`                     | `createTeamsClient(() => options)` renvoie des signaux     |
| Svelte 5 | `Teams.svelte` (export par défaut) | `teamsStore(options)` depuis `/runtime` (un store lisible) |
| Angular  | `<teams-workspace>`                | `TeamsService` (signaux)                                   |
| Vanilla  | `mountTeams(el, props)`            | `createTeams(options)`                                     |

Un bogue de comportement se corrige une seule fois, dans le composant web ou dans le cœur. Un bogue
de câblage ou de hook se corrige dans chaque liaison, dans le même changement.

Au moment de la compilation, `scripts/build-packages.mjs` intègre le composant web privé, CSS
compris, dans chaque paquet de framework, et aplatit ses déclarations ; ainsi, une archive publiée
n’importe que ce que son manifeste déclare (`ooxml-core`, `ooxml-ui`, `lit` et la dépendance
homologue du framework).

## Discussion, présence et appels {#chat-presence-and-calls}

- **Discussion.** Les canaux vivent dans un seul `Y.Map`, les messages dans un `Y.Map` par canal,
  les réactions dans une table plate indexée par message, emoji et utilisateur, afin que des
  réactions simultanées ne s’écrasent jamais. Les modifications et les suppressions sont réservées
  à l’auteur côté client (une suppression est une pierre tombale logique).
- **La présence** repose sur l’awareness de Yjs : personnes en ligne, saisie en cours, statut et
  message de statut. Elle est indicative et non authentifiée. Les pairs reçoivent aussi l’ancienne
  `availability` à trois valeurs.
- **Les épingles** forment une carte Yjs par canal ; tout membre peut épingler ou désépingler.
- **Les appels** sont un maillage complet de connexions entre pairs, selon le schéma de négociation
  parfaite ; la coupure du micro, la caméra et le partage d’écran remplacent les pistes sans
  renégocier. Le maillage reste honnête jusqu’à une douzaine de personnes environ (`maxParticipants`,
  12 par défaut).
- **Les fichiers** sont des liens, jamais des octets intégrés : ils sont envoyés via `uploadFile` ou
  le point de terminaison `/files` du serveur. Les fichiers Word, Excel, PowerPoint et Visio sont
  reconnus à leur nom, afin qu’un hôte puisse les ouvrir dans la visionneuse correspondante (la
  propriété `openers` ou l’événement d’ouverture de fichier).

## Transports {#transports}

En mode `local`, le client synchronise et signale via `BroadcastChannel` : seuls les onglets d’un
même navigateur se voient ; les démos GitHub Pages fonctionnent ainsi. En mode `server`, il utilise
le format de fil de y-websocket pour la synchronisation et un relais JSON pour la signalisation ;
voir [Apportez votre serveur](/fr/server).
