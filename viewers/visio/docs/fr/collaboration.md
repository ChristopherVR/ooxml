# Collaboration {#collaboration}

Fichier > Partager démarre une session en direct entre des fenêtres d'un même navigateur. Une application hôte peut aussi partager entre appareils via la propriété `collaboration`, qui accepte les mêmes options que tout autre éditeur Office.

::: warning Partiel
Cette version n'inclut aucun service hébergé. Avec Fichier > Partager, rien ne quitte votre navigateur. Entre appareils, vous apportez votre propre serveur (une URL compatible y-websocket), fournisseur ou session Yjs ; les pages et les formes sont fusionnées élément par élément, tout le reste est partagé en bloc (voir ci-dessous).
:::

## Intégration : la propriété `collaboration` {#embedding-the-collaboration-property}

`<visio-viewer>` implémente le contrat d'intégration partagé de `ooxml-core/collab` (`CollaborationOptions`, `CollaborationState`, `CollaborativeEditor`), le même que Word, Excel et PowerPoint ; un hôte qui intègre plusieurs produits n'intègre donc qu'une fois.

```ts
viewer.collaboration = {
	roomId: 'floor-plan', // 1 to 128 letters, digits, hyphens or underscores
	serverUrl: 'wss://example.com/collab', // optional: a y-websocket compatible server
	user: { name: 'Ada', color: '#2563eb' }, // optional; `role: 'viewer'` makes this window read-only
	intent: 'join', // optional: never seed an empty room with the drawing already open
};
viewer.addEventListener('collaboration-change', (event) => console.log(event.detail));
viewer.collaborationState; // { active, roomId, status, synced, people }
viewer.reconnectCollaboration();
viewer.stopCollaboration(); // or: viewer.collaboration = null
```

- **Transport** : `session` (une `CollabSession` de l'hôte, jamais détruite par le visualiseur), sinon `provider`, sinon `webrtc`, sinon `serverUrl`, sinon un `BroadcastChannel` entre fenêtres de ce navigateur.
- **Déclaratif** : tant que `collaboration` est définie, le dessin ouvert est partagé, ainsi que celui chargé ensuite. La valeur `null` quitte la salle. Une session ouverte par le visualiseur lui-même est détruite en quittant ; une `session` de l'hôte ne l'est pas.
- **Intention (`intent`)** : avec `join`, une salle vide n'est pas initialisée avec le dessin local, de sorte qu'un arrivant tardif ne puisse pas écraser une salle pas encore chargée. Un dessin chargé ou modifié ensuite est tout de même publié. Par défaut : `create`.
- **Rôles** : `role: 'viewer'` met le visualiseur en lecture seule tant qu'il est dans la salle (`canEdit` est rétabli en sortant). Le rôle n'est qu'indicatif ; un serveur doit le faire respecter.
- **Événements** : `collaboration-change` fournit un `CollaborationState` quand le partage démarre ou s'arrête et quand l'état de la connexion, la première synchronisation ou les personnes changent. Un échec de connexion est signalé par `document-error`. Fichier > Partager utilise le même contrôleur et émet donc aussi l'événement.
- **Bindings** : chaque binding transmet une propriété `collaboration` (React, Vue, Solid, Svelte, `mountViewer`/`update` en vanilla, `[collaboration]` en Angular) et relaie `collaboration-change` (`events['collaboration-change']`, `@collaboration-change` en Vue, `(collaborationChange)` en Angular). Le démontage quitte la salle.
- **Types** : `VisioCollaborationOptions` (les options partagées plus `onWriteBack`, `writeBackDebounceMs` et `writeBackRetry`) et `VisioCollaborationState` sont exportés par `ooxml-ui/visio`.

## Serveur et poste à poste : jeton et WebRTC {#server-and-peer-to-peer-auth-token-and-webrtc}

```ts
// A y-websocket compatible server. The token is sent as the `token` query parameter:
// wss://example.com/collab/floor-plan?token=<authToken>
viewer.collaboration = {
	roomId: 'floor-plan',
	serverUrl: 'wss://example.com/collab',
	authToken: await getToken(),
};

// Peer to peer over y-webrtc. Pass the `WebrtcProvider` class from your own y-webrtc install;
// the viewer does not bundle it. The room password defaults to `authToken`.
import { WebrtcProvider } from 'y-webrtc';
viewer.collaboration = {
	roomId: 'floor-plan',
	authToken: 'room password',
	webrtc: { WebrtcProvider, signaling: ['wss://signal.example.com'] },
};
```

Le visualiseur transmet ces options telles quelles à `openCollaborationSession`. Le jeton n'authentifie qu'auprès du serveur que vous exploitez (le serveur doit le vérifier), et un mot de passe WebRTC ne fait qu'écarter de la salle les pairs qui ne l'ont pas ; ni l'un ni l'autre ne rend les rôles plus qu'indicatifs. Les salles poste à poste ne confirment jamais une première synchronisation : c'est le délai de grâce de la session qui décide quand cette fenêtre peut écrire.

## Curseurs et sélections en direct {#live-cursors-and-selections}

Les collaborateurs d'une même page voient le pointeur des autres (un curseur portant le nom de la personne dans sa couleur) et leurs formes sélectionnées (un contour en pointillés dans sa couleur) au-dessus du canevas. Ils suivent le zoom et le défilement du dessin et disparaissent quand la personne quitte la salle, change de page, retire son pointeur de la page, efface sa sélection ou devient inactive.

Chaque fenêtre publie `{ page, x?, y?, selection? }` : l'identifiant de la page, le pointeur en unités de dessin de la page (pouces) et jusqu'à 64 identifiants de formes sélectionnées. `sanitizeVisioPresence` (de `ooxml-core/visio/ui`) valide chaque charge entrante. Le visualiseur l'applique aux sessions qu'il ouvre lui-même. **Une `session` de l'hôte doit être créée avec lui** (avec `VISIO_INITIAL_PRESENCE` comme `initialPresence`), sinon ses pairs ne portent aucune charge et ni curseurs ni sélections ne sont dessinés (rien ne casse).

Limites : le curseur et la sélection ne sont affichés que pour la page ouverte par les deux personnes ; les contours de sélection suivent le cadre enregistré de la forme, pas sa géométrie dessinée.

## Écriture de retour du propriétaire {#owner-write-back}

Un hôte qui veut enregistrer le dessin partagé peut définir `onWriteBack` et, au besoin, `writeBackDebounceMs` (5000 par défaut) :

```ts
viewer.collaboration = {
	roomId: 'floor-plan',
	user: { name: 'Ada', role: 'owner' },
	onWriteBack: (bytes) => upload(bytes), // VSDX bytes
	writeBackDebounceMs: 5000,
};
```

Seule la fenêtre dont le rôle est `owner` écrit, si bien que les collaborateurs ne se disputent pas l'enregistrement. Le rappel se déclenche quand le dessin est resté calme pendant le délai (après des modifications locales et distantes), et une dernière fois quand le propriétaire quitte la salle si un enregistrement est en attente. Il ne se déclenche jamais pour `collaborator` ni `viewer`. Renvoyez une promesse depuis `onWriteBack` pour confirmer le stockage : une écriture à la fois est en cours, les modifications intermédiaires forment une écriture de suivi, et un rejet est réessayé si vous définissez `writeBackRetry` (`{ attempts, baseDelayMs?, maxDelayMs? }`). `collaborationState.writeBack` (aussi dans `collaboration-change`) indique `{ status: 'idle' | 'pending' | 'saving' | 'saved' | 'failed', lastSavedAt?, error? }` ; seul le propriétaire le renseigne. Un enregistrement en attente est vidé en quittant la salle, sur `pagehide` et quand la page est masquée ; une écriture asynchrone peut tout de même être perdue à la fermeture. Pour une écriture qui doit survivre à la fermeture, utilisez `navigator.sendBeacon` ou `fetch(..., { keepalive: true })` (limite d'environ 64 Kio).

## Ce qui existe aujourd'hui {#what-exists-today}

- Une session en direct fondée sur `ooxml-core/collab` (Yjs) via un `BroadcastChannel`, de sorte que les fenêtres et onglets d'un même navigateur peuvent partager un dessin ; les hôtes partagent entre appareils via `collaboration`.
- Les pages, les formes (jusqu'aux cellules, lignes et sections), les connecteurs, les masters et les images sont fusionnés élément par élément (`ooxml-core/visio/collab`) ; les autres parties du paquet (styles, thème, données, paramètres du document) sont partagées en bloc, la dernière écriture l'emporte par partie. Le runtime de collaboration n'est chargé qu'au démarrage du partage.
- Des modifications simultanées de formes différentes sont toutes deux conservées ; deux modifications d'une même forme donnent la même valeur dans toutes les fenêtres ; une forme ou une page supprimée l'emporte sur une modification simultanée.
- Annuler et rétablir (Ctrl+Z, le ruban) n'agissent que sur les modifications de cette fenêtre et n'annulent jamais celles d'un collaborateur. Les changements distants sont appliqués hors de votre historique, qui commence à l'arrivée dans la salle.
- Les collaborateurs apparaissent via l'élément de présence partagé de `ooxml-ui`.
- Le nom de session compte de 1 à 128 lettres, chiffres, traits d'union ou soulignés. Une fenêtre qui rejoint une session dont la salle contient déjà un dessin l'adopte ; sinon elle initialise la salle (sauf si l'hôte demande `intent: 'join'`).
- Les sessions du panneau conservent leur nom de BroadcastChannel d'origine (`ooxml-collab:visio-<room>`), de sorte que les fenêtres de versions plus anciennes se retrouvent.

## Démo en direct : deux fenêtres, n'importe quel framework {#live-demo-two-windows-any-framework}

Une session de partage n'est qu'un nom. Les deux volets ci-dessous sont des terrains de jeu distincts, construits avec des liaisons de frameworks différents, qui rejoignent au chargement la session `docs-collab` : la fenêtre A est le terrain de jeu vanilla et ouvre l'exemple, la fenêtre B est le terrain de jeu React et reçoit le dessin depuis la salle. Modifiez le texte d'une forme dans un volet et regardez l'autre suivre ; l'annulation inverse l'étape distante.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer fenêtre A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer fenêtre B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Associer deux frameworks au choix {#pair-any-two-frameworks}

Chaque terrain de jeu accepte `?share=<session>`, qui renseigne Fichier > Partager et démarre la session. Seul le nom de la session compte, donc n'importe quelle association fonctionne : la [démo en direct](/fr/#live-demo) de la page d'accueil propose un sélecteur de framework pour chaque fenêtre, et dans deux onglets vous pouvez ouvrir, par exemple, `/demo-vue/?sample=1&share=my-room` et `/demo-angular/?share=my-room`. Vous pouvez aussi le faire à la main dans deux fenêtres quelconques : Fichier, puis Partager, puis le même nom de session.

### Pourquoi Fichier > Partager reste limité à un même navigateur {#why-it-is-same-browser-only}

Le panneau utilise un `BroadcastChannel`, qui relie les fenêtres, onglets et cadres d'un profil de navigateur sur une même origine, quel que soit le framework, et rien d'autre. Deux machines différentes ne peuvent pas se rejoindre ainsi ; pour cela un hôte définit `collaboration` avec `serverUrl`, `provider` ou `session`, et cette version n'inclut aucun serveur.

## Ce qui est pris en charge et ce qui ne l'est pas {#what-is-and-is-not-supported}

| Pris en charge                                                                                       | Non pris en charge                                                                               |
| ---------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| Sessions dans un même navigateur via BroadcastChannel, fondées sur Yjs                               | Un serveur hébergé : apportez votre propre serveur y-websocket, fournisseur ou session           |
| Fusion par forme et par page, la suppression l'emporte, annulation locale uniquement                 | Fusion à l'intérieur d'une forme : deux modifications d'une même forme donnent une forme entière |
| Collaborateurs affichés avec l'élément de présence partagé                                           | Commentaires ; les rôles sont indicatifs (un serveur doit les faire respecter)                   |
| Fichier > Partager et la propriété `collaboration` comme points d'entrée                             | Curseurs entre pages différentes (seule la page partagée est dessinée)                           |
| Curseurs et sélections en direct des collaborateurs d'une même page                                  | Présence d'une `session` hôte créée sans `sanitizeVisioPresence`                                 |
| Écriture de retour du propriétaire : temporisée, confirmation par promesse, nouvelle tentative, état | Garantie qu'une écriture asynchrone se termine à la fermeture de la page                         |
| `authToken` pour un serveur et `webrtc` (votre propre `WebrtcProvider`)                              | y-webrtc ou serveur de signalisation inclus ; validation du jeton (rôle de votre serveur)        |

### Limites de la fusion au niveau des formes {#limits-of-shape-level-merging}

- Deux personnes qui modifient des propriétés différentes d'une même forme (un déplacement et un changement de texte ou de remplissage, des cellules, lignes ou sections différentes) conservent toutes deux leur changement. Si elles modifient la même cellule ou le même texte, il reste une seule valeur, identique dans toutes les fenêtres. Les modifications complètes d'un master et les enfants `Text` ou `Data` entiers forment chacun une unité.
- Si deux fenêtres ajoutent des formes avec le même identifiant en même temps, l'une est renumérotée ; les formules `Sheet.n!...` ajoutées plus tard ou par d'autres personnes qui citent l'ancien numéro ne sont pas réécrites.
- Une connexion (glue) supprimée par une personne et recréée à l'identique par une autre est conservée. Les images de médias supprimés restent dans la salle. Ce qu'une personne supprime (cellule, master, image) pendant qu'une autre le modifie conserve la modification ; une forme ou une page supprimée l'emporte toujours.
- Une salle écrite avec une autre version du schéma du format de collaboration est refusée : la fenêtre signale une erreur et ne lit ni n'écrit dans cette salle. Mettez à jour ensemble toutes les fenêtres d'une salle.
- L'historique est propre à chaque fenêtre, commence à l'arrivée et n'est pas conservé après une reconnexion. Chaque reconstruction relit tout le paquet de la salle (il n'y a pas d'analyse incrémentale) : un très grand dessin est relu à chaque lot de changements distants.

Les preuves sont consignées dans le [registre des capacités](/parity), sous « Collaboration and comments ».
