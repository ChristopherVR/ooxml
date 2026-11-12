---
title: Révision dans votre application
description: Une recette de révision collaborative de présentations dans une application hôte qui gère les comptes, les autorisations, la session, la persistance et la clé de l'IA.
---

# Révision dans votre application

La collaboration est généralement la plus utile au sein d'un flux de travail que vos utilisateurs possèdent déjà : une cellule commerciale, une file de conformité, une fiche CRM, un tableau de projet. Plusieurs personnes ouvrent la même présentation depuis ce système, font de petites modifications à la main ou avec l'assistant IA, et le système garde la trace de qui peut faire quoi et de la version en vigueur.

Cette recette répartit le travail selon une seule ligne de partage :

| Votre application gère                                                  | La visionneuse gère                                                       |
| ----------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| La connexion, les comptes et le nom et la couleur de l'utilisateur      | Le rendu et l'édition de la présentation                                  |
| Qui peut ouvrir, modifier ou seulement consulter chaque élément         | Le schéma des diapositives dans le document Yjs partagé                   |
| Le document Yjs, la conscience de présence, le fournisseur et le relais | La présence (curseurs, sélection, « suivre un collaborateur »)            |
| Le moment où la session est prête et celui où elle cesse de l'être      | Retenir les modifications tant que vous indiquez que la salle est prête   |
| Où vont les versions enregistrées et qui les écrit                      | Transformer la présentation partagée en octets PPTX pour le propriétaire  |
| Le modèle d'IA, sa clé et qui peut l'utiliser                           | Le panneau de l'assistant, ses outils de présentation et les propositions |

L'application d'exemple dans
[`demos/pptx/demo-host-review`](https://github.com/ChristopherVR/ooxml/tree/main/demos/pptx/demo-host-review)
implémente chacune des étapes ci-dessous avec la liaison vanilla. Chaque liaison accepte les mêmes options `collaboration` et `ai`, de sorte que les étapes se lisent de la même façon avec React, Vue, Angular et Svelte.

## 1. Faire correspondre vos permissions à un rôle {#_1-map-your-permissions-to-a-role}

Conservez votre propre modèle d'autorisations. Dérivez le `role` de la visionneuse de ce modèle lorsque vous montez l'éditeur :

| Votre permission (exemple) | `role`           | Remarques                                                      |
| -------------------------- | ---------------- | -------------------------------------------------------------- |
| Propriétaire de l'élément  | `'owner'`        | Modifie, et reçoit les instantanés `onWriteBack` à enregistrer |
| Relecteur                  | `'collaborator'` | Modifie en direct avec les autres                              |
| Observateur                | `'viewer'`       | Lecture seule ; passez aussi `editable: false`                 |
| Aucun accès                | (ne pas monter)  | Votre API ne doit même pas révéler que l'élément existe        |

`role: 'viewer'` n'est appliqué **que dans le navigateur**. Appliquez-le aussi côté serveur (étape 3).

## 2. Gérer la session vous-même {#_2-own-the-session}

Créez vous-même le document Yjs, la conscience de présence et le fournisseur, et passez-les en tant que `collaboration.externalSession` (voir [Collaboration gérée par l'hôte](/fr/guide/host-owned-collaboration) pour le contrat complet). La visionneuse ne se connecte jamais, ne se reconnecte jamais et ne les détruit jamais.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const grant = await api.requestGrant(reviewId); // your API: ticket, role, intent, user
const doc = new Y.Doc();
const awareness = new Awareness(doc);
const provider = new WebsocketProvider(`wss://${location.host}/collab`, reviewId, doc, {
	awareness,
	params: { ticket: grant.ticket },
});

const viewer = createPptxViewer(container, {
	source: sourceBytes, // the same source PPTX for every participant
	editable: grant.role !== 'viewer',
	collaboration: {
		roomId: reviewId,
		serverUrl: '', // unused with an external session
		userName: grant.user.name, // from your directory, not typed by the user
		userColor: grant.user.color,
		role: grant.role,
		sessionIntent: grant.intent,
		externalSession: { doc, awareness, getSnapshot, subscribe },
	},
});
```

`getSnapshot()` renvoie `{ status, synced }`. Ne signalez `synced: true` que lorsque le fournisseur est synchronisé **et** que l'utilisateur détient toujours l'accès avec lequel l'éditeur a été monté. La valeur `false` rend immédiatement l'éditeur en lecture seule.

**Qui initialise la salle.** Une salle vide est remplie par le premier participant ayant `sessionIntent: 'create'` ; tous les autres passent `'join'` et adoptent ce qui s'y trouve. Laissez votre serveur décider, pas les clients : l'exemple accorde `'create'` à un seul rédacteur à la fois, avec un bail court et uniquement tant que la salle est vide, afin que deux personnes ouvrant l'élément au même moment ne publient jamais chacune leur présentation de départ.

## 3. Admettre et contrôler les sockets côté serveur {#_3-admit-and-police-sockets-on-the-server}

Une WebSocket ne peut pas transporter le cookie de session de votre site vers un relais situé sur une autre origine, et un jeton de longue durée dans une URL finit dans les journaux. Émettez plutôt un **ticket signé à courte durée de vie** à partir d'un appel d'API ordinaire et authentifié, puis ne faites confiance au relais qu'à ce que le ticket indique :

- **À l'ouverture** : vérifiez la signature, l'expiration et la salle du ticket. Refusez tout le reste.
- **À chaque message** : une connexion en lecture seule peut demander le document (étape 1 de synchronisation) et partager la présence, mais ses messages d'étape 2 et de mise à jour sont ignorés. C'est la moitié serveur de `role: 'viewer'`.
- **Conscience de présence** : laissez chaque connexion ne parler que pour les identifiants de client de conscience qu'elle a revendiqués en premier, afin que personne ne puisse effacer ou usurper le curseur d'une autre personne.
- **Renouvellement du ticket** : un ticket ne fait qu'admettre une connexion. Renouvelez-le avant son expiration et mettez à jour `provider.params.ticket`, pour qu'une reconnexion automatique utilise un ticket valide.

Vous n'avez pas besoin d'écrire ce relais vous-même. `openteams-server` (le serveur de référence de la visionneuse OpenTeams, un serveur Yjs compatible y-websocket) accepte un hook `authorizeSync` qui renvoie `'write'`, `'read'` ou un refus pour chaque socket, fixe déjà les identifiants de conscience et expose `closeSync`, `syncPrincipals`, `isSyncRoomEmpty` et `handleUpgrade` pour les étapes ci-dessous. Le fichier `server/relay.ts` de l'exemple fait une soixantaine de lignes autour de lui.

## 4. Modifier les accès pendant que les personnes travaillent {#_4-change-access-while-people-are-working}

Lorsqu'un propriétaire modifie l'accès de quelqu'un, fermez les sockets ouvertes de cette personne avec un code de fermeture applicatif (l'exemple utilise `4403`). Le client réagit alors :

1. il signale `synced: false`, ce qui gèle l'éditeur ;
2. il détruit la visionneuse, puis son propre document et son fournisseur ;
3. il demande un nouvel accès et monte une session neuve avec le nouveau rôle, ou affiche « accès retiré » s'il n'y en a pas.

Repartez d'un **nouveau** Y.Doc après une rétrogradation. Les modifications que l'utilisateur a faites juste avant le changement ont été supprimées par le relais, et réutiliser l'ancien document continuerait de les afficher localement.

## 5. Enregistrer les versions avec un seul rédacteur {#_5-save-versions-with-one-writer}

Montez exactement un participant, en général le propriétaire de l'élément, avec `role: 'owner'` et `onWriteBack(bytes)`. La visionneuse temporise les modifications (`writeBackDebounceMs`, 5 secondes par défaut) et vous remet un instantané PPTX de la présentation partagée à conserver comme version. Les autres participants n'écrivent jamais en retour, il n'y a donc pas de course « le dernier enregistrement gagne ».

`onWriteBack` n'est pas un accusé de réception de stockage durable : les nouvelles tentatives, les échecs et le choix d'un rédacteur lorsque le propriétaire est hors ligne restent de votre ressort. Si la salle doit survivre au redémarrage du relais, persistez aussi les mises à jour Yjs côté serveur (par exemple avec votre propre stockage, ou un backend de fournisseur comme y-leveldb).

## 6. Ajouter des modifications assistées par l'IA {#_6-add-ai-assisted-changes}

Passez l'option `ai` uniquement aux comptes autorisés à l'utiliser, et pointez-la vers **votre** route afin que la clé du modèle n'atteigne jamais le navigateur :

```ts
ai: {
	connection: {
		kind: 'endpoint',
		api: `/api/ai/chat?review=${reviewId}`,
		credentials: 'same-origin',
	},
	writePolicy: 'stage', // every change is a proposal the reviewer accepts or rejects
},
```

La route vérifie le compte et l'autorisation de l'appelant, puis exécute le modèle avec les outils de présentation en lecture de schéma seulement, issus de `ooxml-ui/pptx/ai` (voir [Assistant IA](/fr/guide/ai-assistant)). Les appels d'outils s'exécutent dans le navigateur du relecteur, sur la présentation en direct. Une proposition acceptée est une modification ordinaire de ce relecteur : elle passe par l'historique de l'éditeur (un seul Ctrl+Z l'annule) et atteint la salle comme une frappe au clavier. Votre journal d'audit peut enregistrer que l'assistant a été utilisé ; le document partagé n'indique pas quelles modifications en proviennent.

## Lancer l'exemple {#run-the-example}

```sh
bun install                                   # at the repository root
cd demos/pptx/demo-host-review
bun run start                                 # http://localhost:4190
```

Ouvrez la page dans deux navigateurs (ou une fenêtre normale et une fenêtre privée) et connectez-vous avec des personnes différentes : Ana possède l'élément, Ben et Chen relisent, Dee observe. En tant qu'Ana, modifiez l'accès de Dee dans le panneau latéral et observez l'éditeur de Dee passer de la lecture seule à la modification, puis à l'absence d'accès, sans rechargement. Pour essayer l'assistant, définissez `REVIEW_AI_MODEL` (et `REVIEW_AI_API_KEY`, ou `REVIEW_AI_BASE_URL` pour un serveur local compatible OpenAI) avant de le démarrer.

## Limites {#limits}

- L'exemple garde tout en mémoire : comptes, versions et salles partagées disparaissent à l'arrêt du processus. Sa connexion est un sélecteur de compte.
- Le contenu de la conscience de présence (le nom affiché et la couleur) vient du client. Le relais fixe les identifiants de conscience aux connexions mais ne vérifie pas la charge utile par rapport au compte du ticket.
- Les commentaires, l'état d'approbation de la révision et l'attribution de chaque modification ne font pas partie du modèle partagé dans cet exemple ; validez ces flux séparément avant de vous y fier.
- La convergence de pairs navigateur n'est pas la même chose qu'une réouverture de la version enregistrée dans PowerPoint sans réparation ; vérifiez des présentations représentatives dans PowerPoint avant de le promettre.
