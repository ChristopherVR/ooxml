---
title: Collaboration gérée par l'hôte
description: Fournir un document Yjs, une conscience et un état de synchronisation existants, sans donner à la visionneuse la propriété du transport.
---

# Collaboration gérée par l'hôte

Utilisez `collaboration.externalSession` lorsque votre application gère déjà l'authentification, un fournisseur Yjs, le stockage hors ligne ou une connexion partagée. React, Vue, Angular, Svelte et Vanilla acceptent le même contrat. La visionneuse gère son schéma de diapositives et la présence ; l'hôte gère la connexion, la reconnexion, l'autorisation et la durée de vie des ressources.

```ts
import type { CollaborationConfig, ExternalCollaborationSession } from 'pptx-react-viewer';

// These resources and the store belong to your application, not to a viewer mount.
const session: ExternalCollaborationSession = {
	doc: hostDocument,
	awareness: hostAwareness,
	getSnapshot: () => connectionStore.getSnapshot(),
	subscribe: (listener) => connectionStore.subscribe(listener),
};

const collaboration: CollaborationConfig = {
	roomId: 'presentation-123',
	serverUrl: '', // Required by the shared config, but unused for an external session.
	userName: 'Alice',
	sessionIntent: 'join',
	externalSession: session,
};
```

Importez les types depuis le package de votre framework. Passez `collaboration` à sa prop ou option habituelle de visionneuse. Le document doit être un véritable `Y.Doc` ; la conscience (awareness) doit appartenir à ce document. Gardez l'objet de session stable et notifiez les abonnés lorsque `status` ou `synced` changent. Un abonnement renvoie une fonction de désabonnement.

L'hôte et la visionneuse doivent résoudre le même runtime Yjs. La visionneuse laisse Yjs en dehors de son bundle et le déclare comme peer optionnelle. Installez `yjs` dans l'application pour la collaboration intégrée comme pour la collaboration gérée par l'hôte :

```sh
npm install yjs
```

La visualisation et l'édition ordinaires sans collaboration ne nécessitent pas cette peer. L'hôte installe aussi le fournisseur et l'implémentation de conscience de son choix. Dédoublonnez Yjs si votre arbre de dépendances contient plusieurs versions : les types issus de copies d'exécution distinctes ne peuvent pas être mélangés en toute sécurité.

## Utiliser le contrat de collaboration partagé {#using-the-shared-collaboration-contract}

Word, Excel et Visio acceptent une seule forme d'option `collaboration` provenant de `ooxml-core/collab` (`CollaborationOptions` : `roomId`, `serverUrl`, `provider`, `session`, `user`, `intent`). PowerPoint accepte la même forme dans sa prop `collaboration`, aux côtés de son propre `CollaborationConfig`, afin qu'un hôte qui intègre plusieurs produits puisse leur donner une seule session commune :

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({ roomId, provider: transportProvider({ transport }), user });
// <PowerPointViewer collaboration={{ session, user: { name: 'Ada' }, intent: 'join' }} />
```

- `session` est empruntée : elle devient le `externalSession`, et la visionneuse ne la détruit jamais. L'hôte la crée, la connecte et la détruit. `roomId` vaut par défaut la salle de la session.
- `user.name`, `color`, `avatar` et `role` correspondent respectivement à `userName`, `userColor`, `userAvatar` et `role` ; `intent` correspond à `sessionIntent`. Lorsque les deux sont renseignés, les champs explicites hérités l'emportent.
- Chaque binding exécute le même `resolveCollaborationConfig` (depuis `ooxml-ui/pptx`) à l'endroit où il lit la prop : ils se comportent donc de façon identique. Un `CollaborationConfig` hérité est renvoyé sans modification.

Sans `session`, les mêmes options couvrent aussi les cas pour lesquels un hôte ne souhaite pas construire de session :

```ts
// Your own transport: the viewer opens and destroys the session around it.
collaboration = { roomId, provider: transportProvider({ transport }), user: { name: 'Ada' } };
// No server: tabs and windows of this browser share over a BroadcastChannel.
collaboration = { roomId, user: { name: 'Ada' } };
```

Le transport est choisi dans l'ordre partagé : `session` (empruntée, jamais détruite) > `provider` > `serverUrl` > BroadcastChannel. Pour `provider` et la salle BroadcastChannel, la visionneuse appelle `openCollaborationSession` (le runtime `ooxml-core/collab` n'est importé qu'à ce moment-là, en partageant l'unique copie de Yjs), présente la session par le même chemin de session empruntée qu'une `session` fournie par l'hôte, et la détruit à l'arrêt, au démontage ou lorsque la configuration change. La session devient modifiable une fois synchronisée, ou après le délai de grâce accordé à un pair seul sur un maillage ; `intent: 'join'` attend toujours la présentation de la salle, et `role: 'viewer'` reste en lecture seule. L'état passe à `connecting` pendant l'ouverture de la session, et à `error` pour un nom de salle invalide ou un navigateur sans BroadcastChannel.

Un `CollaborationConfig` hérité conserve ses transports : `serverUrl` sélectionne y-websocket, et `transport: 'webrtc'` ou un `serverUrl` vide sélectionne y-webrtc sans serveur (qui synchronise aussi les onglets du même navigateur). Seule une entrée du contrat partagé qui ne nomme ni `serverUrl` ni `transport` utilise la salle BroadcastChannel. La présence de la visionneuse et le battement de cœur d'identité de la session partagent le champ `presence` de la conscience ; la visionneuse conserve sa charge utile de diapositive et de curseur à travers les battements, de sorte que les curseurs distants restent affichés.

## Signaler l'état partagé {#reporting-the-shared-state}

Chaque binding signale le même `CollaborationState` que Word, Excel et Visio (`active`, `roomId`, `status`, `synced`, `people`, la personne locale marquée `self` ; le type provient de `ooxml-core/collab` et est aussi réexporté par chaque package). Il se déclenche lorsqu'une session démarre ou s'arrête (état désactivé, `status: 'off'`) et lorsque le statut, l'indicateur de synchronisation ou les personnes changent, pour les sessions gérées par l'hôte comme pour celles gérées par la visionneuse, et ne répète jamais un état inchangé :

| Binding | Notification                                                |
| ------- | ----------------------------------------------------------- |
| React   | prop `onCollaborationChange`                                |
| Vue     | événement `collaboration-change`                            |
| Angular | sortie `collaborationChange`                                |
| Svelte  | prop de rappel `oncollaborationchange`                      |
| Vanilla | option `onCollaborationChange` et `getCollaborationState()` |

Vanilla conserve `onCollaborationStatus` et `getCollaborationStatus()` pour le statut brut de la connexion.

## La disponibilité relève de la décision de l'hôte {#readiness-is-the-host-s-decision}

`getSnapshot()` renvoie `{ status, synced }`. Le statut est `connecting`, `connected`, `disconnected` ou `error`. `synced` signifie que le document faisant autorité est disponible et que la visionneuse peut publier des modifications. Il est indépendant de l'état du réseau :

- Gardez `synced: false` pendant le chargement ou l'hydratation du document partagé. Aucun délai d'attente n'autorise les écritures à la place de l'hôte.
- Passez `synced: true` après la synchronisation initiale. Les diapositives existantes de la salle sont adoptées avant qu'une présentation de démarrage locale puisse être publiée.
- Gardez-le à true pendant une période hors ligne si les modifications peuvent être mises en file d'attente dans le fournisseur Yjs de l'hôte.
- Passez-le à false pour rendre l'éditeur en lecture seule et suspendre l'écriture différée durable pendant une nouvelle synchronisation. L'édition reprend une fois l'état authentique de la salle adopté.

Le texte en ligne accepté et les aperçus de gestes sont écrits de façon synchrone dans le document emprunté ; un changement ultérieur de disponibilité ne peut donc pas effacer une dernière frappe mise en file. Le fournisseur de l'hôte peut toujours regrouper le trafic réseau. Une modification de texte active se ferme lorsque la disponibilité est retirée ; la reprise adopte la salle faisant autorité au lieu de rejouer un brouillon local périmé.

Utilisez `sessionIntent: 'create'` pour le participant autorisé à amorcer une salle vide, et `'join'` pour les participants dont la présentation de démarrage ne doit pas l'amorcer. Un Fichier > Ouvrir explicite reste un remplacement intentionnel. Coordonnez la création de la salle dans l'hôte : il ne s'agit pas d'une autorisation côté serveur ni d'un protocole d'élection.

## Propriété et persistance {#ownership-and-persistence}

La visionneuse ne crée jamais un second fournisseur, ne déconnecte jamais le fournisseur de l'hôte et ne détruit jamais le document ou la conscience fournis. Au démontage, elle retire ses écouteurs et ne restaure que le champ de présence qu'elle possède encore. Les champs de conscience non liés et les écritures ultérieures de l'hôte sont préservés. Utilisez un seul rédacteur actif de présence de diapositive par instance de conscience.

L'hôte peut persister directement les mises à jour Yjs. Sinon, désignez un participant avec `role: 'owner'` et `onWriteBack(bytes)` pour recevoir des instantanés PPTX à débit limité, en utilisant les mêmes PPTX source et options d'enregistrement que l'éditeur. Les instantanés en attente et en cours sont annulés si la session quitte ou si l'hôte retire la disponibilité. La désignation du rédacteur reste de la responsabilité de l'hôte. Renvoyez une promesse depuis `onWriteBack` pour accuser réception du stockage : la visionneuse signale `saving`, `saved` ou `failed` dans `collaborationState.writeBack` (propriétaires uniquement), sérialise les écritures (une seule en cours, les changements ultérieurs fusionnent en un seul suivi) et réessaie un rejet selon `writeBackRetry` (`{ attempts, baseDelayMs?, maxDelayMs? }`) ; un simple retour compte immédiatement comme accusé. Une écriture différée en attente est vidée sur `pagehide` et lorsque la page passe à l'état masqué. Les navigateurs n'attendent pas les travaux asynchrones lors du déchargement d'une page : une écriture qui nécessite un aller-retour réseau peut donc encore être perdue à la fermeture. Un hôte dont l'écriture `write` doit survivre à une fermeture doit envoyer les octets avec `navigator.sendBeacon(url, bytes)` ou `fetch(url, { method: 'PUT', body: bytes, keepalive: true })`, que le navigateur termine après le départ de la page (les corps keepalive sont plafonnés à environ 64 KiB, un document volumineux peut donc encore ne pas aboutir).
Une session active avec `role: 'viewer'` rend l'éditeur en lecture seule et bloque la publication, pour la collaboration intégrée comme pour celle gérée par l'hôte. Le serveur doit appliquer lui-même le contrôle d'accès.

## Limites d'intégration {#integration-boundaries}

Cette option modifie la propriété de la session, pas le modèle de données de collaboration existant. La visionneuse possède `pptx:slides` et ses types Yjs imbriqués. Donnez à tous les participants le même PPTX source pour les ressources du package, les thèmes et les médias. D'autres données applicatives peuvent occuper des clés Yjs distinctes. Ne réconciliez pas indépendamment, depuis l'hôte, des instantanés complets de diapositives.

Vérifiez les limites de messages et de persistance de l'hôte sur des présentations représentatives, y compris la mise à jour initiale du document et `Y.encodeStateAsUpdate(doc)` après modification. La taille d'un fichier PPTX compressé n'est pas la taille de son instantané Yjs : le schéma partagé transporte aussi le XML analysé, les métadonnées de formes et de groupes et les charges utiles des ressources. Sortir les images du document ne suffit peut-être pas. Cette API de session n'ajoute ni découpage de transport ni modification des limites de l'hôte.

Elle n'ajoute pas d'intégration au backend de commentaires, de nouvelle implémentation d'annulation locale seule, ni de synchronisation de chaque champ au niveau de la présentation. Validez séparément ces flux de travail avant de les exposer dans un produit collaboratif. Une réouverture native dans PowerPoint est aussi une vérification de fidélité distincte de la convergence entre deux pairs de navigateur.

## Essayer les démonstrations {#trying-the-demos}

Démarrez le relais de démonstration (`bun run collab`) et une démonstration de framework, puis ouvrez :

```text
/?externalSession=1&room=example&sample=1&server=ws://localhost:1234
```

Un second participant utilise la même salle sans `sample=1`. La bande de l'hôte expose les commandes de disponibilité et de montage de l'éditeur tout en conservant le document et le fournisseur. La démonstration désactive BroadcastChannel afin que la synchronisation utilise le relais WebSocket.

Ajoutez `&headless=1` pour essayer une interface personnalisée dans l'une des cinq démonstrations. Ces exemples utilisent les blocs de construction publics du binding sans monter une visionneuse complète et cachée. Leur bouton Enregistrer géré par l'hôte appelle le chemin public de sérialisation ; le sélecteur de zoom permet de vérifier les curseurs distants et les contours de sélection à différentes échelles.

Voir [Hooks React](/react/hooks), [Collaboration Vue](/vue/collaboration), [Collaboration Angular](/angular/collaboration), [Collaboration Svelte](/svelte/collaboration) et [Interfaces personnalisées Vanilla](/vanilla/api#custom-shell-collaboration) pour chaque composition.

Pour une application complète autour de cette option (comptes, relais avec contrôle des autorisations, modifications d'accès en direct, versions du propriétaire et route IA contrôlée), voir [Relecture dans votre application](/fr/guide/review-workflow).
