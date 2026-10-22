# Protocole de collaboration

Le composant web prend en charge deux modes de collaboration exclusifs : des étapes ProseMirror ordonnées par une autorité, et la fusion Yjs en option via une `CollabSession` détenue par l'application. Les deux utilisent le même éditeur et les six adaptateurs de frameworks. L'hôte gère le transport, l'authentification, la persistance et le moment de l'export.

## Démo en direct : deux fenêtres, n'importe quel framework

Une session n'est qu'un nom. Les deux volets ci-dessous sont des applications de démo distinctes, construites à partir de différents adaptateurs de frameworks (React et Vue), qui rejoignent la session `docs-collab` : le premier volet ouvre le document d'exemple et héberge la session, le second la rejoint par son nom et reçoit le document, chaque modification et le curseur de l'autre personne. Saisissez du texte dans l'un ou l'autre.

<iframe
	src="/ooxml/docx/demo/?sample=1&room=docs-collab&name=Ada"
	title="Hôte de session docx-viewer (React)"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/docx/demo-vue/?room=docs-collab&name=Grace"
	title="Invité de session docx-viewer (Vue)"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Associer deux frameworks quelconques

Chaque démo accepte `?room=<session>`. La fenêtre qui demande aussi `sample=1` héberge la session ; toute autre fenêtre avec le même `room` la rejoint, quel que soit le framework avec lequel elle a été construite. La [démo en direct](/fr/#live-demo) de la page d'accueil propose un sélecteur de framework pour chaque fenêtre, et dans deux onglets vous pouvez ouvrir, par exemple, `/demo-angular/?sample=1&room=my-room` et `/demo-svelte/?room=my-room`.

### Comment la démo procède, et pourquoi elle ne fonctionne que dans un seul navigateur

Le protocole de l'éditeur est neutre vis-à-vis du transport ; la démo fournit donc le transport : un `BroadcastChannel` nommé d'après la session. La fenêtre hôte exécute l'autorité de référence en mémoire (`createCollaborationAuthority`), ordonne les lots d'étapes de tous les participants et diffuse ceux qui sont acceptés ; un invité part du document initial de l'hôte et rejoue les lots acceptés, de sorte que le framework de chaque fenêtre n'a pas d'importance. Un `BroadcastChannel` relie les fenêtres, onglets et cadres d'un même profil de navigateur, et rien d'autre ; la session prend fin à la fermeture de la fenêtre hôte, et rien n'est persisté. Pour des personnes sur des machines différentes, votre application doit fournir le transport, l'autorité, l'identité et le stockage (voir plus bas).

Il existe aussi une [démo sur une seule page](/demo/collaboration.html){target="_self"} avec deux éditeurs dans la même page et un bouton **Pause delivery** pour essayer des modifications concurrentes ; ajoutez `?guest=vue` pour exécuter son second éditeur dans un autre framework.

## Ce qui est pris en charge, et ce qui ne l'est pas

| Pris en charge                                                                                           | Non pris en charge                                                                         |
| -------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| Lots d'étapes ProseMirror ordonnés et validés par une autorité                                           | Un serveur hébergé, le réseau, l'authentification ou la persistance (l'hôte en est chargé) |
| Fusion Yjs du texte et de la mise en forme, curseurs relatifs, fils de commentaires et annulation locale | Édition concurrente des styles, des définitions de numérotation et des notes hors du corps |
| Une autorité de référence en mémoire pour les tests et la démo locale                                    | Commandes structurelles de tableau pendant la collaboration                                |

## Mode Yjs

Chargez le même paquet DOCX source dans chaque éditeur, avec ses médias, styles, commentaires, notes et parties opaques. L'hôte fournit un `documentId` stable pour ce paquet et désigne exactement un créateur de salle. Cet identifiant est un contrat de l'hôte, pas un jeton d'authentification ni une somme de contrôle automatique du paquet.

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({
	roomId,
	provider: transportProvider({ transport }), // Application-owned transport
	user: { name: 'Ada', role: 'collaborator' },
});
if (!session.synced) {
	await new Promise<void>((resolve) => {
		const off = session.on('synced', (synced) => {
			if (synced) {
				off();
				resolve();
			}
		});
	});
}
editor.startYjsCollaboration(session, {
	documentId: sourcePackageId,
	initializeIfEmpty: designatedCreator,
});
editor.publishPresence({ name: 'Ada', color: '#2563eb' });
```

Attendez la véritable synchronisation initiale du fournisseur, et non la barrière d'écriture à délai de grâce, avant de rejoindre la salle. Le créateur alimente une salle vide ; les autres participants adoptent le corps existant et les attributs de page et de section au lieu de les écraser par un instantané local. Un identifiant de document ou un format de salle qui ne correspond pas fait rejeter l'adhésion. Une session d'autorité et une session Yjs ne peuvent pas piloter simultanément le même éditeur.

Le schéma Word partagé actuel initialise des salles `word-yjs-v3`. Chaque propriété directe modélisée d'une image, d'une référence de note, d'un marqueur de champ ou d'un saut possède un attribut indépendant au-dessus de sa base de propriétés importée. Les changements concurrents de propriétés différentes fusionnent, y compris la suppression d'une propriété ; l'annulation locale conserve les changements de l'autre auteur. Les conflits sur une même propriété suivent l'ordonnancement de Yjs. L'historique de mise en forme ne conserve encore qu'un instantané et qu'un auteur antérieurs par séquence de texte ; l'attribution complète des révisions concurrentes de plusieurs auteurs reste à terminer.

Ceci modifie la compatibilité des salles : la visionneuse mise à jour rejette les salles v1 et v2, et les anciens clients rejettent les salles v3. Les consommateurs du cœur qui utilisent d'anciens schémas peuvent encore employer des salles v1/v2 correspondantes, avec la seule rédaction de commentaires textuels. Pour mettre à niveau une salle persistée, exportez son DOCX actuel avec l'ancien client correspondant, chargez ce paquet dans les clients mis à jour et laissez le créateur désigné initialiser une nouvelle salle v3 vide. Maintenez tous les participants sur le schéma correspondant. L'adaptateur ne réécrit pas une salle active et ne migre pas automatiquement la persistance du fournisseur.

Le mode Yjs utilise la liaison stable pour Yjs 13, `y-prosemirror` 1.3.7. Les attributs racine du document disposent d'un mappage distinct, car la liaison ne les synchronise pas. Le texte, la mise en forme et les marques de révision suivies circulent dans le fragment du corps. Les parties d'images PNG, JPEG, GIF, BMP et SVG nouvellement insérées utilisent le routage d'actifs partagé du cœur avec des charges binaires brutes ; les noms sont uniques par client et les parties sont immuables. Les octets des images restent disponibles pour l'annulation/rétablissement et les exports après l'arrêt de la collaboration. Les actifs existants du paquet restent dans la source chargée correspondante de chaque éditeur. L'hôte contrôle toujours la synchronisation des exports et la persistance de la salle ; cela ne coordonne pas les enregistrements simultanés et ne prouve pas un export sans perte.

Le paramètre d'enregistrement du suivi des modifications circule lui aussi dans les transactions du document. Le basculer met à jour les pairs, participe à l'annulation/rétablissement local et s'exporte sous la forme `w:trackRevisions`. Rejoindre une salle compatible qui n'a pas ce paramètre conserve la valeur du paquet chargé jusqu'à ce qu'un participant la modifie. Les marques de révision conservent le nom de l'auteur de la modification et l'horodatage UTC de la transaction d'enregistrement. Les pairs conservent ces valeurs au lieu de réattribuer les modifications distantes ; les noms et horloges fournis par les clients n'authentifient ni l'auteur ni l'heure.

Ctrl/Cmd+Z et l'annulation du ruban utilisent l'historique Yjs local. L'édition ordinaire conserve aussi les enregistrements de fils et les ancrages dans des transactions annulables ; les instantanés locaux de fils sont exclus des attributs de document Yjs. Après l'arrêt de Yjs, ses fils actuels deviennent l'instantané de l'historique ordinaire. L'historique survit au détachement et au remontage de l'éditeur ; les vues déconnectées adoptent les changements au remontage. La présence (awareness) utilise des positions relatives pour le mappage des curseurs et des sélections, avec l'interface de curseur localisée existante. `publishPresence` met à jour le profil de présence et renvoie `null` en mode Yjs ; `leavePresence` efface le curseur. Le fournisseur transporte directement la présence, de sorte que `presence-send` et `receivePresence` restent des API du mode autorité.

`reconnectCollaboration()` redémarre le fournisseur Yjs en conservant le document partagé. `resyncCollaboration()` demande un échange d'état et renvoie `false` lorsque le fournisseur ne peut pas le faire. L'hôte peut aussi utiliser `session.connect()`, `disconnect()`, `peers()`, les abonnements d'état/synchronisation/erreur et `destroy()`. `stopCollaboration()` détache la liaison Word et conserve son contenu actuel et les nouveaux actifs d'images, y compris lorsqu'elle est démontée ; elle ne détruit pas la session détenue par l'hôte. Détruisez cette session lorsque l'hôte quitte la salle.

Les éditeurs en lecture seule et les rôles de lecteur reçoivent les mises à jour distantes tout en bloquant les écritures locales dans le document. Les rôles des clients sont indicatifs : appliquez aussi l'autorisation dans le serveur ou le fournisseur. L'édition hors du corps et les commandes structurelles de tableau restent désactivées. Les attributs de page et de section sont préservés et synchronisés, mais la collaboration multi-récits et un amorçage complet du paquet restent à terminer.

Les nouvelles salles Yjs partagent aussi les ancrages de commentaires, les réponses, la résolution et la suppression via le volet de révision existant. Le créateur alimente les commentaires chargés. Chaque commentaire possède des attributs de texte indépendants, des plages relatives à l'élément et un enregistrement ; les réponses concurrentes conservent leurs propres identifiants. La résolution utilise l'ordonnancement des conflits des cartes Yjs. Supprimer une racine masque ses réponses, y compris celles écrites en même temps hors ligne. L'annulation restaure ensemble la racine et son ancrage et révèle les réponses survivantes ; l'annulation de chaque auteur conserve les opérations des autres. L'export capture les derniers fils partagés même lorsque l'éditeur est détaché.

Les salles v3 portent une capacité de commentaire `inline-relative-v1` dans leur table d'identités Word. Les images, sauts, références de notes et marqueurs de champs conservent des ancrages indépendants malgré les ajouts et suppressions concurrents, les modifications du texte voisin et l'annulation locale. Les ancrages importés alimentent ces plages lors de la création d'une nouvelle salle. Commenter le résultat ou un marqueur d'un champ complexe équilibré couvre le champ entier, conformément au comportement natif de Word enregistré. Les salles héritées correspondantes conservent leur capacité `independent-v1` limitée au texte.
Les salles compatibles qui en sont dépourvues conservent leurs commentaires chargés et désactivent leur modification. Créez une nouvelle salle à partir d'un instantané canonique enregistré pour effectuer la mise à niveau ; ne changez pas la capacité d'une salle active. Le mode d'étapes d'autorité ne synchronise pas les métadonnées de commentaires, de sorte que la modification des commentaires y reste désactivée. Ces comportements client ne fournissent ni autorisation côté serveur, ni persistance durable, ni les mentions, notifications et attributions de tâches des commentaires modernes de Word.

Essayez la [démo Yjs à deux pairs](/demo/collaboration.html?mode=yjs){target="_self"}. Ses actions Pause delivery, Reconnect providers et Resync providers exercent la récupération en mémoire. Elle ne fournit ni réseau de production ni stockage durable.

## La propriété `collaboration`

Dans le cas courant, `<docx-editor>` accepte le même contrat d'intégration déclaratif que les éditeurs Excel, PowerPoint et Visio (`CollaborationOptions` dans `ooxml-core/collab` ; les alias Word sont `DocxCollaborationOptions` et `DocxCollaborationState`). C'est une fine couche au-dessus du mode Yjs décrit plus haut : vous n'avez ni à créer une session, ni à l'attendre, ni à choisir un `documentId` :

```ts
editor.collaboration = {
	roomId: 'q3-report',
	serverUrl: 'wss://example.com/collab',
	user: { name: 'Ada' },
	intent: 'create',
};
editor.addEventListener('collaboration-change', (event) => render(event.detail)); // { active, roomId, status, synced, people }
editor.collaboration = null; // or editor.stopCollaboration()
```

- Le transport est le premier parmi `session` (la vôtre, jamais détruite par l'éditeur), `provider`, `serverUrl`, sinon un BroadcastChannel entre onglets d'un même navigateur. Une session ouverte par l'éditeur lui-même est détruite quand il la quitte.
- `user.name` vaut par défaut `reviewAuthor` ; un `user.role` égal à `viewer` bloque les modifications locales (à titre indicatif, comme plus haut).
- `collaborationState` et `collaboration-change` rendent compte de toute liaison Yjs, y compris celle démarrée avec `startYjsCollaboration`. `reconnectCollaboration()` relance la connexion. Les échecs arrivent sous forme de `document-error`.
- `stopCollaboration()` sans argument efface aussi la propriété. L'argument `discardPending` du mode autorité est inchangé.
- Charger ou remplacer le document fait d'abord quitter la salle, puis la rejoint ensuite (l'API impérative lève toujours "Stop collaboration before replacing the document."). La jonction attend qu'un document soit monté et qu'il ne soit pas en cours de chargement.

Limites assumées :

- Le nom de la salle est le `documentId` Word. Rien ne vérifie que chaque fenêtre a chargé le même paquet source ; cela reste votre contrat.
- **Deux créateurs entrent en concurrence.** Avec la valeur par défaut `intent: 'create'`, une fenêtre initialise une salle vide. Si plusieurs fenêtres ouvrent une salle vide au même instant, chacune l'initialise et la salle contient finalement les deux corps. Donnez `create` à une seule fenêtre et `intent: 'join'` aux autres ; ce dernier n'initialise jamais et signale `document-error` si la salle est toujours vide après la synchronisation.
- Un créateur attend le premier échange avec la salle, ou l'ouverture de la barrière d'écriture (quelques secondes) lorsqu'il est seul sur un maillage sans serveur ; dans ce second cas, il initialise sans avoir vu la salle, si bien qu'un serveur lent dont le contenu arrive plus tard est initialisé deux fois. Utilisez un serveur ou `join` lorsque cela compte.
- Après avoir quitté puis rejoint la salle autour d'un chargement, une `session` détenue par l'hôte conserve le contenu antérieur de la salle, qui l'emporte sur le document nouvellement chargé ; une session détenue par l'éditeur repart de zéro.

## Protocole

Le composant web expose `startCollaboration`, `getPendingCollaboration`, `receiveCollaboration` et `stopCollaboration`. Les hôtes peuvent utiliser `createCollaborationAuthority(model, options)` pour créer l'autorité de référence en mémoire, avec le même schéma et la même conversion de modèle que l'éditeur.

## Flux côté client et côté hôte

Tous les participants qui éditent le même document utilisent le même `sessionId` ; chaque incarnation d'un éditeur utilise un `clientId` unique. Réutiliser un identifiant de client peut amener les pairs à prendre les étapes des autres pour des accusés de réception. Tous les participants doivent utiliser le même schéma et le même document initial. Pour rejoindre une autorité existante, initialisez le client avec sa version actuelle et son document canonique :

```ts
const client = new CollaborationClient({ sessionId, clientId, version: currentVersion });
const state = EditorState.create({ doc: canonicalDoc, plugins: [history(), client.plugin] });
```

Pour l'élément personnalisé, appelez `editor.startCollaboration({ sessionId, clientId, version })` après avoir chargé l'instantané de modèle correspondant. Gardez chaque récepteur monté tant qu'il participe : `receiveCollaboration` s'applique via la vue vivante de l'éditeur. Écoutez `collaboration-send` ; soumettez chaque `StepBatch` à l'autorité et remettez son lot accepté à chaque éditeur avec `receiveCollaboration(batch)`. Les rejets tels que `stale`, `out-of-order` ou `wrong-session` doivent déclencher une récupération au niveau de l'hôte plutôt que d'être ignorés silencieusement.

Après des transactions locales, appelez `client.createPendingBatch(state)`. Elle renvoie `null` s'il n'y a aucune étape non confirmée, et sinon une enveloppe immuable. Réessayez exactement cette enveloppe jusqu'à ce que l'autorité l'accepte ; ne lancez pas une autre requête tant qu'une est en cours. La méthode renvoie la même enveloppe jusqu'à la réception de son lot accepté correspondant.

L'hôte soumet l'enveloppe à son autorité et relaie le `batch` accepté à tous les participants, y compris l'expéditeur. La `version` d'un lot accepté est la version de l'autorité immédiatement avant les étapes du lot. Les participants transmettent l'enveloppe à `client.receive(state, batch)`. Lorsque le statut est `applied`, distribuez la transaction renvoyée par le chemin habituel de mise à jour d'état de l'éditeur. Ainsi, les mises à jour ordinaires du modèle, les événements d'état et le comportement d'enregistrement s'exécutent aussi pour les changements distants. L'application de lots distants ne consulte pas le réglage de lecture seule de la vue ; ce réglage ne bloque que les modifications locales.

Le protocole accepte au plus 500 étapes et 1 000 000 de caractères JSON sérialisés par lot. L'implémentation de référence `CollaborationAuthority` est une implémentation d'ordonnancement et de rebasage en mémoire. Elle accepte les requêtes de la version courante et fait passer les requêtes périmées à travers les étapes acceptées intermédiaires. Son historique par défaut conserve 2 000 étapes, et elle met en cache 5 000 accusés de réception de requêtes pour les nouvelles tentatives. Une requête périmée plus ancienne que l'historique conservé est rejetée ; une version future est rejetée comme hors séquence. Le client mémorise les 1 000 derniers identifiants de lots reçus pour ignorer les doublons. Les hôtes doivent prévoir un rattrapage ou un rechargement depuis leur instantané canonique après un trou, une version expirée ou un lot trop volumineux rejeté.

## Présence

`PresenceClient` fournit un canal transitoire distinct pour les noms, couleurs, curseurs et sélections des pairs. Il n'utilise ni ne modifie le flux d'étapes du document. Ajoutez son `.plugin` à côté du plugin de collaboration. `publish(state, { name, color })` renvoie un message pour le transport de l'hôte, ou `null` tant que des étapes locales du document attendent un accusé de réception. `leave(state)` renvoie un message de retrait. Après la réception d'un message, distribuez la transaction issue de `receive(state, message)` lorsque son statut est `applied` ; ignorez `duplicate` et récupérez après `stale`, `out-of-order` ou `wrong-session` selon la politique de l'hôte.

Les messages de présence portent l'identifiant de session partagé, l'identifiant du client expéditeur, la version de collaboration validée et une séquence strictement croissante par expéditeur. Relayez les messages dans l'ordre et republiez après l'accusé de réception d'une modification ; les positions sont définies par rapport à la version validée du message. Le plugin fait suivre aux curseurs et sélections existants les transactions ultérieures du document, y compris les modifications reçues pendant que l'éditeur est en lecture seule. Il accepte des noms de 80 caractères imprimables au plus, des couleurs issues de `PRESENCE_PALETTE` et au plus 100 pairs visibles. Les hôtes doivent limiter le débit et faire expirer les pairs déconnectés dans leur couche de transport. La présence n'est ni persistée, ni authentifiée, ni incluse dans l'export DOCX.

## Identités stables

Pour l'élément personnalisé partagé, la présence est activée par `startCollaboration` :

```ts
import { PRESENCE_PALETTE } from 'docx-vanilla-viewer';

editor.addEventListener('presence-send', (event) => {
	transport.sendPresence(event.detail); // Application-owned transport
});
editor.publishPresence({ name: 'Ada', color: PRESENCE_PALETTE[0] });
// Incoming messages: editor.receivePresence(message)
// Explicit departure: editor.leavePresence()
```

L'élément republie son profil enregistré après les changements de sélection et les accusés de réception de modifications. Il renvoie `out-of-order` pour les positions de sélection entrantes tant que des modifications locales sont en attente ; l'hôte peut demander une présence à jour après la synchronisation. Les messages de départ ne contiennent aucune position et peuvent retirer un pair à une autre version du document, tout en appliquant les contrôles de session et de séquence de l'expéditeur.

Les identifiants de paragraphes et de tableaux font partie du contenu du document et circulent dans les étapes ProseMirror. Pour les transactions locales, utilisez `repairCollaborativeDocumentIds(state, clientId, idGenerator)` plutôt que l'utilitaire de réparation d'identifiants non collaboratif. Créez `idGenerator` une seule fois avec `createCollaborationIdGenerator(clientId)` et réutilisez-le pendant toute la durée de vie de cet éditeur ; cela évite la réutilisation d'identifiants après suppression et place dans un espace de noms les identifiants de tableaux, de cellules et de paragraphes insérés localement. L'hôte doit attribuer un identifiant de client distinct à chaque incarnation active d'un éditeur ; l'identifiant de session et l'identifiant de client sont des identifiants de protocole, pas des justificatifs d'identité.

`stopCollaboration()` exige que les modifications en attente soient acquittées. Si l'hôte les abandonne volontairement, appelez `stopCollaboration(true)` pour ignorer explicitement la file de transport locale. Une déconnexion ou un démontage ne doit pas être considéré comme un accusé de réception ; conservez l'instantané et la version canoniques, puis reprenez ou récupérez selon la politique de transport de l'hôte.

## Responsabilités de l'autorité et limites

L'application hôte gère le transport, l'authentification et l'autorisation, le stockage durable, la distribution des instantanés, l'identité des clients, la politique de reconnexion et le moment de l'export. L'autorité en mémoire est utile pour les tests et une démo locale, pas comme serveur multiprocessus ou de production. L'export DOCX doit utiliser un modèle canonique suffisamment synchronisé si l'application a besoin de sémantiques d'enregistrement partagé ; cette bibliothèque ne coordonne ni les verrous de fichiers ni les exports simultanés. Les pairs distants ne reçoivent que des étapes ProseMirror prises en charge par le schéma, et l'autorité valide ces étapes avant de les accepter. Les commandes structurelles de tableau sont désactivées pendant la collaboration, car l'éditeur de tableaux actuel remplace la structure du tableau en bloc ; des transactions granulaires au niveau des cellules sont nécessaires avant de pouvoir activer ces commandes en toute sécurité.

Les liens pris en charge vers des images en ligne, des sauts et des notes sont portés par des attributs de séquence de texte, de sorte que leur import ou leur modification ne dépend pas de marques d'élément que Yjs abandonne. La boîte de dialogue Lien lit et modifie ces cibles via les mêmes commandes du cœur que pour les liens de texte ordinaires. Le changement de cible et la suppression sont synchronisés et prennent en charge l'annulation locale par l'auteur ; les médias d'images partagés restent disponibles dans les exports DOCX. Les liens d'équations et l'attribution des révisions de liens hypertexte restent hors de cette prise en charge.
