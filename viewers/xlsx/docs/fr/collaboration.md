# Collaboration

Plusieurs personnes peuvent modifier un même classeur en même temps. L’éditeur associe le classeur à un document Yjs via `ooxml-core/xlsx/collab` et rejoint une salle grâce à la session partagée `ooxml-core/collab` (la même session et les mêmes fournisseurs que les éditeurs Word et Visio). Chaque modification est envoyée au fil de l’eau, les modifications simultanées de cellules différentes fusionnent, et chacun voit où se trouvent les autres.

## Démarrer le partage

Depuis l’éditeur : **Fichier > Partager** (ou le bouton Partager à l’extrémité droite de la rangée d’onglets du ruban), saisissez un nom de session et choisissez **Démarrer le partage**. Une autre fenêtre ou un autre onglet du même navigateur qui a un classeur ouvert et démarre le partage avec le même nom de session la rejoint. Sans serveur défini par la page, rien ne quitte le navigateur : la session passe par un `BroadcastChannel`.

Depuis le code, définissez la propriété `collaboration` (ou la prop, dans chaque liaison) :

```ts
editor.newWorkbook(); // a window joining an existing room still needs a workbook open
editor.collaboration = {
	roomId: 'budget-2026',
	serverUrl: 'wss://collab.example.com/rooms', // y-websocket compatible; omit for same-browser only
	user: { name: 'Ada', color: '#2563eb' },
};
editor.addEventListener('collaboration-change', (event) => console.log(event.detail));
```

| Option      | Remarques                                                                                                                                      |
| ----------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| `roomId`    | De 1 à 128 lettres, chiffres, tirets ou tirets bas.                                                                                            |
| `serverUrl` | URL de base d’un serveur compatible y-websocket ; l’identifiant de salle y est ajouté.                                                         |
| `provider`  | Une fabrique de fournisseur de votre cru (signalisation WebRTC, ou un fournisseur Yjs existant via `externalProvider`). Prime sur `serverUrl`. |
| `session`   | Une `CollabSession` existante créée avec `sanitizePayload: sanitizeXlsxPresence`. L’éditeur ne détruit jamais une session qu’on lui a fournie. |
| `user`      | `name` (par défaut `authorName`), `color` (par défaut une couleur de la palette), `role` (`viewer` met l’éditeur en lecture seule).            |

Tant que `collaboration` est définie, un classeur ouvert ensuite rejoint la même salle ; quand la salle contient déjà un classeur, son contenu remplace celui qui a été ouvert. La définir à `null` quitte la salle. Les formes impératives sont `startCollaboration(options)` (se résout une fois la salle rejointe), `stopCollaboration()`, `reconnectCollaboration()`, `collaborationState` et `share()` (ouvre Fichier > Partager). Chaque liaison transmet `collaboration`, signale `collaboration-change` (`onCollaborationChange`, Vue `collaboration-change`, Angular `collaborationChange`, Svelte `oncollaborationchange`) et expose `share()` sur son handle ; démonter l’éditeur quitte la salle.

## Ce qui est partagé

- Les valeurs des cellules, les formules (avec leurs valeurs en cache), les formats numériques et la mise en forme des cellules.
- Les feuilles : ajout, renommage, déplacement, suppression, masquage et couleur d’onglet.
- Les hauteurs de ligne, largeurs de colonne, lignes et colonnes masquées, cellules fusionnées et noms définis.
- La sélection de chaque personne, dessinée dans la grille des autres sous forme de contour à sa couleur avec son nom, et les personnes de la session sous forme d’avatars dans la barre de titre.

Annuler (Ctrl+Z, le ruban et la barre d’outils Accès rapide) n’annule que vos propres modifications tant que le classeur est partagé ; le travail d’un collaborateur n’est jamais annulé. Quand la connexion tombe, l’éditeur le signale et conserve vos modifications ; elles sont envoyées au retour de la connexion, et un second avis confirme la reconnexion.

## Ce qui n’est pas partagé

Les graphiques, images et autres dessins, commentaires, liens hypertexte, mises en forme conditionnelles, validations de données, tableaux et filtres, tableaux croisés dynamiques, volets figés et autres paramètres d’affichage, mises en page, protections, propriétés du document et paramètres de calcul restent propres à chaque fenêtre et sont enregistrés depuis celle-ci. Les insertions et suppressions de lignes et de colonnes circulent sous forme des cellules qu’elles déplacent ; une modification simultanée d’une cellule déplacée atterrit donc à son ancienne adresse. Une formule saisie par un collaborateur avec le nom d’une feuille renommée au même moment conserve l’ancien nom. Les règles de résolution des conflits sont décrites dans le dépôt `ooxml` (`docs/collab-area.md`).

## Ce que l’hôte prend en charge

L’éditeur ne fournit aucun serveur. Votre application gère le transport (l’URL du serveur ou le fournisseur), l’authentification et l’autorisation (les rôles sont indicatifs côté client), la persistance des salles, ainsi que le moment et la fenêtre depuis lesquels le classeur est enregistré : l’éditeur ne coordonne pas les enregistrements simultanés. Ce n’est pas la co-édition d’Excel et il ne prétend pas à la parité avec elle.
