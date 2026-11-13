# API de l'élément et événements

Chaque adaptateur de framework encapsule le même élément personnalisé `<docx-editor>`. Les adaptateurs transmettent ces propriétés et réémettent ces événements selon les conventions de leur framework (voir les [adaptateurs de frameworks](/bindings)).

## Attributs et propriétés

| Propriété                      | Attribut          | Description                                                                                                                                                                                                                                                                        |
| ------------------------------ | ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `documentModel`                |                   | Le `DocumentModel` affiché. L'assigner est un remplacement externe, pas un événement d'édition.                                                                                                                                                                                    |
| `readOnly`                     | `read-only`       | Bloque les modifications locales. Les lots de collaboration distants s'appliquent toujours.                                                                                                                                                                                        |
| `locale`                       | `locale`          | Langue de l'interface : `en`, `fr`, `de`, `es` ou `zh-CN`. Le contenu du document n'est jamais traduit.                                                                                                                                                                            |
| `fileName`                     | `file-name`       | Nom du document ouvert (`Document1.docx` par défaut).                                                                                                                                                                                                                              |
| `reviewAuthor`                 | `review-author`   | Auteur enregistré sur les modifications suivies et les commentaires (`Author` par défaut).                                                                                                                                                                                         |
| `theme`                        | `theme`           | `light`, `dark` ou `auto`. Voir les [thèmes](/theming).                                                                                                                                                                                                                            |
| `themeColors`                  |                   | Remplacements partiels de jetons, appliqués comme propriétés personnalisées `--dve-*`.                                                                                                                                                                                             |
| `showThumbnails`               | `show-thumbnails` | Rail de vignettes de pages à gauche (mise en page Impression uniquement).                                                                                                                                                                                                          |
| `showToolbar`                  | `show-toolbar`    | `false` masque le ruban ; la barre de titre et la barre d'état restent affichées.                                                                                                                                                                                                  |
| `hiddenActions`                |                   | Commandes du ruban à masquer par identifiant stable ; voir le tableau dans les [adaptateurs](/bindings).                                                                                                                                                                           |
| `ribbonAddIns`                 |                   | Onglets ajoutés par l'hôte après les onglets intégrés ; voir les [adaptateurs](/bindings#ribbon-add-in-tabs).                                                                                                                                                                      |
| `collaboration`                |                   | Collaboration déclarative (`DocxCollaborationOptions` ou `null`) ; voir [collaboration](/fr/collaboration#la-propriete-collaboration).                                                                                                                                             |
| `customization`                |                   | Capacités, pages de l’hôte dans l’onglet Fichier et commandes de menu de l’hôte (voir [Personnalisation](#customization)).                                                                                                                                                         |
| `surfaceOnly`                  |                   | Monte uniquement le document et sa surface d'édition. À choisir avant que l'élément ne soit connecté.                                                                                                                                                                              |
| `surfaceActions`               |                   | Dialogues de l'hôte pour le mode surface : `onOpenHyperlinkDialog`, `onOpenFindReplace` et `onOpenImagePicker`.                                                                                                                                                                    |
| `googleFonts`                  |                   | Polices web pour les polices du document qui ne sont ni installées ni intégrées. Non défini, suit Fichier > Options > Centre de gestion de la confidentialité (désactivé tant qu'on ne l'active pas) ; `false` les désactive ; `{ baseUrl }` définit le point de terminaison CSS2. |
| `storage`                      |                   | Stockage géré par l'hôte pour Enregistrer, Enregistrer sous et Ouvrir ; voir « Stockage géré par l'hôte » plus bas.                                                                                                                                                                |
| `fileProviders`                |                   | Emplacements cloud pour Fichier > Ouvrir ; voir « Emplacements cloud dans Fichier > Ouvrir » plus bas.                                                                                                                                                                             |
| `loading` (lecture)            |                   | Vrai pendant le chargement d'un document.                                                                                                                                                                                                                                          |
| `collaborationState` (lecture) |                   | La salle de collaboration : si elle est active, son identifiant, son statut, son état de synchronisation et les personnes présentes.                                                                                                                                               |
| `dirty` (lecture)              |                   | Vrai après une modification, faux après un enregistrement, un chargement ou `markClean()`.                                                                                                                                                                                         |

## Personnalisation {#customization}

`customization` accepte un objet `WordCustomization`, le contrat partagé de `ooxml-core/embedding` que les autres éditeurs utilisent aussi. Chaque liaison le transmet comme prop ou entrée `customization`. Il est appliqué quand l’identité de l’objet change ; gardez-le donc stable. `null` rétablit les valeurs par défaut.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [
			{ id: 'send', label: 'Send to chat', onSelect: ({ selectedText }) => send(selectedText) },
		],
	},
};
```

- `capabilities` : `false` supprime tous les chemins vers la capacité (`edit`, `print`, `export`, `save`, `open`, `share`, `copy`, `comment`, `review`, `find`, `settings`) : ruban, pages de l’onglet Fichier, menus et raccourcis. `hiddenActions` est inchangé ; les deux sont réunis.
- `backstage.extraPages` : pages de l’hôte `{ id, label, group?, order?, render }` dans l’onglet Fichier, après les pages intégrées de leur groupe. `render` est une fonction de montage, une balise d’élément personnalisé ou `{ element, properties }` et reçoit un `WordRegionContext`.
- `contextMenu.extraCommands` : commandes de l’hôte `{ id, label, icon?, group?, when?, disabled?, onSelect }` au-dessus (`top`, par défaut) ou au-dessous (`bottom`) des entrées intégrées. `onSelect` s’exécute après la fermeture du menu avec un `WordMenuContext`. Les libellés sont affichés tels quels (localisez-les vous-même).

`edit: false` est un verrouillage en lecture seule à part entière, à côté de l’interrupteur de lecture seule de l’hôte et d’un rôle de collaboration `viewer` : aucun ne lève un autre. Les capacités sont une politique d’interface, pas une protection du document. Le remplacement de régions (boîte d’options, volets) n’est pas encore proposé pour Word. Tous les produits réunis : [guide d’intégration](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Méthodes

| Méthode                                                         | Description                                                                                                                                                       |
| --------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes)`                                                   | Charge des octets DOCX ou DOC hérité. Rejette et émet `document-error` en cas d'échec.                                                                            |
| `save()`                                                        | Se résout en un `Blob` du document sérialisé.                                                                                                                     |
| `saveBytes()`                                                   | Se résout en le `Uint8Array` brut.                                                                                                                                |
| `download(fileName?)`                                           | Enregistre et lance un téléchargement dans le navigateur.                                                                                                         |
| `markClean()`                                                   | Efface l'indicateur dirty une fois le document persisté par l'hôte.                                                                                               |
| `setLoadedDocument(session)`                                    | Monte un `LoadedDocument` déjà analysé par `loadDocument`.                                                                                                        |
| `openCloudFile(file)`                                           | Ouvre un fichier d'un `CloudFileProvider`, comme le fait Fichier > Ouvrir > OneDrive.                                                                             |
| `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` | État et rappels pour les contrôles de barre d'outils, de navigation et d'état construits par l'hôte.                                                              |
| `subscribeEditorState(listener)`                                | Appelle `listener` lors des changements de sélection et de mise en forme, y compris ceux qui ne modifient pas le document. Renvoie une fonction de désabonnement. |
| `startCollaboration(config)`                                    | Rejoint une session de collaboration ; voir la [collaboration](/fr/collaboration).                                                                                |
| `startYjsCollaboration(session, options)`                       | Rejoint une salle Yjs synchronisée avec son paquet source correspondant.                                                                                          |
| `reconnectCollaboration()`, `resyncCollaboration()`             | Rétablit la connexion Yjs ou demande un échange d'état au fournisseur.                                                                                            |
| `getPendingCollaboration()`                                     | Le lot d'étapes en attente, ou `null`.                                                                                                                            |
| `receiveCollaboration(batch)`                                   | Applique un lot accepté par l'autorité.                                                                                                                           |
| `stopCollaboration(discardPending?)`                            | Quitte la session ; les modifications en attente doivent être acquittées ou explicitement ignorées.                                                               |
| `publishPresence`, `receivePresence`, `leavePresence`           | Canal de présence transitoire ; voir la [collaboration](/fr/collaboration#présence).                                                                              |

## Événements

Tous les événements remontent (bubble) et sont composés, de sorte qu'ils traversent la frontière du shadow DOM.

| Événement              | Détail                               | Signification                                                                                          |
| ---------------------- | ------------------------------------ | ------------------------------------------------------------------------------------------------------ |
| `document-change`      | `DocumentModel`                      | Le document a changé par l'édition ou par des actions de révision.                                     |
| `document-error`       | `Error`                              | Le chargement, l'enregistrement, l'impression ou une modification a échoué.                            |
| `document-warning`     | `string`                             | Avertissement non fatal.                                                                               |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }` | L'utilisateur a basculé Édition/Lecture dans l'interface.                                              |
| `file-command`         | `{ command, fileName? }`             | Annulable. `command` vaut `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` ou `share`. |
| `dirty-change`         | `boolean`                            | L'état « modifications non enregistrées » a basculé.                                                   |
| `page-change`          | `{ page, pageCount }`                | La page ou le nombre de pages de la mise en page Impression a changé (pagination approximative).       |
| `ribbon-action`        | `RibbonAction`                       | Une commande du ruban a été activée.                                                                   |
| `ribbon-customize`     | `readonly string[]`                  | L'utilisateur a modifié les commandes affichées du ruban ; persistez les identifiants si besoin.       |
| `office-ribbon-add-in` | `{ tab, command }`                   | Une commande d'un onglet de l'hôte (`ribbonAddIns`) a été choisie.                                     |
| `presence-send`        | `PresenceMessage`                    | La présence locale a changé ; transmettez-la aux autres clients.                                       |
| `collaboration-send`   | `StepBatch`                          | Des étapes locales sont en attente ; transmettez-les à l'autorité.                                     |
| `collaboration-change` | `DocxCollaborationState`             | L'état de la collaboration a changé : `{ active, roomId, status, synced, people }`.                    |

Le paquet exporte la table typée des événements (`DocxEditorEventMap`) et la liste d'exécution `DOCX_EDITOR_EVENTS`.

Les liaisons de framework transmettent les props `fileName`, `reviewAuthor` et `themeColors`, ainsi que les événements `file-command`, `document-warning`, `readonly-change`, `ribbon-action` et `ribbon-customize`, sous les noms propres à leur framework (par exemple `onFileCommand`, `fileCommand` et `onfilecommand`). La table se trouve dans [liaisons de framework](/bindings#binding-props-and-events).

## Parties shadow

La barre de titre (nom du fichier, état d'enregistrement, accès rapide et « Dites-moi ») est la partie shadow `title-bar`.
Un hôte qui dessine sa propre barre de document la masque avec `docx-editor::part(title-bar) { display: none; }` ;
le ruban et la barre d'état restent affichés.

## Emplacements cloud dans Fichier > Ouvrir

`fileProviders` (une propriété de l'élément et une prop de chaque liaison) accepte des
`CloudFileProvider` de `ooxml-core/embedding`, par exemple `createOneDriveFileProvider`. Fichier >
Ouvrir affiche alors « Cet appareil » et un emplacement par fournisseur. Un fichier ouvert depuis un
emplacement est enregistré par son propre stockage et est en lecture seule quand le fournisseur
indique qu'il ne peut pas être modifié ; voir
[l'intégration OneDrive](/onedrive#open-from-onedrive-in-file-open).

## Stockage géré par l'hôte

`storage` (une propriété de l'élément et une prop de chaque liaison) accepte un `DocumentStorage` de
`ooxml-core/embedding`. Quand il est défini :

- Fichier > Enregistrer, Enregistrer sous et Ctrl+S transmettent les octets enregistrés à
  `storage.save(bytes, meta)` au lieu de les télécharger. `meta` vaut `{ fileName, mimeType, reason }`,
  avec `reason` égal à `'save'` ou `'saveAs'`. Une fois la promesse résolue, le document est marqué
  comme enregistré ; si elle se résout avec `{ fileName }`, la barre de titre prend ce nom. Un rejet
  est signalé par `document-error` et le document reste modifié (dirty).
- Fichier > Ouvrir appelle `storage.open()` quand le stockage le fournit et ouvre le
  `{ bytes, fileName }` renvoyé (`undefined` signifie que l'utilisateur a annulé). Sans `open`, le
  sélecteur de fichiers est utilisé.
- Un seul enregistrement s'exécute à la fois : Enregistrer pendant un enregistrement en cours est
  ignoré, pas mis en file d'attente. Une modification faite pendant l'enregistrement de l'hôte laisse
  le document modifié, puisque les octets enregistrés ne la contiennent pas.
- `file-command` est toujours émis en premier. Un hôte qui appelle `preventDefault()` gère la
  commande lui-même et `storage` n'est pas appelé. Enregistrer sous annonce `saveAs` puis `save`.
- Un hôte qui a retiré Enregistrer (`customization.capabilities.save: false`) ne reçoit ni
  l'événement `save` ni `storage.save` ; Enregistrer sous (la capacité `export`) ne fait alors que
  télécharger.

Exporter, Exporter en texte et Imprimer utilisent toujours le navigateur. Word n'a pas de
récupération automatique, donc les méthodes de récupération de `DocumentStorage` ne sont pas
utilisées.

```ts
editor.storage = {
	async save(bytes, { fileName, mimeType }) {
		await fetch(`/files/${encodeURIComponent(fileName)}`, {
			method: 'PUT',
			headers: { 'content-type': mimeType },
			body: bytes,
		});
	},
};
```

## Lecture à voix haute {#read-aloud}

Le hook des blocs de construction (`useViewerBuildingBlocks`) ajoute cinq champs de lecture à voix haute à
`toolbarProps`. Ce ne sont pas des propriétés ni des événements de l'élément. Ils utilisent l'implémentation
propre de l'éditeur (la synthèse vocale du navigateur ; rien n'est envoyé). Le
[guide des blocs de construction](/fr/building-blocks#read-aloud) décrit leur comportement.

| Champ             | Type                     | Comportement                                                                                                                                                                 |
| ----------------- | ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `canReadAloud`    | `boolean`                | Vrai quand un document est affiché et que le navigateur peut parler. Les documents en lecture seule comptent.                                                                |
| `readingAloud`    | `boolean`                | Vrai uniquement dans l'éditeur qui possède la voix. Si un autre éditeur lit, il vaut faux.                                                                                   |
| `readAloudIssue`  | `ReadAloudIssue \| null` | Le dernier problème : `'unsupported'` tant que le navigateur n'a pas de synthèse vocale, `'empty'` ou `'error'`. `null` quand il n'y en a pas.                               |
| `onReadAloud`     | `() => ReadAloudStart`   | Lit la sélection, ou du point d'insertion à la fin sans sélection. Relance si la lecture est déjà en cours. Renvoie `'started'` ou la raison pour laquelle rien n'a démarré. |
| `onStopReadAloud` | `() => boolean`          | Arrête la lecture de cet éditeur. Renvoie s'il lisait.                                                                                                                       |

Les codes sont de simples unions de chaînes. La bibliothèque ne fournit que les codes ; l'hôte formule le message.

```ts
type ReadAloudIssue = 'unsupported' | 'empty' | 'error';
type ReadAloudStart = 'started' | ReadAloudIssue;
```

`'error'` signifie que la voix a échoué. Un énoncé annulé ou interrompu n'est pas signalé comme une erreur.

Un bouton React de l'hôte qui bascule la lecture et formule le cas vide :

```tsx
import { DocumentCanvas, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel }) {
	const { canvasProps, toolbarProps } = useViewerBuildingBlocks({ documentModel });
	return (
		<>
			{toolbarProps && toolbarProps.readAloudIssue !== 'unsupported' && (
				<>
					<button
						disabled={!toolbarProps.canReadAloud}
						aria-pressed={toolbarProps.readingAloud}
						onClick={() =>
							toolbarProps.readingAloud
								? toolbarProps.onStopReadAloud()
								: toolbarProps.onReadAloud()
						}
					>
						{toolbarProps.readingAloud ? 'Stop reading' : 'Read aloud'}
					</button>
					{toolbarProps.readAloudIssue === 'empty' && <span role="status">Nothing to read.</span>}
				</>
			)}
			<DocumentCanvas {...canvasProps} />
		</>
	);
}
```

La lecture s'arrête aussi quand le document est modifié ou remplacé, et quand l'éditeur est démonté.
La `Toolbar` optionnelle dessine son propre bouton à bascule « Read aloud » à partir des mêmes champs.
