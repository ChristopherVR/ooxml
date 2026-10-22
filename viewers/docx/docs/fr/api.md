# API de l'élément et événements

Chaque adaptateur de framework encapsule le même élément personnalisé `<docx-editor>`. Les adaptateurs transmettent ces propriétés et réémettent ces événements selon les conventions de leur framework (voir les [adaptateurs de frameworks](/bindings)).

## Attributs et propriétés

| Propriété         | Attribut          | Description                                                                                                                |
| ----------------- | ----------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `documentModel`   |                   | Le `DocumentModel` affiché. L'assigner est un remplacement externe, pas un événement d'édition.                            |
| `readOnly`        | `read-only`       | Bloque les modifications locales. Les lots de collaboration distants s'appliquent toujours.                                |
| `locale`          | `locale`          | Langue de l'interface : `en`, `fr`, `de`, `es` ou `zh-CN`. Le contenu du document n'est jamais traduit.                    |
| `fileName`        | `file-name`       | Nom du document ouvert (`Document1.docx` par défaut).                                                                      |
| `reviewAuthor`    | `review-author`   | Auteur enregistré sur les modifications suivies et les commentaires (`Author` par défaut).                                 |
| `theme`           | `theme`           | `light`, `dark` ou `auto`. Voir les [thèmes](/theming).                                                                    |
| `themeColors`     |                   | Remplacements partiels de jetons, appliqués comme propriétés personnalisées `--dve-*`.                                     |
| `showThumbnails`  | `show-thumbnails` | Rail de vignettes de pages à gauche (mise en page Impression uniquement).                                                  |
| `showToolbar`     | `show-toolbar`    | `false` masque le ruban ; la barre de titre et la barre d'état restent affichées.                                          |
| `hiddenActions`   |                   | Commandes du ruban à masquer par identifiant stable ; voir le tableau dans les [adaptateurs](/bindings).                   |
| `ribbonAddIns`    |                   | Onglets ajoutés par l'hôte après les onglets intégrés ; voir les [adaptateurs](/bindings#ribbon-add-in-tabs).              |
| `customization`   |                   | Capacités, pages de l’hôte dans l’onglet Fichier et commandes de menu de l’hôte (voir [Personnalisation](#customization)). |
| `dirty` (lecture) |                   | Vrai après une modification, faux après un enregistrement, un chargement ou `markClean()`.                                 |

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

| Méthode                                               | Description                                                                                         |
| ----------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `load(bytes)`                                         | Charge des octets DOCX ou DOC hérité. Rejette et émet `document-error` en cas d'échec.              |
| `save()`                                              | Se résout en un `Blob` du document sérialisé.                                                       |
| `saveBytes()`                                         | Se résout en le `Uint8Array` brut.                                                                  |
| `download(fileName?)`                                 | Enregistre et lance un téléchargement dans le navigateur.                                           |
| `markClean()`                                         | Efface l'indicateur dirty une fois le document persisté par l'hôte.                                 |
| `startCollaboration(config)`                          | Rejoint une session de collaboration ; voir la [collaboration](/fr/collaboration).                  |
| `startYjsCollaboration(session, options)`             | Rejoint une salle Yjs synchronisée avec son paquet source correspondant.                            |
| `reconnectCollaboration()`, `resyncCollaboration()`   | Rétablit la connexion Yjs ou demande un échange d'état au fournisseur.                              |
| `getPendingCollaboration()`                           | Le lot d'étapes en attente, ou `null`.                                                              |
| `receiveCollaboration(batch)`                         | Applique un lot accepté par l'autorité.                                                             |
| `stopCollaboration(discardPending?)`                  | Quitte la session ; les modifications en attente doivent être acquittées ou explicitement ignorées. |
| `publishPresence`, `receivePresence`, `leavePresence` | Canal de présence transitoire ; voir la [collaboration](/fr/collaboration#présence).                |

## Événements

Tous les événements remontent (bubble) et sont composés, de sorte qu'ils traversent la frontière du shadow DOM.

| Événement              | Détail                | Signification                                                                                    |
| ---------------------- | --------------------- | ------------------------------------------------------------------------------------------------ |
| `document-change`      | `DocumentModel`       | Le document a changé par l'édition ou par des actions de révision.                               |
| `document-error`       | `Error`               | Le chargement, l'enregistrement, l'impression ou une modification a échoué.                      |
| `document-warning`     | `string`              | Avertissement non fatal.                                                                         |
| `readonly-change`      | `boolean`             | L'utilisateur a basculé Édition/Lecture dans l'interface.                                        |
| `file-command`         | `{ command }`         | Annulable. `command` vaut `new`, `open`, `save`, `export`, `print` ou `share`.                   |
| `dirty-change`         | `boolean`             | L'état « modifications non enregistrées » a basculé.                                             |
| `page-change`          | `{ page, pageCount }` | La page ou le nombre de pages de la mise en page Impression a changé (pagination approximative). |
| `ribbon-action`        | `RibbonAction`        | Une commande du ruban a été activée.                                                             |
| `ribbon-customize`     | `readonly string[]`   | L'utilisateur a modifié les commandes affichées du ruban ; persistez les identifiants si besoin. |
| `office-ribbon-add-in` | `{ tab, command }`    | Une commande d'un onglet de l'hôte (`ribbonAddIns`) a été choisie.                               |
| `presence-send`        | `PresenceMessage`     | La présence locale a changé ; transmettez-la aux autres clients.                                 |
| `collaboration-send`   | `StepBatch`           | Des étapes locales sont en attente ; transmettez-les à l'autorité.                               |

Le paquet exporte la table typée des événements (`DocxEditorEventMap`) et la liste d'exécution `DOCX_EDITOR_EVENTS`.
