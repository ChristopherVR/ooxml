# API de l’élément

`<xlsx-editor>` (classe `XlsxEditorElement`) est l’unique éditeur que chaque paquet monte. Enregistrez-le avec `defineXlsxEditor()` (exporté par chaque paquet ; l’enregistrement est idempotent et ne fait rien sans `customElements`), ou laissez `mountEditor` et les composants de framework s’en charger.

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const editor = document.createElement('xlsx-editor');
editor.fileName = 'Budget.xlsx';
document.body.append(editor);
await editor.load(bytes, 'Budget.xlsx');
```

## Attributs et propriétés

| Attribut           | Propriété        | Type                                                           | Défaut       |
| ------------------ | ---------------- | -------------------------------------------------------------- | ------------ |
| `locale`           | `locale`         | `en \| fr \| de \| es \| zh-CN` (toute balise se ramène à une) | `en`         |
| `read-only`        | `readOnly`       | booléen                                                        | `false`      |
| `file-name`        | `fileName`       | chaîne                                                         | `Book1.xlsx` |
| `author-name`      | `authorName`     | chaîne                                                         | `Author`     |
| `theme`            | `theme`          | `light \| dark \| auto`                                        | `auto`       |
| `show-toolbar`     | `showToolbar`    | booléen (`"false"` masque)                                     | `true`       |
| `show-formula-bar` | `showFormulaBar` | booléen (`"false"` masque)                                     | `true`       |

Propriétés uniquement :

| Propriété            | Type                                     | Remarques                                                                                                                  |
| -------------------- | ---------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `themeColors`        | `Partial<Record<...>>`                   | Surcharges de jetons appliquées sous forme de propriétés personnalisées `--xve-*` ([thèmes](/theming)).                    |
| `hiddenActions`      | `string[]`                               | Contrôles du ruban, éléments de la barre de titre et entrées de menu contextuel à masquer par identifiant stable.          |
| `ribbonAddIns`       | `RibbonAddInTab[]`                       | Onglets que l’hôte ajoute après ceux intégrés ([liaisons](/bindings#ribbon-add-in-tabs)).                                  |
| `customization`      | `ExcelCustomization \| null`             | Capacités, pages de l’hôte dans l’onglet Fichier et commandes de menu de l’hôte (voir [Personnalisation](#customization)). |
| `workbook`           | `Workbook \| null`                       | Obtient le classeur ouvert, ou en définit un pour le remplacer.                                                            |
| `dirty`              | `boolean` (lecture seule)                | Modifications non enregistrées.                                                                                            |
| `activeSheet`        | `number`                                 | Indice, à partir de zéro, de la feuille visible.                                                                           |
| `collaboration`      | `XlsxCollaborationOptions \| null`       | Partage le classeur ouvert dans une salle tant qu’elle est définie ([collaboration](/fr/collaboration)).                   |
| `collaborationState` | `XlsxCollaborationState` (lecture seule) | Indique si le partage est actif, l’état de la connexion et qui est dans la salle.                                          |

## Personnalisation {#customization}

`customization` accepte un objet `ExcelCustomization`, le contrat partagé de `ooxml-core/embedding` que les autres éditeurs utilisent aussi. Chaque liaison le transmet comme prop ou entrée `customization`. Il est appliqué quand l’identité de l’objet change ; gardez-le donc stable. `null` rétablit les valeurs par défaut.

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send range', onSelect: ({ ref }) => send(ref) }],
	},
};
```

- `capabilities` : `false` supprime tous les chemins vers la capacité (`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`) : ruban, pages de l’onglet Fichier, menus et raccourcis. `hiddenActions` est inchangé ; les deux sont réunis.
- `backstage.extraPages` : pages de l’hôte `{ id, label, group?, order?, render }` dans l’onglet Fichier, après les pages intégrées de leur groupe. `render` est une fonction de montage, une balise d’élément personnalisé ou `{ element, properties }` et reçoit un `ExcelRegionContext`.
- `contextMenu.extraCommands` : commandes de l’hôte `{ id, label, icon?, group?, when?, disabled?, onSelect }` au-dessus (`top`, par défaut) ou au-dessous (`bottom`) des entrées intégrées. `onSelect` s’exécute après la fermeture du menu avec un `ExcelMenuContext`. Les libellés sont affichés tels quels (localisez-les vous-même).

`edit: false` est un verrouillage en lecture seule à part entière, à côté de l’interrupteur de lecture seule de l’hôte et d’un rôle de collaboration `viewer` : aucun ne lève un autre. Les capacités sont une politique d’interface, pas une protection du document. Le remplacement de régions (boîte d’options, volets) n’est pas encore proposé pour Excel. Tous les produits réunis : [guide d’intégration](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Méthodes

| Méthode                       | Renvoie               | Remarques                                                                                                                              |
| ----------------------------- | --------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`      | `Promise<void>`       | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv` ; le format est deviné à partir des octets. Émet `workbook-error` et rejette en cas d’échec. |
| `newWorkbook()`               | `void`                | Un classeur vierge nommé `Book1.xlsx`.                                                                                                 |
| `save()`                      | `Promise<Blob>`       | `.xlsx`. N’efface pas `dirty`.                                                                                                         |
| `saveBytes(format?)`          | `Promise<Uint8Array>` | `'xlsx'` (par défaut) ou `'csv'` (la feuille active, valeurs mises en forme).                                                          |
| `download(name?)`             | `Promise<void>`       | Enregistre, télécharge, puis `markClean()`.                                                                                            |
| `markClean()`                 | `void`                |                                                                                                                                        |
| `select(ref)`                 | `void`                | Une cellule ou plage A1 de la feuille active (`'C3'`, `'A1:D20'`).                                                                     |
| `getSelection()`              | `string`              | La sélection courante sous forme de référence A1.                                                                                      |
| `setActiveSheet(index)`       | `void`                |                                                                                                                                        |
| `undo()`, `redo()`            | `void`                | Pendant le partage, seules les modifications de cette fenêtre sont annulées.                                                           |
| `startCollaboration(options)` | `Promise<void>`       | Rejoint une salle ; rejette si le partage est déjà actif ou si aucun classeur n’est ouvert.                                            |
| `stopCollaboration()`         | `void`                | Quitte la salle et efface `collaboration` ; le classeur reste tel quel.                                                                |
| `reconnectCollaboration()`    | `void`                | Relance la connexion en conservant les modifications faites hors ligne.                                                                |
| `share()`                     | `void`                | Ouvre Fichier > Partager.                                                                                                              |
| `focusGrid()`                 | `void`                | Place le focus clavier sur la grille pour que la saisie modifie la cellule active.                                                     |

## Événements

Tous les événements sont des `CustomEvent` qui remontent et traversent la frontière du shadow DOM (`composed`).

| Événement              | `detail`                                                                                   | Remarques                                                                                                           |
| ---------------------- | ------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                             | Après une modification, un chargement ou un nouveau classeur.                                                       |
| `workbook-error`       | `{ error, message }`                                                                       | Un chargement ou un enregistrement a échoué.                                                                        |
| `workbook-warning`     | `{ warnings }`                                                                             | Contenu lu par le chargeur mais qu’il n’a pas pu modéliser.                                                         |
| `readonly-change`      | `{ readOnly }`                                                                             |                                                                                                                     |
| `ribbon-action`        | `{ id }`                                                                                   | Une commande du ruban s’est exécutée.                                                                               |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' }` | **Annulable** : `preventDefault()` pour la gérer vous-même.                                                         |
| `selection-change`     | `{ sheet, ref, active }`                                                                   | `ref` est la plage sélectionnée, `active` la cellule active.                                                        |
| `sheet-change`         | `{ index, name }`                                                                          |                                                                                                                     |
| `dirty-change`         | `{ dirty }`                                                                                |                                                                                                                     |
| `office-ribbon-add-in` | `{ tab, command }`                                                                         | Une commande d’un onglet de l’hôte (`ribbonAddIns`) a été choisie.                                                  |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                               | Le partage a démarré ou s’est arrêté, la connexion a changé, ou quelqu’un est arrivé ou parti.                      |
| `validation-error`     | `{ message, title?, ref, style? }`                                                         | Une modification de cellule en cours n’a pas passé la validation et reste ouverte jusqu’à correction ou annulation. |

```ts
editor.addEventListener('file-command', (event) => {
	if (event.detail.command === 'save') {
		event.preventDefault();
		void editor
			.save()
			.then(uploadToMyStorage)
			.then(() => editor.markClean());
	}
});
```

## Parties shadow

L’élément expose ses régions sous forme de parties shadow CSS, pour les styler avec `xlsx-editor::part(grid)` et pour les tests : `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs` et `status-bar`.
