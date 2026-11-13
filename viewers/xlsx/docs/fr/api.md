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

| Propriété            | Type                                     | Remarques                                                                                                                                                                                               |
| -------------------- | ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `themeColors`        | `Partial<Record<...>>`                   | Surcharges de jetons appliquées sous forme de propriétés personnalisées `--xve-*` ([thèmes](/theming)).                                                                                                 |
| `validateReplace`    | `boolean`                                | Remplacer et Remplacer tout vérifient la validation des données ; les cellules refusées restent inchangées (`false` par défaut, comme dans Excel).                                                      |
| `hiddenActions`      | `string[]`                               | Contrôles du ruban, éléments de la barre de titre et entrées de menu contextuel à masquer par identifiant stable.                                                                                       |
| `ribbonAddIns`       | `RibbonAddInTab[]`                       | Onglets que l’hôte ajoute après ceux intégrés ([liaisons](/bindings#ribbon-add-in-tabs)).                                                                                                               |
| `customization`      | `ExcelCustomization \| null`             | Capacités, pages de l’hôte dans l’onglet Fichier et commandes de menu de l’hôte (voir [Personnalisation](#customization)).                                                                              |
| `workbook`           | `Workbook \| null`                       | Obtient le classeur ouvert, ou en définit un pour le remplacer.                                                                                                                                         |
| `dirty`              | `boolean` (lecture seule)                | Modifications non enregistrées.                                                                                                                                                                         |
| `activeSheet`        | `number`                                 | Indice, à partir de zéro, de la feuille visible.                                                                                                                                                        |
| `collaboration`      | `XlsxCollaborationOptions \| null`       | Partage le classeur ouvert dans une salle tant qu’elle est définie ([collaboration](/fr/collaboration)).                                                                                                |
| `collaborationState` | `XlsxCollaborationState` (lecture seule) | Indique si le partage est actif, l’état de la connexion et qui est dans la salle.                                                                                                                       |
| `googleFonts`        | `false \| { baseUrl }`                   | Repli sur des polices web pour les polices que le lecteur n’a ni installées ni intégrées. Non défini, le choix du Centre de gestion de la confidentialité du lecteur s’applique (désactivé par défaut). |
| `surfaceOnly`        | booléen                                  | À définir avant le montage pour omettre l’interface de l’application ; associez-le à `surfaceActions`.                                                                                                  |
| `surfaceActions`     | `SurfaceHostActions`                     | Rappels de l’hôte qui remplacent les boîtes de dialogue lorsque `surfaceOnly` est défini. Un flux de travail sans rappel est signalé comme non activé dans `getToolbarProps().enabled`.                 |
| `loading`            | `boolean` (lecture seule)                | Indique si un classeur est en cours d’ouverture.                                                                                                                                                        |

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

| Méthode                          | Renvoie                   | Remarques                                                                                                                                                                                                                                           |
| -------------------------------- | ------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`         | `Promise<void>`           | `.xlsx`, `.xlsm`, `.xltx`, `.xls`, `.csv` ; le format est deviné à partir des octets. Émet `workbook-error` et rejette en cas d’échec.                                                                                                              |
| `newWorkbook()`                  | `void`                    | Un classeur vierge nommé `Book1.xlsx`.                                                                                                                                                                                                              |
| `save()`                         | `Promise<Blob>`           | `.xlsx`. N’efface pas `dirty`.                                                                                                                                                                                                                      |
| `saveBytes(format?)`             | `Promise<Uint8Array>`     | `'xlsx'` (par défaut) ou `'csv'` (la feuille active, valeurs mises en forme).                                                                                                                                                                       |
| `download(name?)`                | `Promise<void>`           | Enregistre, télécharge, puis `markClean()`.                                                                                                                                                                                                         |
| `markClean()`                    | `void`                    |                                                                                                                                                                                                                                                     |
| `select(ref)`                    | `void`                    | Une cellule ou plage A1 de la feuille active (`'C3'`, `'A1:D20'`).                                                                                                                                                                                  |
| `getSelection()`                 | `string`                  | La sélection courante sous forme de référence A1.                                                                                                                                                                                                   |
| `setActiveSheet(index)`          | `void`                    |                                                                                                                                                                                                                                                     |
| `undo()`, `redo()`               | `void`                    | Pendant le partage, seules les modifications de cette fenêtre sont annulées.                                                                                                                                                                        |
| `startCollaboration(options)`    | `Promise<void>`           | Rejoint une salle ; rejette si le partage est déjà actif ou si aucun classeur n’est ouvert.                                                                                                                                                         |
| `stopCollaboration()`            | `void`                    | Quitte la salle et efface `collaboration` ; le classeur reste tel quel.                                                                                                                                                                             |
| `reconnectCollaboration()`       | `void`                    | Relance la connexion en conservant les modifications faites hors ligne.                                                                                                                                                                             |
| `share()`                        | `void`                    | Ouvre la boîte de dialogue Partager.                                                                                                                                                                                                                |
| `focusGrid()`                    | `void`                    | Place le focus clavier sur la grille pour que la saisie modifie la cellule active.                                                                                                                                                                  |
| `openCloudFile(file)`            | `Promise<void>`           | Ouvre un fichier qu’un hôte a obtenu d’un `CloudFileProvider`, comme Fichier > Ouvrir > OneDrive : en lecture seule lorsque `file.canEdit` vaut false, Enregistrer et Enregistrer sous routés vers `file.storage`. Rejette si le chargement échoue. |
| `commitEdit()`                   | `boolean`                 | Valide une modification de cellule en cours avant qu’un hôte ne navigue ou n’enregistre. `false` signifie que la validation l’a bloquée.                                                                                                            |
| `getToolbarProps()`              | `SpreadsheetToolbarProps` | L’état headless de la barre d’outils et les indicateurs `enabled` pour les contrôles personnalisés ([blocs de construction](/building-blocks)).                                                                                                     |
| `subscribeEditorState(listener)` | `() => void`              | Appelle `listener` à chaque changement d’état de l’éditeur ; renvoie la fonction de désabonnement.                                                                                                                                                  |

## Événements

Tous les événements sont des `CustomEvent` qui remontent et traversent la frontière du shadow DOM (`composed`).

| Événement              | `detail`                                                                                              | Remarques                                                                                                                                         |
| ---------------------- | ----------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                                        | Après une modification, un chargement ou un nouveau classeur.                                                                                     |
| `workbook-error`       | `{ error, message }`                                                                                  | Un chargement ou un enregistrement a échoué.                                                                                                      |
| `workbook-warning`     | `{ warnings }`                                                                                        | Contenu lu par le chargeur mais qu’il n’a pas pu modéliser.                                                                                       |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }`                                                                  | L’état effectif de lecture seule. `locked` est défini par un rôle de collaboration `viewer` ; `hostReadOnly` est le commutateur propre de l’hôte. |
| `ribbon-action`        | `{ id }`                                                                                              | Une commande du ruban s’est exécutée.                                                                                                             |
| `ribbon-customize`     | `{ hiddenActions }`                                                                                   | L’utilisateur a modifié les commandes du ruban affichées (Fichier > Options > Personnaliser le ruban).                                            |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' \| 'share' }` | **Annulable** : `preventDefault()` pour la gérer vous-même.                                                                                       |
| `selection-change`     | `{ sheet, ref, active }`                                                                              | `ref` est la plage sélectionnée, `active` la cellule active.                                                                                      |
| `sheet-change`         | `{ index, name }`                                                                                     |                                                                                                                                                   |
| `dirty-change`         | `{ dirty }`                                                                                           |                                                                                                                                                   |
| `office-ribbon-add-in` | `{ tab, command }`                                                                                    | Une commande d’un onglet de l’hôte (`ribbonAddIns`) a été choisie.                                                                                |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                                          | Le partage a démarré ou s’est arrêté, la connexion a changé, ou quelqu’un est arrivé ou parti.                                                    |
| `validation-error`     | `{ message, title?, ref, style? }`                                                                    | Une modification de cellule en cours n’a pas passé la validation et reste ouverte jusqu’à correction ou annulation.                               |

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

## Stockage de l'hôte {#host-owned-storage}

`storage` (une propriété de l'élément et une prop de chaque binding) accepte un `DocumentStorage` de
`ooxml-core/embedding`. Lorsqu'elle est définie :

- Fichier > Enregistrer, Enregistrer sous et Ctrl+S transmettent le classeur enregistré à
  `storage.save(bytes, meta)` au lieu de le télécharger. `meta` vaut `{ fileName, mimeType, reason }`,
  avec `reason` égal à `'save'` ou `'saveAs'` ; les sources `.xls` et `.csv` sont enregistrées en
  `.xlsx`, comme pour un téléchargement. Une fois la promesse résolue, le classeur est marqué comme non
  modifié ; si elle est résolue avec `{ fileName }`, la barre de titre prend ce nom. Un rejet est
  signalé par `workbook-error` et le classeur reste modifié.
- Fichier > Ouvrir appelle `storage.open()` lorsque le stockage en dispose et ouvre le
  `{ bytes, fileName }` avec lequel elle est résolue (`undefined` signifie que l'utilisateur a annulé).
  Sans `open`, le sélecteur de fichiers est utilisé.
- Un seul enregistrement s'exécute à la fois : Enregistrer pendant un enregistrement en cours est
  ignoré, pas mis en file d'attente. Une modification faite pendant l'enregistrement de l'hôte laisse
  le classeur modifié, puisque les octets enregistrés ne la contiennent pas.
- `file-command` est toujours déclenché en premier ; un hôte qui appelle `preventDefault()` gère la
  commande lui-même et `storage` n'est pas appelé.

Exporter, Enregistrer au format CSV et Imprimer continuent de télécharger ou d'imprimer dans le
navigateur. Excel n'a pas de récupération automatique, donc les méthodes de récupération de
`DocumentStorage` ne sont pas utilisées.

```ts
editor.storage = {
	save: (bytes, { fileName }) => myDrive.put(fileName, bytes),
	open: () => myDrive.pick(), // resolves { bytes, fileName } or undefined
};
```

## Emplacements cloud dans Fichier > Ouvrir {#cloud-places-in-file-open}

`fileProviders` (une propriété de l'élément et une prop de chaque binding) accepte un tableau de
`CloudFileProvider` (`ooxml-core/embedding`), par exemple `createOneDriveFileProvider` de
`ooxml-core/embedding/onedrive`. Vide ou omis, Fichier > Ouvrir reste inchangé. Avec des fournisseurs,
Fichier > Ouvrir liste "Cet appareil" (Parcourir) et un emplacement par fournisseur, qui affiche le
navigateur partagé `<office-ui-cloud-files>` limité aux fichiers `.xlsx`. Ses textes suivent la
`locale` de l'éditeur.

- L'ouverture d'un fichier demande s'il faut abandonner les modifications non enregistrées, comme
  Ouvrir, puis charge les octets par le chemin de chargement normal, reprend le nom du fichier et ferme
  la vue Fichier.
- Le classeur s'enregistre ensuite via le `storage` propre au fichier, et non le `storage` de l'hôte :
  Enregistrer et Ctrl+S écrivent dans le même élément, Enregistrer sous écrit une copie, et le classeur
  n'est marqué comme non modifié qu'une fois la promesse du fournisseur résolue. Un enregistrement
  rejeté est signalé par `workbook-error` et le classeur reste modifié. Ouvrir, créer ou charger un
  autre classeur rend Enregistrer au `storage` de l'hôte.
- Un fichier que le fournisseur indique comme `canEdit: false` s'ouvre en lecture seule (`readOnly`
  vaut true, un événement `readonly-change` est déclenché). Le plus strict entre le `readOnly` de
  l'hôte et celui du fichier l'emporte, et la valeur propre de l'hôte n'est pas modifiée.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken,
		signIn,
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

## Menus et flux de travail gérés par l’hôte {#host-menus-and-workflows}

Une surface gérée par l’hôte (les [blocs de construction](/fr/building-blocks), ou `<xlsx-editor>` avec `surfaceOnly` et `surfaceActions`) n’a pas de menus contextuels ni de boîtes de dialogue propres pour Format de cellule, Rechercher et remplacer ou Insérer une fonction. L’hôte les dessine à l’aide des callbacks ci-dessous ou des opérations de `toolbarProps`. Le comportement est décrit dans [Menus contextuels et opérations de flux](/fr/building-blocks#context-menus-and-workflow-operations) ; cette section donne les types et les conditions de refus.

Les callbacks se définissent dans `surfaceActions` ou se passent comme options à `useViewerBuildingBlocks`, `createViewerBuildingBlocks` et `mountEditorSurface`.

### Callbacks de menu contextuel {#context-menu-callbacks}

```ts
onContextMenu?: (request: ContextMenuRequest) => boolean | void;
onContextMenuItems?: (request: ContextMenuRequest) => ContextMenuCustomization | void;
```

`onContextMenu` s’exécute lors d’un clic droit ou de Maj+F10 sur une cellule, un en-tête de ligne, un en-tête de colonne ou un onglet de feuille. Renvoyez `true` pour supprimer le menu natif et dessiner le vôtre ; toute autre valeur affiche le menu natif. `onContextMenuItems` ne s’exécute que lorsque `onContextMenu` n’a pas renvoyé `true`. Il filtre et étend le menu natif, qui reste dessiné par Excel.

```ts
type ContextMenuKind = 'cell' | 'row' | 'column' | 'tab';

interface SurfaceActionContext {
	sheet: number;
	sheetName: string;
	selection: string; // texte A1
	activeCell: string; // texte A1
}

interface ContextMenuItem {
	id: string;
	label: string; // localisé
	command?: string;
	enabled: boolean;
	checked?: boolean;
	shortcut?: string;
	separatorBefore?: boolean;
}

interface ContextMenuModel extends SurfaceActionContext {
	menu: ContextMenuKind;
	menuSheet: number; // l’onglet cliqué pour 'tab', sinon la feuille active
	readOnly: boolean;
	items: ContextMenuItem[];
	execute(id: string): Promise<boolean>;
}

interface ContextMenuRequest extends ContextMenuModel {
	x: number; // coordonnées client du clic droit
	y: number;
}

interface HostContextMenuItem {
	id: string; // unique parmi les éléments de l’hôte d’un menu ; un identifiant vide ou répété est ignoré
	label: string;
	enabled?: boolean | (() => boolean); // une fonction est réévaluée à l’exécution de l’élément
	requiresEdit?: boolean; // désactivé, et refusé à l’exécution, dans un classeur en lecture seule
	checked?: boolean;
	separatorBefore?: boolean;
	position?: 'top' | 'bottom'; // 'top' par défaut
	onSelect(request: ContextMenuRequest): void | Promise<void>;
}

interface ContextMenuCustomization {
	hide?: readonly string[]; // identifiants des éléments natifs de request.items
	add?: readonly HostContextMenuItem[];
}
```

`request.execute(id)` résout `true` lorsque l’élément a été exécuté. Elle résout `false` lorsque la modification de cellule en cours ne passe pas la validation (la modification reste ouverte), lorsque `id` ne figure pas dans `items`, ou lorsque l’élément ne peut pas être exécuté à cet instant (lecture seule, protection de feuille, ou commande indisponible sur la surface). Un menu d’onglet active d’abord sa feuille. Les éléments de l’hôte de `onContextMenuItems` exécutent `onSelect(request)` après la fermeture du menu, sauf s’ils sont désactivés entre-temps.

### toolbarProps.contextMenu {#toolbarprops-contextmenu}

```ts
contextMenu(menu: ContextMenuKind, sheet?: number): ContextMenuModel;
```

Renvoie le même modèle sans clic droit, par exemple pour un hôte qui dessine ses propres onglets de feuille ou sa barre d’outils. `sheet` vaut par défaut la feuille active et désigne l’onglet concerné par un menu `'tab'`. Les commandes indisponibles sur la surface figurent avec `enabled: false`. Format de cellule est activé lorsque `onOpenFormatCells` est installé et que le classeur est modifiable. Les entrées qui exécutent une action de la barre d’onglets intégrée, comme Renommer, figurent désactivées.

### Callbacks de flux {#workflow-callbacks}

```ts
onOpenFormatCells?: (request: FormatCellsRequest) => void;
onOpenFindReplace?: (request: FindReplaceRequest) => void;
onOpenInsertFunction?: (request: InsertFunctionRequest) => void;

interface FormatCellsRequest extends SurfaceActionContext, FormatCellsOperations {}
interface FindReplaceRequest extends SurfaceActionContext, FindOperations {}
interface InsertFunctionRequest extends SurfaceActionContext, InsertFunctionOperations {}
```

Chaque requête contient le contexte ci-dessus plus les opérations de ce flux, décrites plus bas. Le callback ne s’exécute que si le flux est disponible : Format de cellule et Insérer une fonction exigent les droits de modification, Rechercher et remplacer non. La modification de cellule en attente est d’abord validée. Les fonctions `onOpen*` de `toolbarProps` renvoient un `Promise<boolean>`, valant `false` lorsque le callback manque, que les droits de modification manquent ou que la modification en attente ne passe pas la validation.

### toolbarProps.workflows {#toolbarprops-workflows}

`toolbarProps.workflows` expose les mêmes opérations sans callback, donc un panneau toujours ouvert n’a pas besoin de callback `onOpen*`. Chaque opération relit l’état de lecture seule, la protection de feuille et la modification de cellule en attente au moment de son exécution.

#### Format de cellule {#format-cells-operations}

```ts
interface FormatCellsOperations {
	getState(): FormatCellsState;
	canApply(): boolean;
	apply(patch: StylePatch): boolean;
}

interface FormatAspect<T> {
	value: T; // les propriétés sur lesquelles toutes les cellules examinées s’accordent
	mixed: string[]; // les noms des propriétés qui diffèrent
}

interface FormatCellsState {
	font: FormatAspect<Font>;
	alignment: FormatAspect<Alignment>;
	border: FormatAspect<Border>;
	protection: FormatAspect<Protection>;
	numFmt: string | null; // null lorsque la sélection mélange des formats
	fill: Fill | null; // null lorsque la sélection mélange des remplissages
	complete: boolean; // false lorsque la sélection compte plus de 4096 cellules
}
```

- `getState()` renvoie la mise en forme de la sélection. Lorsque `complete` vaut `false`, les listes de propriétés mixtes sont partielles.
- `canApply()` vaut true lorsque le classeur est modifiable et que la protection de feuille autorise la mise en forme.
- `apply(patch)` applique un `StylePatch` à la sélection en une seule étape d’annulation. Il renvoie `false`, sans rien modifier, lorsque `canApply()` vaut `false` ou que la modification de cellule en attente ne passe pas la validation.

#### Rechercher et remplacer {#find-and-replace-operations}

```ts
interface FindQuery {
	text: string;
	sheet?: number;
	range?: CellRange; // uniquement avec sheet
	matchCase?: boolean;
	wholeCell?: boolean;
	lookIn?: 'formulas' | 'values' | 'comments';
	wildcards?: boolean; // true par défaut
	order?: 'rows' | 'columns';
}

interface FindMatch {
	sheet: number;
	row: number;
	col: number;
	text: string;
}

interface FindOperations {
	canReplace(): boolean;
	findAll(query: FindQuery): FindMatch[];
	findNext(query: FindQuery, options?: { backwards?: boolean }): FindMatch | undefined;
	select(match: FindMatch): boolean;
	replace(query: FindQuery, replacement: string, options?: ReplaceOptions): boolean;
	replaceAll(query: FindQuery, replacement: string, options?: ReplaceOptions): number;
}
```

- `canReplace()` vaut true lorsque le classeur est modifiable. La recherche n’exige jamais de droits de modification.
- `findAll(query)` renvoie toutes les correspondances et fonctionne dans un classeur en lecture seule. Il renvoie `[]` pour un `text` vide, ou lorsque l’éditeur n’est pas connecté ou charge encore.
- `findNext(query, { backwards })` sélectionne la correspondance suivante (ou précédente) après la cellule active, en reprenant au début si besoin. Il renvoie `undefined` lorsqu’il n’y a pas de correspondance ou que la modification de cellule en attente ne passe pas la validation.
- `select(match)` sélectionne une correspondance issue de `findAll`. Il renvoie `false` lorsque l’éditeur n’est pas prêt ou que la feuille n’existe pas.
- `replace(query, replacement)` remplace la correspondance de la cellule active, s’il y en a une, puis sélectionne la suivante. Il renvoie `true` uniquement si une cellule a été remplacée. Il renvoie `false` lorsque le classeur est en lecture seule, que la modification en attente ne passe pas la validation, ou que la cellule active est une cellule verrouillée d’une feuille protégée. Comme le Remplacer d’Excel, il n’applique pas la validation des données.
- `replaceAll(query, replacement)` renvoie le nombre de cellules modifiées. Il renvoie `0` lorsque le classeur est en lecture seule ou que la modification en attente ne passe pas la validation. Les cellules verrouillées par la protection de feuille sont ignorées. La validation des données n’est pas appliquée.

Les deux acceptent un troisième argument facultatif `{ validate?, onRefused? }` (`ReplaceOptions`). Avec `validate: true`, ou l'option d'éditeur `validateReplace`, une valeur refusée par la validation des données de la cellule reste inchangée et est signalée une fois par appel à `onRefused` sous la forme `{ sheet, row, col, ref, value, message, title? }` ; `validate: false` soustrait un appel à l'option de l'éditeur.

#### Insérer une fonction {#insert-function-operations}

```ts
interface FunctionDoc {
	name: string;
	category: string;
	syntax: string; // par exemple SUM(number1, [number2], ...)
	description: string; // localisé
}

interface InsertFunctionOperations {
	categories(): string[];
	search(options?: { text?: string; category?: 'all' | 'recent' | string }): FunctionDoc[];
	canInsert(): boolean;
	insert(name: string, args?: string): { text: string; caret: number } | undefined;
}
```

- `categories()` renvoie les noms de catégories distincts, triés.
- `search(options)` compare `text` d’abord au nom, puis aux mots de la description. `category` vaut `'all'` (par défaut), `'recent'` pour les fonctions utilisées récemment, ou un nom de catégorie.
- `canInsert()` vaut true lorsque le classeur est modifiable.
- `insert(name, args?)` insère `NAME(`, ou `NAME(args)` lorsque `args` est fourni, au point d’insertion du brouillon de formule ouvert. Si aucun brouillon n’est ouvert, il démarre `=NAME(`. L’insertion ne se fait jamais avant le `=` initial. Il renvoie le `text` et le `caret` du brouillon après l’insertion, pour que la saisie puisse continuer à partir de là. Il renvoie `undefined`, sans rien modifier, lorsque le classeur est en lecture seule, que `name` n’est pas une fonction documentée, ou que le brouillon ouvert est du texte brut (il ne commence pas par `=`).

### Indicateurs toolbarProps.enabled {#toolbarprops-enabled}

| Indicateur           | Vrai lorsque                                                                                                                                    |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `contextMenu`        | `onContextMenu` est installé et l’éditeur est connecté.                                                                                         |
| `find`               | L’éditeur est connecté, chargé et dispose d’une session. Aucun droit de modification requis.                                                    |
| `replace`            | Le classeur est modifiable (`workflows.find.canReplace()`).                                                                                     |
| `applyFormatCells`   | La mise en forme peut être appliquée maintenant : classeur modifiable et protection de feuille autorisant (`workflows.formatCells.canApply()`). |
| `insertFunctionCall` | Le classeur est modifiable (`workflows.insertFunction.canInsert()`).                                                                            |
| `formatCells`        | `onOpenFormatCells` est installé et le classeur est modifiable.                                                                                 |
| `findReplace`        | `onOpenFindReplace` est installé.                                                                                                               |
| `insertFunction`     | `onOpenInsertFunction` est installé et le classeur est modifiable.                                                                              |

Les indicateurs sont calculés pour chaque instantané de `toolbarProps`. Chaque opération revérifie son état au moment de son exécution.

### Exemple {#host-menus-and-workflows-example}

```tsx
import { useViewerBuildingBlocks } from '@christophervr/xlsx-react-viewer';

const { toolbarProps } = useViewerBuildingBlocks({
	workbook,
	canEdit: true,
	onContextMenu(request) {
		if (request.menu !== 'cell') return; // les en-têtes et les onglets gardent le menu natif
		showMyMenu(request.x, request.y, request.items, (id) => request.execute(id));
		return true; // supprime le menu natif
	},
	onOpenFindReplace: ({ findAll, select }) => showFindPanel({ findAll, select }),
});

// Plus tard, depuis un panneau à vous (toolbarProps est null tant que la zone n’est pas montée) :
toolbarProps?.workflows.formatCells.apply({ font: { bold: true } });
```

## Parties shadow

L’élément expose ses régions sous forme de parties shadow CSS, pour les styler avec `xlsx-editor::part(grid)` et pour les tests : `ribbon`, `name-box`, `formula-bar`, `grid`, `sheet-tabs`, `status-bar`, `title-bar`, `backstage` et `page-layout` (vue Mise en page).
