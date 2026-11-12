# Liaisons de frameworks

Toutes les liaisons montent `<xlsx-editor>` par la même fonction `mountEditor` de `packages/bindings/src/index.ts`, afin que chaque framework bénéficie exactement de la même sémantique.

## Props {#props}

| Prop             | Type                          | Défaut       | Remarques                                                                                                                                                                  |
| ---------------- | ----------------------------- | ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook`       | `Workbook`                    | aucun        | Affiché lorsqu’un **nouvel objet** est passé. Le classeur émis par l’éditeur n’est jamais réaffecté au parent (pas de boucle).                                             |
| `bytes`          | `Uint8Array \| ArrayBuffer`   | aucun        | Octets du fichier ; rechargé à chaque nouveau tableau transmis.                                                                                                            |
| `src`            | `string`                      | aucun        | URL récupérée et chargée lorsque la chaîne change ; le nom du fichier correspond par défaut au dernier segment de l’URL.                                                   |
| `fileName`       | `string`                      | `Book1.xlsx` | Transmis uniquement lorsque le parent le modifie, afin que Fichier > Ouvrir dans l’éditeur puisse renommer le classeur.                                                    |
| `readOnly`       | `boolean`                     | `false`      |                                                                                                                                                                            |
| `locale`         | `string`                      | `en`         | `en`, `fr`, `de`, `es`, `zh-CN` ou une balise qui correspond à l’une d’elles (`de-DE`).                                                                                    |
| `theme`          | `'light' \| 'dark' \| 'auto'` | `auto`       | `auto` suit le système d’exploitation.                                                                                                                                     |
| `authorName`     | `string`                      | `Author`     | Enregistré sur les nouveaux commentaires.                                                                                                                                  |
| `showToolbar`    | `boolean`                     | `true`       | Le ruban.                                                                                                                                                                  |
| `showFormulaBar` | `boolean`                     | `true`       | La zone de nom et la barre de formule.                                                                                                                                     |
| `hiddenActions`  | `string[]`                    | `[]`         | Contrôles du ruban, éléments de la barre de titre et entrées du menu contextuel à masquer, par identifiant stable.                                                         |
| `ribbonAddIns`   | `RibbonAddInTab[]`            | aucun        | Onglets que l’hôte ajoute après ceux intégrés ; voir [Onglets complémentaires du ruban](#ribbon-add-in-tabs).                                                              |
| `themeColors`    | `Partial<Record<...>>`        | `{}`         | Surcharges de jetons, voir [thèmes](/fr/theming).                                                                                                                          |
| `collaboration`  | `XlsxCollaborationOptions`    | aucun        | Partage le classeur ouvert dans une salle tant qu’il est défini ; un nouvel objet rejoint la salle, `null` la quitte. Voir [collaboration](/fr/collaboration).             |
| `storage`        | `DocumentStorage`             | aucun        | Stockage géré par l’hôte : Enregistrer et Enregistrer sous passent par `storage.save()`, et Fichier > Ouvrir par `storage.open()` lorsqu’il existe.                        |
| `fileProviders`  | `CloudFileProvider[]`         | aucun        | Emplacements cloud proposés par Fichier > Ouvrir à côté de « Cet appareil » ; un classeur ouvert depuis l’un d’eux s’enregistre dans ce fichier. Gardez le tableau stable. |
| `customization`  | `ExcelCustomization`          | aucun        | `capabilities`, `backstage.extraPages` et `contextMenu.extraCommands` ; gardez l’objet stable (`null` restaure les valeurs par défaut).                                    |

## Rappels et événements {#callbacks-and-events}

| Rappel (React, Solid, vanilla) | Événement Vue          | Sortie Angular        | Prop Svelte             | Charge utile                                |
| ------------------------------ | ---------------------- | --------------------- | ----------------------- | ------------------------------------------- |
| `onWorkbookChange`             | `workbook-change`      | `workbookChange`      | `onworkbookchange`      | `Workbook`                                  |
| `onWorkbookError`              | `workbook-error`       | `workbookError`       | `onworkbookerror`       | `Error`                                     |
| `onSelectionChange`            | `selection-change`     | `selectionChange`     | `onselectionchange`     | `{ sheet, ref, active }`                    |
| `onDirtyChange`                | `dirty-change`         | `dirtyChange`         | `ondirtychange`         | `boolean`                                   |
| `onReadOnlyChange`             | `readonly-change`      | `readOnlyChange`      | `onreadonlychange`      | `boolean`                                   |
| `onFileCommand`                | `file-command`         | `fileCommand`         | `onfilecommand`         | `CustomEvent<FileCommandDetail>`, annulable |
| `onWorkbookWarning`            | `workbook-warning`     | `workbookWarning`     | `onworkbookwarning`     | `string[]`                                  |
| `onRibbonAction`               | `ribbon-action`        | `ribbonAction`        | `onribbonaction`        | `{ id }`                                    |
| `onRibbonCustomize`            | `ribbon-customize`     | `ribbonCustomize`     | `onribboncustomize`     | `string[]` (actions masquées)               |
| `onSheetChange`                | `sheet-change`         | `sheetChange`         | `onsheetchange`         | `{ index, name }`                           |
| `onRibbonAddIn`                | `office-ribbon-add-in` | `ribbonAddIn`         | `onribbonaddin`         | `{ tab, command }`                          |
| `onCollaborationChange`        | `collaboration-change` | `collaborationChange` | `oncollaborationchange` | `XlsxCollaborationState`                    |
| `onReady`                      | `ready`                | `ready`               | `onready`               | l’élément `<xlsx-editor>`                   |

Les échecs de chargement (fichiers corrompus, chiffrés ou non pris en charge) parviennent à `onWorkbookError` ; un échec de récupération de `src` aussi. Une récupération `src` encore en cours est abandonnée lorsque de nouveaux `bytes` ou un nouveau `workbook` arrivent, ou lorsque `src` est effacé ou modifié.

`onFileCommand` reçoit l’événement `file-command` annulable lui-même, et non son détail : appelez `event.preventDefault()` pour empêcher une commande du menu Fichier (Nouveau, Ouvrir, Enregistrer, Enregistrer sous, Exporter, Exporter en CSV, Imprimer, Partager), et lisez `event.detail.command` pour savoir laquelle. Vue, Angular et Svelte émettent ou rappellent avec ce même objet événement. Chaque autre rappel reçoit directement sa charge utile. Voir l’[API de l’élément](/fr/api).

Chaque prop n’est transmise que lorsque sa valeur change (`hiddenActions` et `themeColors` sont comparées par valeur), afin qu’un nouveau rendu du parent n’annule jamais ce que l’utilisateur a changé dans l’éditeur : Modification / Affichage, thème ou langue de Fichier > Options, ou Personnaliser le ruban. Un parent contrôlé garde son état synchronisé grâce à `onReadOnlyChange` et `onRibbonCustomize` (dans Angular, `[(readOnly)]`).

## Onglets complémentaires du ruban {#ribbon-add-in-tabs}

Un hôte peut ajouter ses propres onglets du ruban après ceux d’Excel, comme le fait un complément Office.
`ribbonAddIns` prend des données simples et constitue une prop de chaque liaison ; une commande choisie
exécute son rappel `run` et est signalée par `onRibbonAddIn` (Vue
`office-ribbon-add-in`, Angular `ribbonAddIn`, Svelte `onribbonaddin`) avec
`{ tab, command }`.

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [
					{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() },
					{ id: 'options', label: 'Options', size: 'small' },
					{ id: 'send', label: 'Send', items: [{ id: 'send-mail', label: 'By mail' }] },
				],
			},
		],
	},
];

<SpreadsheetEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

Les commandes `size: 'small'` remplissent des colonnes de trois et `items` crée une liste déroulante.
Transmettre à nouveau un descripteur égal, comme le fait un framework à chaque rendu, conserve les
panneaux et remplace seulement les rappels ; omettre la prop supprime les
onglets. Un onglet ne peut pas reprendre l’identifiant d’un onglet intégré, et ses commandes
n’atteignent jamais la gestion de commandes propre à l’éditeur. Le même descripteur (`RibbonAddInTab`, issu de
`ooxml-core/ribbon` et réexporté par chaque paquet de framework) fonctionne dans les éditeurs
Word, Excel, PowerPoint et Visio. Aucun onglet complémentaire n’est fourni avec l’éditeur.

## Poignée {#handle}

Chaque adaptateur expose la même poignée : React par `ref`, Vue par la référence de gabarit, Angular par l’instance du composant, Solid par `editorRef`, Svelte par les exports du composant et vanilla comme valeur de retour de `mountEditor`.

| Membre                          | Effet                                                                                                                                                   |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `element`                       | Le `<xlsx-editor>` monté.                                                                                                                               |
| `load(bytes, fileName?)`        | Ouvre des octets `.xlsx`, `.xlsm`, `.xls` ou `.csv` ; rejette la promesse en cas d’échec.                                                               |
| `newWorkbook()`                 | Remplace le classeur par un classeur vierge.                                                                                                            |
| `save()`                        | Se résout en un Blob `.xlsx`. Ne réinitialise pas `dirty`.                                                                                              |
| `saveBytes(format?)`            | Octets `'xlsx'` (par défaut) ou `'csv'` (feuille active).                                                                                               |
| `download(fileName?)`           | Enregistre, télécharge dans le navigateur et marque le classeur comme enregistré.                                                                       |
| `markClean()`                   | Réinitialise `dirty` après que vous avez vous-même conservé un Blob enregistré.                                                                         |
| `select(ref)`, `getSelection()` | Sélection sous forme de référence A1 (`'B2:D8'`).                                                                                                       |
| `setActiveSheet(index)`         | Change de feuille.                                                                                                                                      |
| `share()`                       | Ouvre la boîte de dialogue Partager. Le démontage de l’éditeur quitte toute salle partagée.                                                             |
| `openCloudFile(file)`           | Ouvre un fichier d’un `CloudFileProvider` comme le fait Fichier > Ouvrir > OneDrive : son nom, son état lecture seule et son stockage pour Enregistrer. |
| `dirty`                         | Modifications non enregistrées. `isDirty()` de Svelte est déprécié ; lisez `dirty`.                                                                     |

Appeler une méthode de la poignée avant que l’éditeur soit monté lève l’erreur `Editor is not mounted`.

## Rendu côté serveur {#server-rendering}

Chaque paquet s’importe sans DOM. L’enregistrement (`defineXlsxEditor`) et le montage sont des opérations côté client.
