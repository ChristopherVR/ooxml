# Blocs de construction

Chaque paquet de framework (`@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer`, `xlsx-angular-viewer`,
`xlsx-svelte-viewer`, `xlsx-solid-viewer` et `xlsx-vanilla-viewer`) exporte un hook ou une fabrique
de blocs de construction (`useViewerBuildingBlocks`, `createViewerBuildingBlocks`), un `SpreadsheetCanvas`
et les contrôles facultatifs `Toolbar`, `FormulaBar` et `SheetTabs`. Tous utilisent la session de classeur,
la grille, les formules, la validation, la protection et l’historique existants ; il n’existe pas de
second magasin d’état. La forme suit le hook `useViewerBuildingBlocks` de la visionneuse PowerPoint.

```tsx
import {
	SpreadsheetCanvas,
	FormulaBar,
	SheetTabs,
	useViewerBuildingBlocks,
} from '@christophervr/xlsx-react-viewer';

function EmbeddedSheet({ workbook, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		workbook,
		canEdit,
		handle,
	});
	return (
		<>
			{toolbarProps && <FormulaBar {...toolbarProps} />}
			<SpreadsheetCanvas {...canvasProps} />
			{toolbarProps && <SheetTabs {...toolbarProps} />}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Le hook renvoie aussi `mode` (`'view'` ou `'edit'`) ainsi que les valeurs réactives `dirty`, `selection` (texte A1),
`collaboration` (les participants de la salle et l’état de la connexion) et `editError` (voir
[Échecs de validation](#validation-failures)). Affichez le canevas pendant le chargement ; `toolbarProps`
vaut null jusqu’au montage. Donnez une hauteur à son conteneur. Le canevas ne monte que la
grille. Fichier, paramètres, partage, barre de titre, ruban, barre d’état, barre de formule et onglets
de feuille sont absents, sauf si l’hôte fournit des contrôles distincts.

`toolbarProps` contient la sélection ou la cellule active, l’état des feuilles, le texte de la formule, la mise en forme et
les états `enabled`. Reliez `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange`,
`onFontSizeChange`, `onNumberFormatChange`, `onAlignmentChange`, `onUndo` et `onRedo`
à vos contrôles personnalisés. Les champs de mise en forme valent null pour les valeurs mixtes. L’inspection de la mise en forme est
limitée à 4096 cellules : au-delà, les sélections renvoient null et `formattingComplete=false`.

`enabled` comporte un indicateur par rappel : `bold`, `italic`, `underline`, `fontFamily`, `fontSize`,
`alignment`, `numberFormat`, `undo`, `redo`, ainsi que les actions de l’hôte `openShareDialog`, `settings`,
`formatCells`, `findReplace` et `insertFunction`. `canUndo` et `canRedo` reprennent les indicateurs annuler et rétablir.
`options` liste les choix qu’un contrôle personnalisé de police, de taille, d’alignement ou de format de nombre peut proposer
(`fontFamilies`, `fontSizes`, `alignments`, `numberFormats`), et `labels` contient les noms
localisés des contrôles.

`onSelect`, `onSheetChange`, `onFormulaChange` et `commitEdit` renvoient false lorsque la validation de la cellule
bloque l’action. Les rappels de mise en forme renvoient une `Promise<boolean>`. Ils valident une modification
de cellule active avant d’agir, et conservent sa sélection lorsque des contrôles personnalisés prennent le focus. Le texte de la formule est
validé par Entrée dans la FormulaBar facultative ; une saisie refusée définit aria-invalid.
La poignée existante conserve l’enregistrement, la sérialisation et le suivi des modifications. Un enregistrement est rejeté si la saisie
ne peut pas être validée. Les autorisations et la protection de la feuille restent prioritaires.

Pour chaque framework, `mountEditorSurface(host, options)` expose la même surface, la poignée existante,
`getToolbarProps()`, `subscribe(listener)`, `update` et `destroy`. Les hôtes gèrent eux-mêmes les
fichiers et les dialogues facultatifs ; aucune barre d’outils cachée de la bibliothèque n’est nécessaire.

`formula`, `formulaCaret`, `editing` et `editMode` lisent le pont d’édition en direct de la grille.
`onFormulaInput(text, caret?)` met à jour ce même brouillon ; `onFormulaChange(text)` le valide.
Un brouillon refusé peut être corrigé par `onFormulaInput` sans avoir d’abord validé la
valeur invalide. `onCancelEdit()` l’annule. Il n’existe pas de magasin d’état de formule distinct.

L’API de composition requiert le runtime `ooxml-ui` de la même version de l’espace de travail ; les paquets de framework en dépendent et sont publiés avec lui.
Les imports d’éditeur complet existants restent compatibles avec le runtime publié précédemment. Un runtime plus ancien
refuse explicitement une demande de composition au lieu de monter par erreur l’interface de l’application.

## Rappels de l’hôte pour les dialogues {#host-callbacks-for-dialogs}

L’application ouvre les dialogues de partage, de paramètres, Format de cellules, Rechercher et remplacer et Insérer
une fonction. Une surface ne les monte jamais. Passez plutôt un rappel, comme le font les blocs de construction
de PowerPoint (`onOpenSettings`, `onOpenShareDialog`), et ouvrez votre propre interface :

```ts
useViewerBuildingBlocks({
	workbook,
	onOpenShareDialog: ({ activeCell }) => openMyShareDialog(activeCell),
	onOpenFormatCells: ({ selection }) => openMyFormatPanel(selection),
});
```

Les rappels disponibles sont `onOpenShareDialog`, `onOpenSettings`, `onOpenFormatCells`,
`onOpenFindReplace` et `onOpenInsertFunction`. Chacun reçoit `{ sheet, sheetName, selection,
activeCell }`. `toolbarProps` expose les `onOpenShareDialog`, `onOpenSettings`,
`onOpenFormatCells`, `onOpenFindReplace` et `onOpenInsertFunction` correspondants (ils valident d’abord une
modification en cours et renvoient une `Promise<boolean>`), et `enabled.openShareDialog` et ses équivalents valent vrai uniquement lorsque
le rappel a été fourni (les actions de mise en forme et d’insertion exigent aussi l’autorisation de modification). `handle.share()`
appelle `onOpenShareDialog` ; sans ce rappel, il ne fait rien, et `enabled.openShareDialog` vaut false. Un
flux de travail sans rappel est inaccessible, il n’est jamais monté en silence. Les hôtes qui veulent les
dialogues intégrés affichent à la place le `SpreadsheetEditor` complet.

## Menus contextuels et opérations de flux {#context-menus-and-workflow-operations}

Les types, signatures et conditions de refus sont dans la [référence de l’API de l’élément](/fr/api#host-menus-and-workflows).

`onOpenFormatCells`, `onOpenFindReplace` et `onOpenInsertFunction` reçoivent, en plus du contexte
ci-dessus, les opérations pour terminer le flux. Les mêmes opérations sont dans
`toolbarProps.workflows` : un panneau toujours ouvert n’a donc pas besoin de rappel. Chaque
opération relit à l’exécution la lecture seule, la protection de la feuille et la modification de
cellule en cours, passe par la session d’édition (annuler, rétablir, état modifié et
`onWorkbookChange` se comportent comme dans l’interface fournie) et signale un refus par sa valeur
de retour.

- `workflows.formatCells` : `getState()` renvoie `{ font, alignment, border, protection }` (chacun
  `{ value, mixed }` : les propriétés communes et les noms qui diffèrent), `numFmt` et `fill`
  (null si mixte) et `complete` (false si la sélection est trop grande). `canApply()` et
  `apply(patch)` appliquent un correctif de style à la sélection en une seule étape d’annulation.
- `workflows.find` : `findAll(query)`, `findNext(query, { backwards })`, `select(match)`,
  `replace(query, text)`, `replaceAll(query, text)` et `canReplace()`. La recherche fonctionne dans
  un classeur en lecture seule ; le remplacement y est refusé et les cellules verrouillées par la
  protection sont ignorées. Comme le Remplacer d’Excel lui-même, `replace` et `replaceAll` n’appliquent pas la validation des données par défaut : le remplacement est écrit même si la règle de la cellule le refuserait à la saisie. Pour la vérifier, passez `{ validate: true, onRefused }` en troisième argument ou activez l’option d’éditeur `validateReplace` (prop `validateReplace` de chaque liaison et propriété de `<xlsx-editor>`), que la boîte de dialogue Rechercher et remplacer fournie respecte aussi. Une valeur refusée par la règle laisse sa cellule inchangée et figure dans `onRefused(cells)` (`{ sheet, row, col, ref, value, message, title? }`) ; la boîte de dialogue affiche un avis avec le nombre. `{ validate: false }` exclut un appel de l’option d’éditeur.
- `workflows.insertFunction` : `categories()`, `search({ text, category })` (avec `name`,
  `category`, `syntax` et une `description` localisée), `canInsert()` et `insert(name, args?)`, qui
  insère `NAME(` au curseur du brouillon de formule actif (en créant un si besoin) et renvoie
  `{ text, caret }` pour que votre champ suive. Un brouillon en texte simple est refusé.

Pour afficher votre propre menu contextuel, passez `onContextMenu(request)`. Il est appelé lors d’un
clic droit sur une cellule, un en-tête de ligne ou de colonne et (si une barre d’onglets est montée)
un onglet de feuille. Renvoyez `true` pour supprimer le menu natif et affichez `request.items`
(`id`, `label`, `command`, `enabled`, `checked`, `shortcut`, `separatorBefore`) à `request.x`,
`request.y` ; exécutez un choix avec `request.execute(id)`. Toute autre valeur affiche le menu
natif. Sans clic droit, `toolbarProps.contextMenu('cell' | 'row' | 'column' | 'tab', sheet?)`
renvoie le même modèle, ce qui permet à un hôte avec ses propres onglets d’obtenir le menu d’onglet.
Une surface enregistre peu de commandes : la plupart des éléments sont listés avec `enabled: false` ;
Format Cells est activé si `onOpenFormatCells` est installé et le classeur modifiable.
`enabled.contextMenu`, `enabled.find`, `enabled.replace`, `enabled.applyFormatCells` et
`enabled.insertFunctionCall` indiquent ce qui fonctionne maintenant. Pour conserver le menu natif et le modifier, passez plutôt `onContextMenuItems(request)` (il ne
s’exécute que si `onContextMenu` n’a pas renvoyé `true`) et renvoyez `{ hide, add }`. `hide` liste
les id d’éléments natifs de `request.items` ; `add` liste des éléments de l’hôte `{ id, label,
enabled, requiresEdit, checked, separatorBefore, position, onSelect(request) }`, placés au-dessus
(`position: 'top'`, par défaut) ou au-dessous (`'bottom'`) des éléments natifs. `enabled` peut être
une fonction. Les éléments natifs gardent leur état désactivé et sont revérifiés à l’exécution ; un
élément de l’hôte l’est aussi (son `enabled` et `requiresEdit` face à un classeur en lecture seule),
donc un menu resté ouvert ne peut pas exécuter un choix périmé. Il complète `hiddenActions` et
`contextMenu.extraCommands` de `customization`, qui s’appliquent à l’application fournie.

## Échecs de validation {#validation-failures}

Une modification de cellule en cours doit être validée avant la mise en forme, la navigation ou l’enregistrement. Lorsque la
validation des données la refuse, les rappels renvoient `false`, la modification reste ouverte et l’échec est
signalé de trois façons :

- `onValidationError({ message, title?, ref, style? })` dans les options (l’élément émet
  un événement `validation-error` avec le même détail) ;
- la chaîne réactive `editError` (aussi `toolbarProps.editError`), remise à null dès que le brouillon est
  modifié, corrigé ou annulé ;
- la `FormulaBar` facultative définit `aria-invalid` et affiche le message dans un élément `span` avec `role="alert"`.

### L’alerte de validation des données {#the-data-validation-alert}

Les blocs de construction ne dessinent jamais l’alerte Réessayer / Annuler de la grille. Les trois styles d’alerte d’Excel
relèvent de l’hôte, qui les affiche via `onOpenValidationAlert(request)` (également accepté dans `surfaceActions`
et par chaque adaptateur de framework). La demande contient `{ ref, sheet, sheetName, value, style,
title, message }`, où `value` est le texte saisi par l’utilisateur, ainsi que les méthodes qui correspondent aux
boutons d’Excel :

| `style`       | Boutons Excel       | Méthodes                                      |
| ------------- | ------------------- | --------------------------------------------- |
| `stop`        | Réessayer / Annuler | `retry()`, `cancel()`                         |
| `warning`     | Oui / Non / Annuler | `accept()` (Oui), `retry()` (Non), `cancel()` |
| `information` | OK / Annuler        | `accept()` (OK), `cancel()`                   |

`retry()` garde l’éditeur ouvert avec le brouillon, `cancel()` restaure l’ancienne valeur et `accept()`
valide la valeur malgré tout, exactement comme les boutons intégrés. Résolvez une demande une seule fois ; les appels suivants
sont ignorés. Tant qu’elle n’est pas résolue, l’éditeur de cellule ignore les validations suivantes ; résolvez-la donc
toujours (un rappel qui lève une exception compte comme `retry()`). `onValidationError` et `editError`
signalent toujours l’échec.

Sans le rappel, rien n’est monté : chaque style agit comme Réessayer, la modification reste ouverte et
l’échec est signalé par `onValidationError` / `editError`.
`toolbarProps.enabled.validationAlert` vaut true uniquement lorsque l’hôte a fourni le rappel. Le
`SpreadsheetEditor` complet et `<xlsx-editor>` conservent leur alerte intégrée.

Les contrôles personnalisés doivent aussi afficher `editError` et laisser l’utilisateur corriger le brouillon avec
`onFormulaInput` ou l’abandonner avec `onCancelEdit`.

## Barre d’outils facultative {#optional-toolbar}

Le `Toolbar` de la bibliothèque affiche chaque contrôle que les props prennent en charge, il peut donc remplacer une barre
construite à la main : annuler, rétablir, famille de police, taille de police, gras, italique, souligné, alignement
à gauche, au centre ou à droite, et format de nombre. Les boutons exposent `aria-pressed` (`mixed` pour une sélection mixte),
les listes de choix affichent une valeur vide pour une mise en forme mixte, et tous les contrôles respectent `enabled`. Les boutons ne
prennent pas le focus, si bien que la grille conserve sa sélection pendant la mise en forme. La barre d’outils n’affiche pas les
flux de dialogues ci-dessus ; reliez `onOpenShareDialog` et les autres rappels `onOpen*` à vos propres
contrôles. `toolbarModel(toolbarProps)` (exporté par `xlsx-vanilla-viewer` et le runtime Svelte) renvoie
la même liste de contrôles pour les rendus personnalisés.

## Collaboration {#collaboration}

La coédition en temps réel fonctionne sur une surface : passez `collaboration` (un identifiant de salle, une URL de serveur, un fournisseur ou
une session) dans les options, et le hook renvoie l’état réactif `collaboration` (`active`,
`roomId`, `status`, `synced`, `people`). `onUndo` et `onRedo` n’annulent que les modifications de cette fenêtre
pendant le partage. Quitter ou démonter le composant quitte la salle. La liste de présence de la barre de titre relève de
l’application ; un hôte affiche donc lui-même `collaboration.people`.

## Vanilla {#vanilla}

`xlsx-vanilla-viewer` exporte `createViewerBuildingBlocks(options)`, construit sur la même session que
les autres frameworks :

```ts
import { createViewerBuildingBlocks } from 'xlsx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ workbook, canEdit: true });
const stops = [
	blocks.mountToolbar(toolbarHost),
	blocks.mountFormulaBar(formulaHost),
	blocks.mountSheetTabs(tabsHost),
	blocks.mountCanvas(canvasHost),
];
blocks.subscribe((state) => render(state.dirty, state.editError, state.toolbarProps));
blocks.update({ workbook, canEdit: false });
blocks.getState().toolbarProps?.onBold();
blocks.destroy();
```

La surface est **en lecture seule tant que l’hôte ne passe pas `canEdit: true`**, comme les blocs de construction
de PowerPoint et de Visio (l’éditeur `SpreadsheetEditor` complet reste modifiable par défaut). L’autorisation est
appliquée avant le montage du canevas, si bien que `mode` démarre à `'view'` et qu’aucun clignotement en mode modification n’apparaît.
`readOnly` l’emporte toujours sur `canEdit: true`.

Chaque `mount*` renvoie son dispositif de destruction. `getState()` renvoie le même instantané que les hooks
(`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration`, `editError`), et
`handle` est la poignée impérative habituelle. Les contrôles mettent à jour leurs éléments sur place, si bien que
le focus et les brouillons survivent aux mises à jour.

## Composition native avec les frameworks {#native-framework-composition}

Le code de chaque framework est livré par son propre paquet. Tous les adaptateurs possèdent un moteur de
grille existant et exposent le même `EditorHandle` ; leurs instantanés réactifs natifs exposent
`toolbarProps`, `loading`, `error`, `mode`, `dirty`, `selection`, `collaboration` et `editError`. Les contrôles natifs facultatifs reçoivent une
prop/entrée `toolbarProps`. Les contrôles personnalisés appellent les actions de ce même instantané.

### Vue {#vue}

Appelez le composable pendant `setup`. Passez un getter pour suivre les changements d’autorisation, de document, d’octets et de
rappels. Les refs calculées sont déballées dans les gabarits ; le démontage libère le moteur.

```vue
<script setup lang="ts">
import {
	useViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-vue-viewer';
const props = defineProps<{ workbook: import('xlsx-vue-viewer').Workbook; canEdit: boolean }>();
const { canvasProps, toolbarProps, loading, error, handle } = useViewerBuildingBlocks(() => ({
	workbook: props.workbook,
	canEdit: props.canEdit,
}));
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<FormulaBar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<SpreadsheetCanvas v-bind="canvasProps" style="height: 600px" />
	<SheetTabs v-if="toolbarProps" :toolbar-props="toolbarProps" />
</template>
```

### Angular {#angular}

Fournissez `ViewerBuildingBlocks` par hôte, puis injectez cette instance dans l’hôte ou dans ses
contrôles personnalisés. `options` est un signal inscriptible ; les instantanés sont des signaux calculés. Le
`DestroyRef` de l’injecteur gère le nettoyage. Le service expose aussi les émetteurs d’événements `workbookChange`,
`workbookError`, `selectionChange` et `dirtyChange` pour les abonnements de l’hôte, en plus des
rappels des options. Utilisez `takeUntilDestroyed()` lorsque vous vous abonnez dans les hôtes.

```ts
@Component({
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [SpreadsheetCanvasComponent, ToolbarComponent, FormulaBarComponent, SheetTabsComponent],
	template: `
		@if (blocks.toolbarProps(); as p) {
			<spreadsheet-toolbar [toolbarProps]="p" />
			<spreadsheet-formula-bar [toolbarProps]="p" />
			<button (click)="p.onBold()" [disabled]="!p.canEdit">Custom bold</button>
			<spreadsheet-sheet-tabs [toolbarProps]="p" />
		}
		<spreadsheet-canvas [mount]="blocks.canvasProps.mount" style="display:block;height:600px" />
	`,
})
export class WorkbookHost {
	readonly workbook = input.required<Workbook>();
	readonly canEdit = input(true);
	readonly blocks = inject(ViewerBuildingBlocks);
	constructor() {
		effect(() => this.blocks.options.set({ workbook: this.workbook(), canEdit: this.canEdit() }));
	}
}
```

Importez les composants, le service et `Workbook` depuis `xlsx-angular-viewer`, et `Component`, `input`,
`effect` et `inject` depuis `@angular/core`. `SpreadsheetBuildingBlocks` est le nom propre à ce produit
du même service. Sa `handle` conserve le chargement, l’enregistrement et la sélection.

### Solid {#solid}

Créez le bloc dans un propriétaire Solid. Les getters d’entrée et les accesseurs de sortie font partie du graphe
réactif de Solid ; le nettoyage suit ce propriétaire. Les composants acceptent des getters réactifs comme props.

```tsx
import { Show } from 'solid-js';
import {
	createViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-solid-viewer';
function WorkbookHost(props: { workbook: Workbook; canEdit: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		workbook: props.workbook,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<Show when={blocks.toolbarProps()}>
				{(p) => (
					<>
						<Toolbar toolbarProps={p()} />
						<FormulaBar toolbarProps={p()} />
						<button onClick={() => p().onBold()}>Custom bold</button>
						<SheetTabs toolbarProps={p()} />
					</>
				)}
			</Show>
			<SpreadsheetCanvas {...blocks.canvasProps} />
		</>
	);
}
```

### Svelte 5 {#svelte-5}

Créez le magasin pendant l’initialisation du composant. `$blocks` s’abonne nativement ; utilisez un
`$effect` pour transmettre les options réactives. Le nettoyage du moteur suit la destruction du composant.
Les sources des composants et leurs déclarations sont publiées sous forme de sous-chemins explicites.

```svelte
<script lang="ts">
  import { createViewerBuildingBlocks, type Workbook } from 'xlsx-svelte-viewer/runtime';
  import SpreadsheetCanvas from 'xlsx-svelte-viewer/SpreadsheetCanvas.svelte';
  import Toolbar from 'xlsx-svelte-viewer/Toolbar.svelte';
  import FormulaBar from 'xlsx-svelte-viewer/FormulaBar.svelte';
  import SheetTabs from 'xlsx-svelte-viewer/SheetTabs.svelte';
  let { workbook, canEdit = false }: {workbook: Workbook; canEdit?: boolean} = $props();
  const blocks = createViewerBuildingBlocks();
  $effect(() => blocks.update({workbook, canEdit}));
</script>
{#if $blocks.toolbarProps}
  <Toolbar toolbarProps={$blocks.toolbarProps} />
  <FormulaBar toolbarProps={$blocks.toolbarProps} />
  <button onclick={() => $blocks.toolbarProps?.onBold()}>Custom bold</button>
  <SheetTabs toolbarProps={$blocks.toolbarProps} />
{/if}
<SpreadsheetCanvas {...blocks.canvasProps} />
```

`FormulaBar` partage le brouillon en direct de la grille. Entrée valide et définit `aria-invalid` sur une
saisie refusée. Réessayer permet de corriger ; Échap annule. Les mises en forme en conflit, la sélection,
la navigation entre feuilles et l’enregistrement restent bloqués tant que la validation refuse un brouillon. La lecture seule
bloque toute modification tout en préservant la sélection, la navigation entre feuilles et la sérialisation.

## Apparence partagée {#shared-appearance}

Placez le canevas et les contrôles natifs facultatifs sous un même ancêtre `data-office-theme="light|dark|auto"`. Les cinq
intégrations natives de framework installent les mêmes styles de contrôles et
lisent les jetons `--office-*` hérités. Personnalisez `--office-surface`, `--office-background`,
`--office-foreground`, `--office-border`, `--office-accent`, `--office-ring`, `--office-font`
et `--office-radius` à ce niveau. Les changements à l’exécution préservent les brouillons et le focus.

Les options existantes `theme` et `themeColors` du canevas continuent de s’appliquer localement ; elles ne
réécrivent pas les ancêtres ni les contrôles voisins de l’hôte. Les surcharges `--xve-*` héritées sur un
ancêtre commun atteignent toujours à la fois le canevas et les contrôles facultatifs. Voir [Thèmes](/fr/theming) pour le
périmètre et des exemples de compatibilité.
