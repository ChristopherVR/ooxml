# Blocs de construction

`docx-react-viewer` exporte `useViewerBuildingBlocks`, `DocumentCanvas` et les composants facultatifs `Toolbar`, `PageNavigation` et `StatusBar`. Ils utilisent le moteur existant de WordEditor (édition, mise en page, historique et sérialisation) et suivent la forme du `useViewerBuildingBlocks` de PowerPoint : le hook renvoie des propriétés, et l'hôte affiche ce qu'il veut.

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, navigationProps, statusProps, loading, error } =
		useViewerBuildingBlocks({
			documentModel,
			canEdit,
			handle,
			onDirtyChange: (dirty) => console.log(dirty),
			// Host-owned dialogs: omit one and its action reports as not enabled.
			onOpenHyperlinkDialog: (request) => openMyLinkDialog(request),
			onOpenFindReplace: (request) => openMyFindBar(request),
			onOpenImagePicker: (request) => openMyFilePicker(request),
		});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{statusProps && <span>{statusProps.wordText}</span>}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Affichez le canevas même pendant le chargement. Les propriétés valent `null` tant qu'il n'est pas monté. Donnez une hauteur au conteneur du canevas.

## Ce que renvoie le hook {#what-the-hook-returns}

| Champ                      | Contenu                                                                                 |
| -------------------------- | --------------------------------------------------------------------------------------- |
| `canvasProps`              | À répartir sur `DocumentCanvas`. Monte et libère l'éditeur.                             |
| `toolbarProps`             | État de la sélection et de la mise en forme, plus les callbacks d'édition (ci-dessous). |
| `navigationProps`          | Page, nombre de pages, mode d'affichage et callbacks de navigation entre les pages.     |
| `statusProps`              | Nombre de mots et zoom, avec leurs callbacks.                                           |
| `loading`, `error`, `mode` | État du chargement, dernière erreur de chargement ou de document, `'view'` ou `'edit'`. |

### toolbarProps {#toolbarprops}

État : `selection` (`{ from, to }`), `bold`, `italic`, `underline` (`null` lorsque la sélection mélange les valeurs), `fontFamily`, `fontSize` (`null` si mélangé), `alignment` (`'left' | 'center' | 'right' | 'justify'` ou `null`), `list` (`'bullet' | 'number' | 'other' | null`), `canEdit`, `canUndo`, `canRedo` et `labels` traduits.

Callbacks, qui renvoient chacun `false` lorsque l'action n'est pas autorisée (lecture seule ou chargement en cours) : `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange(family)`, `onFontSizeChange(points)`, `onAlignmentChange(alignment)`, `onListChange('bullet' | 'number' | 'remove')`, `onUndo`, `onRedo`. L'état `list` et `onListChange` utilisent les mêmes valeurs `'bullet' | 'number'`.

### navigationProps {#navigationprops}

`page`, `pageCount`, `pageText` (« Page 2 sur 5 »), `pagesAvailable`, `canGoPrevious`, `canGoNext`, `viewMode` (`'draft' | 'print'`) et les callbacks `onGoToPage(n)`, `onPreviousPage`, `onNextPage` et `onViewModeChange(mode)`. Les pages n'existent qu'en mise en page Impression : en mode Brouillon, `pagesAvailable` vaut faux et les callbacks de page renvoient `false`. Basculez avec `onViewModeChange('print')`.

### statusProps {#statusprops}

`words`, `wordText` (« 42 mots »), `zoom` (en pourcentage), `minZoom`, `maxZoom` et les callbacks `onZoomChange(percent)`, `onZoomIn` et `onZoomOut`. Le défilement met aussi à jour `navigationProps`.

## Boîtes de dialogue de l'hôte {#host-dialogs}

Une surface n'affiche aucune boîte de dialogue propre. Ces workflows appartiennent à l'hôte et portent les noms des callbacks `onOpen...` de PowerPoint. Passez-les en options ; `toolbarProps` transporte alors un déclencheur `open...` correspondant (`openHyperlinkDialog`, `openFindReplace`, `openImagePicker`), un indicateur `can...` et `hostWorkflows` (les callbacks fournis).

| Option                  | Objet de requête                                                                 | Activé pendant que                                              |
| ----------------------- | -------------------------------------------------------------------------------- | --------------------------------------------------------------- |
| `onOpenHyperlinkDialog` | `{ href, text, apply(href, text?), remove() }` pour le lien à la sélection       | le document est modifiable                                      |
| `onOpenFindReplace`     | `{ search, findNext, findPrevious, replaceOne, replaceAll, status, canReplace }` | le document est chargé (la lecture seule conserve la recherche) |
| `onOpenImagePicker`     | `{ insert(file) }` : l'hôte choisit un fichier et le remet                       | le document est modifiable                                      |

Si un callback est omis, `canInsertLink`, `canFindReplace` ou `canInsertPicture` vaut faux et le déclencheur renvoie `false`. Les requêtes revérifient l'autorisation à leur utilisation, de sorte qu'une boîte de dialogue restée ouverte lors d'un passage en lecture seule ne puisse pas modifier le document.

## Lecture à voix haute {#read-aloud}

`toolbarProps` fournit un contrôle de lecture à voix haute qui utilise l'implémentation de l'éditeur
(la synthèse vocale du navigateur ; rien n'est envoyé). Il lit la sélection ou, sans sélection, du
point d'insertion à la fin du document, et fonctionne dans un document en lecture seule : il ne le
modifie jamais et ne le marque pas comme modifié. L'appel ne déplace ni la sélection ni le focus ;
un bouton de votre interface peut donc prendre le focus avant.

| Propriété ou callback | Signification                                                                                          |
| --------------------- | ------------------------------------------------------------------------------------------------------ |
| `canReadAloud`        | Un document est affiché et le navigateur peut parler                                                   |
| `readingAloud`        | Cet éditeur lit en ce moment                                                                           |
| `readAloudIssue`      | `'unsupported'` (pas de synthèse vocale), `'empty'` (rien à lire), `'error'` ou `null`                 |
| `onReadAloud()`       | Lance la lecture (la relance si besoin) ; renvoie `'started'`, `'unsupported'`, `'empty'` ou `'error'` |
| `onStopReadAloud()`   | Arrête la lecture de cet éditeur ; renvoie s'il lisait                                                 |

Le navigateur n'a qu'une voix. Démarrer dans un éditeur la prend à un autre, dont `readingAloud`
devient `false`. Un éditeur cesse toujours de lire quand son document est remplacé, qu'un fichier se charge ou qu'il est démonté. Formulez les messages vous-même ; la bibliothèque ne fournit que les
codes. La `Toolbar` optionnelle dessine un bouton à bascule « Read aloud » et l'omet là où la
synthèse vocale n'existe pas. La référence des champs, avec les types et un exemple, se trouve dans
l'[API de l'élément](/fr/api#read-aloud).

Une modification ordinaire (locale ou d'un collaborateur) n'arrête pas la lecture par défaut : la documentation
publique de Microsoft sur Word (support.microsoft.com, « Listen to your Word documents » et « Use Immersive Reader in
Word ») décrit le démarrage et la sortie de la lecture à voix haute, et indique qu'on peut cliquer dans le document pour
lire ou modifier, mais ne dit pas que modifier l'arrête. La lecture continue donc avec le texte du départ. Passez
`readAloud: { stopOnEdit: true }` comme option des blocs de construction (ou renseignez la propriété `readAloud` de
l'éditeur) pour l'arrêter à chaque modification. Chaque liaison la transmet : la prop `readAloud` dans React, Vue,
Solid, Angular et Svelte, et l'option des blocs vanilla.

## Composants de bibliothèque {#library-components}

`Toolbar` affiche annuler et rétablir, la police et sa taille, gras, italique, souligné, quatre alignements, les listes à puces et numérotées, et un bouton pour chaque workflow de l'hôte que vous avez fourni. `PageNavigation` affiche la page précédente, « Page n sur m », la page suivante et le sélecteur Brouillon ou Mise en page Impression. `StatusBar` affiche le nombre de mots et une liste de zoom. Tous sont facultatifs et pilotés par leurs propriétés ; chaque contrôle a un nœud stable, de sorte que le focus clavier survit aux mises à jour de mise en forme. Ils utilisent les styles Office partagés. Les contrôles personnalisés n'en ont jamais besoin.

```tsx
{
	toolbarProps && <Toolbar {...toolbarProps} />;
}
{
	navigationProps && <PageNavigation {...navigationProps} />;
}
{
	statusProps && <StatusBar {...statusProps} />;
}
```

## Surfaces en lecture seule {#read-only-surfaces}

La surface est **en lecture seule tant que l'hôte ne passe pas `canEdit: true`**, comme les blocs de construction PowerPoint et Visio (le `WordEditor` complet reste modifiable par défaut). L'autorisation est appliquée avant le montage du canevas, de sorte que `mode` démarre à `'view'` et qu'il n'y a aucun clignotement en mode édition. `readOnly` l'emporte toujours sur `canEdit: true`. En lecture seule, chaque callback d'édition renvoie `false` et le document ne peut pas changer. Le lecteur peut tout de même donner le focus à la surface, sélectionner du texte à la souris ou au clavier (flèches, Maj+flèches, Début, Fin, Ctrl+A), copier, et lire la sélection depuis `toolbarProps.selection`. Les raccourcis clavier propres à l'application (rechercher, enregistrer, imprimer) ne sont pas installés sur une surface ; reliez-les à vos propres contrôles.

**Migrer un hôte.** Les blocs de construction étaient autrefois modifiables par défaut. Un hôte qui s'appuyait sur ce comportement doit désormais passer `canEdit: true` (dans l'objet d'options ou, pour un adaptateur natif, dans les options réactives) ; sans cela, la surface s'ouvre en lecture seule et chaque callback d'édition renvoie `false`.

## Cycle de vie {#lifecycle}

Utilisez `bytes` pour charger un fichier, ou les méthodes du handle `load`, `save`, `download`, `markClean` et `dirty`. L'enregistrement ne marque pas le document comme propre tant que l'hôte n'a pas confirmé la persistance. Le canevas ne monte ni barre de titre, ni ruban, ni backstage, ni barre d'état, ni boîte de dialogue d'application.

Chaque paquet de framework exporte aussi `mountEditorSurface(host, options)` pour les adaptateurs à cycle de vie natif. Son binding expose `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` et `subscribe(listener)`, ainsi que le handle de l'éditeur, `update` et `destroy`. Les options de surface excluent celles qui ne concernent que le ruban (`showToolbar`, `showThumbnails`, `hiddenActions`, `ribbonAddIns`, `onRibbonAddIn`), de sorte qu'un hôte TypeScript ne puisse pas en passer une qui serait ignorée.

L'API de composition exige le runtime `ooxml-ui` correspondant, publié avec ces ajouts. Les imports existants de l'éditeur complet restent compatibles avec le runtime publié précédemment. Un runtime plus ancien refuse explicitement une demande de composition au lieu de monter une interface d'application par accident.

## Composition native par framework {#native-framework-composition}

Les adaptateurs de framework vivent dans leur propre paquet de framework. Ils partagent l'assistant de durée de vie du moteur et les descripteurs de contrôles (`bindings/src/controls.ts`) que dessinent les composants facultatifs. Chaque adaptateur expose `toolbarProps`, `navigationProps`, `statusProps`, `loading`, `error`, `mode` et `canvasProps`, ainsi que le `handle` de l'éditeur (React prend une référence `handle` comme option au lieu d'en renvoyer une), et les composants `Toolbar`, `PageNavigation` et `StatusBar` (Angular : `ToolbarComponent`, `PageNavigationComponent`, `StatusBarComponent` ; Svelte : `docx-svelte-viewer/Toolbar`, `/PageNavigation`, `/StatusBar`). Leur état utilise les primitives réactives habituelles du framework, et leurs composants de canevas montent la surface de document Word existante.

### Vue {#vue}

```vue
<script setup lang="ts">
import { computed } from 'vue';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-vue-viewer';
const props = defineProps<{ documentModel: import('docx-core').DocumentModel; canEdit: boolean }>();
const { canvasProps, toolbarProps, error, handle } = useViewerBuildingBlocks(
	computed(() => ({ documentModel: props.documentModel, canEdit: props.canEdit })),
);
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<DocumentCanvas v-bind="canvasProps" class="document-surface" />
	<p v-if="error" role="alert">{{ error.message }}</p>
</template>
```

Le composable observe une ref, une computed ref ou un getter, et libère le moteur avec son scope. `toolbarProps`, `loading`, `error` et `mode` sont des computed refs en lecture seule. `handle` expose l'API impérative existante de l'éditeur.

### Angular {#angular}

Fournissez un service `ViewerBuildingBlocks` par composant hôte. Les signaux lus dans le template participent à la détection de changements d'Angular, y compris dans les applications sans zone. `DestroyRef` libère le moteur lorsque son provider est détruit.

```ts
import { Component, signal } from '@angular/core';
import { createDocument } from 'docx-core';
import {
	createViewerBuildingBlocks,
	ViewerBuildingBlocks,
	DocumentCanvasComponent,
	ToolbarComponent,
} from 'docx-angular-viewer';

@Component({
	selector: 'embedded-word',
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent],
	providers: [ViewerBuildingBlocks],
	template: `
		@if (blocks.toolbarProps(); as toolbar) {
			<word-toolbar [toolbarProps]="toolbar" />
		}
		<word-document-canvas [canvasProps]="blocks.canvasProps" />
	`,
})
export class EmbeddedWord {
	readonly canEdit = signal(true);
	readonly model = createDocument();
	readonly blocks = createViewerBuildingBlocks(() => ({
		documentModel: this.model,
		canEdit: this.canEdit(),
	}));
}
```

Le service expose les signaux en lecture seule `toolbarProps`, `loading`, `error` et `mode`, un `handle` stable, et des émetteurs d'événements typés `documentChange`, `documentError`, `dirtyChange` et `pageChange`. Abonnez-vous à ces émetteurs ou fournissez les callbacks d'options existants. Pour les hôtes impératifs, injectez le service et appelez `update(options)` au lieu de l'assistant.

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
  import { writable } from 'svelte/store';
  import { createDocument } from 'docx-core';
  import { createViewerBuildingBlocks } from 'docx-svelte-viewer/runtime';
  import DocumentCanvas from 'docx-svelte-viewer/DocumentCanvas';
  import Toolbar from 'docx-svelte-viewer/Toolbar';
  const options = writable({ documentModel: createDocument(), canEdit: true });
  const blocks = createViewerBuildingBlocks(options);
  const { toolbarProps, error } = blocks;
</script>
{#if $toolbarProps}<Toolbar toolbarProps={$toolbarProps} />{/if}
<DocumentCanvas {...blocks.canvasProps} />
{#if $error}<p role="alert">{$error.message}</p>{/if}
```

Appelez la fabrique pendant l'initialisation du composant. Elle accepte un store lisible d'options ou un objet d'options initial ; `blocks.update(next)` prend aussi en charge les mises à jour `$effect` pilotées par les runes. Les instantanés sont des stores lisibles. `documentCanvas` est exporté comme action pour un balisage hôte personnalisé. La destruction du composant se désabonne des options et détruit le moteur.

### Solid {#solid}

```tsx
import { createSignal, Show } from 'solid-js';
import { createDocument } from 'docx-core';
import { createViewerBuildingBlocks, DocumentCanvas, Toolbar } from 'docx-solid-viewer';
function EmbeddedWord() {
	const [canEdit, setCanEdit] = createSignal(true);
	const model = createDocument();
	const blocks = createViewerBuildingBlocks(() => ({ documentModel: model, canEdit: canEdit() }));
	return (
		<>
			<Show when={blocks.toolbarProps()}>{(toolbar) => <Toolbar toolbarProps={toolbar()} />}</Show>
			<DocumentCanvas {...blocks.canvasProps} />
		</>
	);
}
```

La fabrique suit son accesseur d'options dans un effet et renvoie des accesseurs de signaux pour `toolbarProps`, `loading`, `error` et `mode`. Son propriétaire contrôle le nettoyage via `onCleanup`. Les contrôles facultatifs de la barre d'outils conservent leurs nœuds DOM pendant que l'état de mise en forme change, de sorte que le focus clavier est préservé. Tous les adaptateurs acceptent `bytes` de façon réactive, effacent les chargements échoués lorsqu'un nouveau modèle ou un nouveau chargement arrive, et gardent leur handle stable lors du remplacement du canevas.

### Vanilla {#vanilla}

`docx-vanilla-viewer` exporte la session partagée sous la forme `createViewerBuildingBlocks(options)` (`canvasProps`, `handle`, `state`, `subscribe`, `update`, `destroy`) et des assistants de montage qui renvoient une fonction d'arrêt : `mountDocumentCanvas(host, blocks)`, `mountToolbar(host, blocks)`, `mountPageNavigation(host, blocks)` et `mountStatusBar(host, blocks)`.

```ts
import {
	createViewerBuildingBlocks,
	mountDocumentCanvas,
	mountToolbar,
	mountStatusBar,
} from 'docx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ documentModel, canEdit: true });
const stops = [
	mountDocumentCanvas(document.querySelector('#page')!, blocks),
	mountToolbar(document.querySelector('#tools')!, blocks),
	mountStatusBar(document.querySelector('#status')!, blocks),
];
blocks.subscribe((state) => console.log(state.toolbarProps?.bold)); // your own controls
// later
stops.forEach((stop) => stop());
blocks.destroy();
```

## Style partagé {#shared-styling}

Les composants natifs facultatifs de barre d'outils utilisent les styles partagés de `ooxml-ui`, y compris les espacements, les fonds, les bordures, les états enfoncé et désactivé, et les anneaux de focus clavier. Placez la barre d'outils et le canevas sous un même ancêtre `data-office-theme`, et personnalisez les variables `--office-*` à cet endroit. Les changements s'appliquent immédiatement à tous les frameworks. Les surcharges `--dve-*` existantes sur un ancêtre restent prises en charge. Les contrôles créés par l'hôte gardent leur propre style, sauf si l'hôte adopte l'assistant partagé `prepareOfficeUiBlock(element, 'toolbar', 'docx')` de `ooxml-ui/theme`.
