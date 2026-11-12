# Blocs de construction

`visio-react-viewer` exporte `useViewerBuildingBlocks`, `DiagramCanvas` et, en option, `Toolbar`, `PageNavigation`, `Inspector` et `ShapesPane`. Ils réutilisent ViewerController, le moteur de rendu SVG existant, l'inspecteur de sélection et la fenêtre des formes. Chaque autre framework dispose des mêmes éléments (voir plus bas), y compris le JavaScript vanilla.

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error, markClean } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
		onOpenSettings: () => openMySettingsDialog(),
	});
	return (
		<>
			<PageNavigation {...toolbarProps} />
			<DiagramCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

`content` accepte la même source (octets ou Blob) que le visualiseur. `handle` reçoit le ViewerController existant ; utilisez `load`, `applyEdits`, `exportVsdx` et son `state.edit.dirty` réactif pour la persistance côté hôte. `onDocumentChange`, `onDirtyChange` et `onDocumentError` sont facultatifs. Donnez au conteneur du canevas une hauteur et laissez-le monté pendant le chargement.

Le `toolbarProps` à plat constitue tout le contrat dont une interface personnalisée a besoin ; chaque champ provient du contrôleur existant, de sorte que rien n'est stocké deux fois :

| Champ                                                                                                                               | Signification                                                                                                                                                                                                                                                                                                                                                |
| ----------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `mode`, `canEdit`                                                                                                                   | `mode` vaut `'edit'` seulement si l'hôte a accordé `canEdit` **et** si le dessin ouvert est un `.vsdx` adossé à sa source (un `.vsd` ancien ou un dessin réduit au modèle indique `'view'`). `canEdit` n'est vrai que si l'hôte autorise la modification, que le dessin est adossé à sa source et qu'aucun chargement ni aucune modification n'est en cours. |
| `pageIndex`, `pages`, `zoom`, `selection`                                                                                           | Page courante, liste des pages, facteur de zoom et formes sélectionnées.                                                                                                                                                                                                                                                                                     |
| `canUndo`, `canRedo`, `canDuplicate`, `canGroup`, `canUngroup`, `canDelete`, `canCut`, `canCopy`, `canPaste`                        | Indicateurs d'activation. Les actions de modification sont fausses en lecture seule ; `canCopy` exige un dessin adossé à sa source et une sélection sur la page courante, mais pas l'autorisation de modification.                                                                                                                                           |
| `canZoomIn`, `canZoomOut`                                                                                                           | Limites de zoom (de 10 % à 800 %).                                                                                                                                                                                                                                                                                                                           |
| `search`                                                                                                                            | `{ query, resultCount, activeIndex, truncated }` issu de la recherche du contrôleur. `activeIndex` vaut -1 tant que l'hôte ne navigue pas.                                                                                                                                                                                                                   |
| `tool`, `tools`, `canDraw`                                                                                                          | Outil actif du canevas (`'pointer'` ou l'un de `rectangle`, `ellipse`, `line`, `freeform`, `arc`, `pencil`, `text`, `connector`), la liste avec les libellés localisés et les indicateurs d'activation, et si le dessin est autorisé.                                                                                                                        |
| `labels`                                                                                                                            | Libellés typés et localisés (`DiagramToolbarLabels`), une entrée pour chaque commande, en `en`, `de`, `es`, `fr` et `zh-CN`.                                                                                                                                                                                                                                 |
| `onPageChange`, `onZoomChange`, `onZoomIn`, `onZoomOut`, `onFit`                                                                    | Actions d'affichage, toujours disponibles.                                                                                                                                                                                                                                                                                                                   |
| `onUndo`, `onRedo`, `onDuplicate`, `onGroup`, `onUngroup`, `onDelete`, `onCut`, `onCopy`, `onPaste`, `onToolChange`, `onTextChange` | Actions de modification. Elles conservent les contrôles d'autorisation, de chargement et de format source du contrôleur. Les actions du presse-papiers utilisent le canevas monté et le presse-papiers du navigateur, elles nécessitent donc un geste de l'utilisateur.                                                                                      |
| `onSearchChange`, `onNextSearchResult`, `onPreviousSearchResult`                                                                    | Recherche, toujours disponible.                                                                                                                                                                                                                                                                                                                              |

`Inspector` consomme `inspectorProps`, et `ShapesPane` consomme `shapesProps`. Le canevas lui-même ne monte ni barre de titre, ni ruban, ni commandes de page, ni volet des formes ou des gabarits, ni inspecteur, ni interface de fichier, ni récupération, ni boîtes de dialogue de paramètres ou de partage. Les limites existantes de modification et de protection des VSDX s'appliquent toujours.

## Autorisation et mode {#permission-and-mode}

`canEdit` suit les blocs de construction de PowerPoint : le dessin est **en lecture seule tant que l'hôte ne transmet pas `canEdit: true`**, et l'autorisation s'applique à la création du contrôleur, avant le montage du canevas, de sorte que rien n'est modifiable un bref instant (le visualiseur complet conserve son mode modifiable par défaut). Retirer l'autorisation annule les modifications en attente, supprime l'outil de dessin et retire les poignées de redimensionnement, de rotation et de points d'extrémité. La navigation entre les pages, la sélection, la recherche, le défilement et le déplacement, le zoom, l'ajustement, la copie et la sérialisation restent disponibles. La prise en charge de la modification selon le format source est distincte de l'autorisation accordée par l'hôte, c'est pourquoi `mode` vérifie aussi cette prise en charge.

## Ce que peut faire le canevas {#what-the-canvas-can-do}

Le canevas autonome utilise les gestes du visualiseur complet (un câblage commun, et non une copie) : déplacement, texte en ligne (double-clic, F2 ou saisie sur une forme sélectionnée), poignées de redimensionnement et de rotation, points d'extrémité des lignes, outils de dessin, outil texte et outil connecteur choisi avec `onToolChange`. L'outil texte trace une zone et ouvre un brouillon de texte brut (Ctrl+Entrée ajoute la zone de texte, Échap annule) ; l'outil connecteur trace de forme à forme ou de point de connexion à point de connexion et colle les extrémités exactement comme le fait le visualiseur complet, avec un tracé à angle droit. Supprimer, Ctrl+Z, Ctrl+Y et Échap (quitter l'outil de dessin) fonctionnent sur le canevas qui a le focus, et les événements copier, couper et coller du navigateur sont pris en compte.

Chacune de ces actions est soumise à `canEdit`, exactement comme dans le visualiseur complet. Les refus de modification sont annoncés dans une zone live polie à l'intérieur du canevas.

## Interface optionnelle de la bibliothèque {#optional-library-ui}

`Toolbar` affiche chaque commande que prend en charge `toolbarProps` : historique, zoom et ajustement, dupliquer, grouper, dissocier et supprimer, couper/copier/coller, outils de dessin, recherche et les boutons de boîte de dialogue de l'hôte décrits ci-dessous, avec leur état d'activation. Elle est facultative : passez `include` (`'file' | 'history' | 'zoom' | 'selection' | 'clipboard' | 'tools' | 'search'`) pour n'afficher que certains groupes, ou ignorez-la et reliez `toolbarProps` à vos propres commandes. `PageNavigation` est le sélecteur de page.

`ShapesPane` est la fenêtre des formes (gabarits) : Rechercher des formes, le gabarit de document du dessin et les gabarits intégrés. Cliquer sur un gabarit, ou le déposer sur le canevas monté, passe par le même flux de création que le visualiseur complet, et donc par la même autorisation. Montez le canevas sur le même contrôleur ; sans canevas, les gabarits sont ajoutés au centre de la page.

## Boîtes de dialogue gérées par l'hôte {#host-owned-dialogs}

Comme les rappels `onOpen...` de PowerPoint, les boîtes de dialogue appartiennent à l'hôte : la bibliothèque n'en monte jamais en mode surface. Transmettez l'un des rappels `onOpenFile`, `onOpenPageSetup`, `onOpenSettings` et `onOpenShareDialog` au hook ou à la fabrique. Ils reviennent dans `toolbarProps` avec `canOpenFile`, `canOpenPageSetup`, `canOpenSettings` et `canOpenShareDialog` ; un rappel que vous omettez est absent et son indicateur vaut `false` (non disponible), et la `Toolbar` de la bibliothèque n'affiche un bouton que pour ceux que vous avez fournis.

## Persistance et suivi des modifications {#persistence-and-dirty-tracking}

`controller.exportVsdx()` renvoie les octets et ne modifie jamais `dirty`. Une fois que l'hôte les a enregistrés, il appelle `markClean()` (sur le résultat du hook, sur le contrôleur ou sur le handle). Cela efface `state.edit.dirty`, déclenche `onDirtyChange(false)` et un événement `document-change` de type `save`. Les modifications et étapes d'annulation ultérieures sont mesurées à partir de ce point enregistré.

## Enregistrer pendant la modification du texte d'une forme {#saving-while-editing-text}

Si l'utilisateur appuie sur un Enregistrer géré par l'hôte pendant qu'il tape dans une forme, le brouillon n'est pas encore dans le document. Appelez d'abord `await commitTextEdit()` (sur le résultat du hook, la session, les blocs vanilla, le service Angular ou le contrôleur/handle), puis `exportVsdx()`. L'appel ne se résout qu'une fois la modification du brouillon terminée, y compris si un blur ou Échap l'avait déjà lancée, et renvoie un `TextEditOutcome` :

- `committed` : le brouillon a été appliqué comme une modification annulable ; `text` est ce qui a été appliqué.
- `unchanged` (`reason` : `no-draft` ou `same-text`) : rien à appliquer.
- `refused` (`reason` : `read-only` ou `rejected`) : l'éditeur l'a refusé ; `message` donne la raison et `text` conserve le brouillon pour le récupérer.
- `cancelled` (`reason` : `stale-document`, `superseded` ou `destroyed`) : le document du brouillon a disparu ou la modification a été remplacée ; `text` est le brouillon.

Chaque résultat contient `sourceGeneration` et `documentGeneration`, ainsi que `pageId` et `shapeId` lorsqu'il y avait un brouillon. Un brouillon appartient au document dans lequel il a été tapé : charger un autre fichier ou démonter le canevas l'abandonne, il n'est donc jamais appliqué à un autre document. Si un chargement a abandonné un brouillon non enregistré, le prochain appel à `commitTextEdit()` le renvoie une seule fois comme `cancelled` / `stale-document` avec son `text`, son `pageId`, son `shapeId` et l'ancienne `sourceGeneration`, afin qu'un Enregistrer concurrent d'un changement de document puisse récupérer ce qui a été tapé. La validation et l'annulation/rétablissement suivent le chemin de modification normal.

## Paquets {#packages}

Chaque paquet de framework réexporte `ooxml-ui/visio`, de sorte que `mountDiagramCanvas(host, controller)`, `mountDiagramInspector(host, controller)`, `mountDiagramShapes(host, controller, { locale? })`, `diagramToolbarProps(controller, fit, locale?, hostCallbacks?)` et `fitDiagramCanvas` sont disponibles depuis n'importe lequel d'entre eux. Abonnez-vous via le contrôleur existant et appelez les fonctions de désabonnement renvoyées au démontage. Aucun second magasin d'état n'est introduit.

Les paquets du visualiseur et `ooxml-ui` sont publiés ensemble, et chaque paquet de visualiseur dépend de la version de `ooxml-ui` avec laquelle il a été publié ; l'API de composition est donc toujours présente. Un hôte qui mélange les versions à la main reçoit une erreur explicite d'un `ooxml-ui` plus ancien, au lieu d'afficher par accident une interface d'application.

## Collaboration {#collaboration}

Les blocs de construction n'acceptent pas d'option `collaboration` : partager un dessin dans une salle est une propriété du `<visio-viewer>` complet (voir [collaboration](/fr/collaboration)). Une interface personnalisée peut proposer sa propre entrée de partage via `onOpenShareDialog`.

## Composition native par framework {#native-framework-composition}

La composition propre à chaque framework réside dans chaque paquet de framework. Tous les adaptateurs possèdent un `ViewerController`, exposent son état par la réactivité du framework, annulent les chargements remplacés, et détruisent les abonnements et les gestionnaires du canevas avec leur propriétaire. Un canevas retiré indépendamment peut se remonter sur ce même contrôleur sans recharger la source.

### Vue {#vue}

```vue
<script setup lang="ts">
import {
	DiagramCanvas,
	Toolbar,
	PageNavigation,
	Inspector,
	useViewerBuildingBlocks,
} from 'visio-vue-viewer';
const props = defineProps<{ content?: Uint8Array; canEdit?: boolean }>();
const blocks = useViewerBuildingBlocks(() => ({ content: props.content, canEdit: props.canEdit }));
const { toolbarProps, loading, error, mode } = blocks;
defineExpose({ controller: blocks.controller });
</script>
<DiagramCanvas v-bind="blocks.canvasProps" />
<Toolbar :toolbar-props="toolbarProps" />
<PageNavigation :toolbar-props="toolbarProps" />
<Inspector v-bind="blocks.inspectorProps" />
```

`toolbarProps`, `state`, `loading`, `error` et `mode` sont des refs ; les templates Vue les déballent automatiquement. Passez un getter d'options ou une ref réactive pour suivre les changements de source, d'autorisation et de langue.

### Solid {#solid}

```tsx
import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-solid-viewer';
function EmbeddedDiagram(props: { content?: Uint8Array; canEdit?: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		content: props.content,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<DiagramCanvas {...blocks.canvasProps} />
			<Toolbar {...blocks.toolbarProps()} />
		</>
	);
}
```

`toolbarProps`, `state`, `loading`, `error` et `mode` sont des accesseurs. Le contrôleur est disponible via `blocks.controller` ; la destruction suit le propriétaire Solid. `PageNavigation` accepte le même descripteur de barre d'outils à plat, `Inspector` accepte `inspectorProps` et `ShapesPane` accepte `shapesProps`.

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
 import { writable } from 'svelte/store';
 import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-svelte-viewer';
 const options = writable({ canEdit: true });
 const blocks = createViewerBuildingBlocks(options);
 const toolbarProps = blocks.toolbarProps;
 // options.set({ canEdit: false }) updates permissions through the store subscription.
</script>
<DiagramCanvas {...blocks.canvasProps} />
<Toolbar toolbarProps={$toolbarProps} />
```

L'argument d'options accepte un store en lecture ou un objet initial simple ; les hôtes qui utilisent un objet simple peuvent appeler `blocks.update(options)` depuis un effet de rune. `toolbarProps`, `state`, `loading`, `error` et `mode` sont des stores en lecture. `PageNavigation` prend `toolbarProps`, `Inspector` prend `inspectorProps` et `ShapesPane` prend `shapesProps`. Le contrôleur est disponible via `blocks.controller`.

### JavaScript vanilla {#vanilla-javascript}

```ts
import { createViewerBuildingBlocks } from 'visio-vanilla-viewer';
const blocks = createViewerBuildingBlocks({ canEdit: false, onOpenSettings: openSettings });
blocks.mountCanvas(document.querySelector('#canvas')!);
blocks.mountToolbar(document.querySelector('#toolbar')!, ['history', 'zoom']);
blocks.mountPageNavigation(document.querySelector('select')!);
blocks.mountInspector(document.querySelector('#inspector')!);
blocks.mountShapes(document.querySelector('#shapes')!);
blocks.update({ content: bytes, canEdit: true });
blocks.subscribe(() => render(blocks.toolbarProps, blocks.mode));
// later: blocks.markClean(); blocks.destroy();
```

La fabrique vanilla exécute la même session que les autres adaptateurs. `toolbarProps`, `state`, `loading`, `error` et `mode` sont des getters, `subscribe(listener)` signale chaque changement d'état ou d'option, et chaque aide `mount...` renvoie sa fonction de désabonnement ; `destroy()` libère tout ce qui reste monté. La barre d'outils montée met à jour ses éléments sur place, de sorte que le focus survit aux rafraîchissements.

### Angular {#angular}

```ts
import { Component, Input, inject, type OnChanges } from '@angular/core';
import {
	ViewerBuildingBlocks,
	DiagramCanvasComponent,
	ToolbarComponent,
} from 'visio-angular-viewer';
@Component({
	selector: 'embedded-diagram',
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [DiagramCanvasComponent, ToolbarComponent],
	template: '<visio-diagram-canvas /><visio-diagram-toolbar />',
})
export class EmbeddedDiagram implements OnChanges {
	@Input() content?: Uint8Array;
	@Input() canEdit = false;
	readonly blocks = inject(ViewerBuildingBlocks);
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit });
	}
}
```

Le service délimité expose les signaux en lecture seule `state`, `toolbarProps`, `loading`, `error` et `mode`, le contrôleur existant, ainsi que les émetteurs d'événements `documentChange`, `dirtyChange` et `documentError`. `DiagramCanvasComponent`, `ToolbarComponent`, `PageNavigationComponent`, `InspectorComponent` et `ShapesPaneComponent` injectent ce propriétaire. Placez le fournisseur sur chaque hôte d'éditeur, plutôt qu'au niveau de toute l'application, pour isoler plusieurs documents. Vous pouvez aussi envelopper les commandes dans `<visio-building-blocks [content]="bytes" [canEdit]="allowed" (dirtyChange)="saveNeeded = $event">` avec `ViewerBuildingBlocksComponent`, qui fournit les entrées, les sorties et un fournisseur délimité.

Toutes les commandes facultatives s'affichent avec des modèles natifs du framework. Les hôtes peuvent les omettre et appeler les actions à plat de la barre d'outils depuis leur propre interface. Les rappels de la barre d'outils conservent les contrôles d'autorisation, de chargement et de format source du contrôleur. Des tests de régression dans le navigateur couvrent les six compositions natives, y compris la modification, annuler/rétablir, le retrait d'autorisation, la sérialisation et le rechargement, la reprise sur erreur et le démontage.

Pour un hôte qui utilise déjà les signaux Angular, `createViewerBuildingBlocks(() => ({
content: bytes(), canEdit: allowed() }))` injecte le service `ViewerBuildingBlocks` délimité et relie le getter d'options à un effet Angular. Appelez-la dans le contexte d'injection du composant. Les émetteurs d'événements du service sont actifs même avant le premier `update` explicite.

Changer la prop `mount` d'un canevas autonome libère son propriétaire précédent avant d'attacher le suivant au même hôte. Changer le contrôleur d'un inspecteur remplace de la même façon son abonnement. Cela permet de passer d'une session d'édition indépendante à une autre sans gestionnaires d'entrée ni abonnements d'état orphelins.

## Thème Office partagé {#shared-office-theme}

Chaque barre d'outils optionnelle, sélecteur de page et inspecteur natif suit les mêmes variables CSS `--office-*` que le canevas du dessin. Les blocs installent automatiquement leur feuille de style partagée ; il n'existe aucun objet de thème propre à un framework ni de second magasin d'état du document. Définissez les thèmes sur votre propre élément hôte :

```html
<section
	data-office-theme="dark"
	style="--office-accent: #a78bfa; --office-font: Inter, sans-serif"
>
	<!-- DiagramCanvas, Toolbar, PageNavigation and Inspector -->
</section>
```

Utilisez `data-office-theme="light"`, `"dark"` ou `"auto"` sur la racine du document ou sur une section qui la contient. `auto` suit la préférence du système, y compris ses changements en cours d'exécution. Des hôtes clair/sombre délimités peuvent coexister sur une même page. Sans attribut délimité, les blocs héritent du thème racine et des variables de l'hôte à travers la frontière d'ombre du canevas. Changer le thème ne recharge pas le canevas, ne réinitialise pas la sélection et ne modifie ni les couleurs du document ni les octets sérialisés.

Pour une palette propre à un hôte, remplacez `--office-background`, `--office-surface`, `--office-foreground`, `--office-muted-foreground`, `--office-border`, `--office-accent`, `--office-ring` et `--office-font` sur la section englobante. Les remplacements Visio historiques `--vv-surface`, `--vv-ink`, `--vv-border`, `--vv-accent`, `--vv-focus` et `--vv-font` restent pris en charge, avec priorité sur leurs équivalents canoniques. Le canevas conserve en plus `--vv-background`, `--vv-secondary` et les remplacements existants de la barre d'état. La page du diagramme garde son propre aspect de document tandis que les commandes qui l'entourent suivent le thème de l'interface.
