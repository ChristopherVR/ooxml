---
title: Ajustement à la fenêtre d'affichage
description: Configurer l'ajustement des diapositives dans un hôte personnalisé sans modifier les dimensions de diapositive définies ni le zoom de l'utilisateur.
---

# Ajustement à la fenêtre d'affichage

Les cinq liaisons d'interface acceptent les options facultatives de l'hôte `fitPadding` et `maxFitScale`. Utilisez `fitPadding: 0` et `maxFitScale: null` lorsque l'hôte fournit déjà ses propres marges et que la diapositive doit pouvoir dépasser sa taille d'origine en pixels. Si vous omettez ces options, la politique d'ajustement existante est conservée.

## Options et valeurs par défaut {#options-and-defaults}

`fitPadding` est une marge en pixels CSS non mis à l'échelle **de chaque côté** : `8` réserve 16 pixels sur chaque axe ; `{ horizontal: 4, vertical: 16 }` réserve 8 pixels à l'horizontale et 32 à la verticale. `maxFitScale` est un plafond positif du facteur d'ajustement ; `null` signifie sans limite. Les valeurs invalides reviennent aux valeurs par défaut partagées.

Les cinq liaisons utilisent les mêmes valeurs par défaut, celles de l'éditeur React : sur une fenêtre d'affichage de bureau (768 px ou plus), la marge est de 4 / 16 px par côté et le facteur d'ajustement est plafonné à 1. Sur une fenêtre d'affichage plus étroite que 768 px, la marge est de 12 / 12 px par côté.

| Liaison | Marge horizontale / verticale par côté | Échelle d'ajustement maximale |
| ------- | -------------------------------------- | ----------------------------- |
| React   | 4 / 16 px (12 / 12 px sous 768 px)     | 1                             |
| Vue     | 4 / 16 px (12 / 12 px sous 768 px)     | 1                             |
| Angular | 4 / 16 px (12 / 12 px sous 768 px)     | 1                             |
| Svelte  | 4 / 16 px (12 / 12 px sous 768 px)     | 1                             |
| Vanilla | 4 / 16 px (12 / 12 px sous 768 px)     | 1                             |

Chaque package d'interface exporte les types `ViewportFitOptions` et `ViewportFitPadding`. Ces options affectent l'ajustement ordinaire de la visionneuse, et non les dimensions de diapositive définies, le contenu enregistré, les miniatures, l'export, le mode présentation, ni le réglage distinct du zoom utilisateur. Vanilla écrit toujours la marge résolue en ligne sur sa fenêtre d'affichage ; une règle de marge dans une feuille de style sur cet élément n'a donc aucun effet. Utilisez `fitPadding` à la place.

## Mesure {#measurement}

Chaque liaison ajuste à la taille cliente fractionnaire de la fenêtre d'affichage (sa zone de remplissage sans barres de défilement, à la précision avec laquelle le navigateur la met en page), et non aux `clientWidth` et `clientHeight` entiers en pixels. Une fenêtre d'affichage mise en page à 801,5 px est ajustée à 801,5 px ; ainsi, avec `fitPadding: 0`, la diapositive remplit exactement la fenêtre sans jamais la dépasser. Les bordures sont retirées à leur largeur effective, ce qui conserve leur fraction sous zoom du navigateur (une bordure de `1px` vaut 1,111 px à 90 %).

La taille exacte n'est utilisée que dans un contexte simple, où rien ne transforme ni ne zoome en CSS la fenêtre d'affichage (le zoom du navigateur convient). Là, le rectangle englobant de la fenêtre correspond à ses `width` et `height` calculés, exprimés dans ses propres pixels CSS. Lorsqu'une `transform` ou un `zoom` CSS sur la fenêtre d'affichage ou sur un ancêtre met à l'échelle ce rectangle, la diapositive est aussi mise en page et peinte à travers cette échelle, et son propre arrondi de mise en page peut décaler ses bords d'une fraction de pixel. L'ajustement réserve alors un pixel, de sorte que la diapositive reste dans la fenêtre d'affichage et que celle-ci ne défile pas. Un pixel est aussi réservé lorsque la fenêtre d'affichage affiche une barre de défilement, puisque la taille exacte ne peut alors pas être retrouvée.

Les changements ultérieurs sont pris en compte par des `ResizeObserver` sur la boîte de contenu et la boîte de bordure de la fenêtre d'affichage ; ainsi, une modification de la seule marge de la fenêtre réajuste aussi la diapositive.

## Configurer votre visionneuse {#configure-your-viewer}

Donnez à l'hôte une largeur et une hauteur réelles. Un enfant flex peut aussi avoir besoin de `min-height: 0` pour pouvoir s'ajuster dans l'espace disponible. Une marge d'ajustement nulle ne peut pas récupérer l'espace occupé par les barres d'outils, les panneaux latéraux ou les autres éléments du chrome de l'hôte.

### React {#react}

```tsx
<PowerPointViewer content={content} fitPadding={0} maxFitScale={null} />
```

Les mêmes options peuvent être passées à `useViewerBuildingBlocks` lorsque vous composez `Toolbar` et `SlideCanvas` dans une coque personnalisée. Par exemple :

```tsx
const { canvasProps, toolbarProps } = useViewerBuildingBlocks({
	content,
	canEdit: true,
	fitPadding: 0,
	maxFitScale: null,
});
```

Affichez les props renvoyées comme d'habitude ; l'ajustement utilise la fenêtre d'affichage mesurée du canevas, et non tout l'hôte, barre d'outils comprise. Les décalages des règles existants restent actifs. Définissez `showRulers={false}` sur un `SlideCanvas` personnalisé lorsqu'il doit s'aligner directement sur les bords de la fenêtre d'affichage.

`SlideCanvas` suit automatiquement sa fenêtre d'affichage, y compris les démontages et remontages conditionnels pendant que le hook de blocs de construction reste monté. Si vous utilisez `useZoomViewport` depuis `pptx-react-viewer/internals` avec votre propre élément de fenêtre d'affichage, liez `ref={zoom.setCanvasViewportNode ?? zoom.canvasViewportRef}` afin que l'ajustement et le zoom natif à la molette suivent l'élément courant. L'objet `canvasViewportRef` existant reste disponible pour lire la fenêtre d'affichage.

### Vue {#vue}

```vue
<PowerPointViewer :content="content" :fit-padding="0" :max-fit-scale="null" />
```

Les mêmes props sont disponibles sur `SlideCanvas`.

### Angular {#angular}

```html
<pptx-viewer [content]="content()" [fitPadding]="0" [maxFitScale]="null" />
```

### Svelte {#svelte}

```svelte
<PowerPointViewer {source} fitPadding={0} maxFitScale={null} />
```

### JavaScript Vanilla {#vanilla-js}

```ts
const viewer = createPptxViewer(container, {
	source,
	fitPadding: 0,
	maxFitScale: null,
});
```

La mise en page existante des règles est préservée dans chaque liaison. Désactivez les règles lorsque l'hôte a besoin d'un ajustement aligné sur les bords ; modifier `fitPadding` ne supprime pas l'espace des règles et ne change pas les autres contrôles de mise en page de la visionneuse.
