---
title: Thèmes
description: Le système de thème ViewerTheme commun aux liaisons - la forme de l'objet thème, la manière dont themeToCssVars produit les propriétés personnalisées --pptx-*, les préréglages intégrés, l'application par liaison et le sélecteur Fichier > Options > Apparence.
---

# Thèmes

L'interface de la visionneuse (barre d'outils, ruban, boîtes de dialogue, backstage) est thématisée au moyen d'un objet `ViewerTheme` : un ensemble de propriétés personnalisées CSS (`--pptx-*`) appliquées à la racine de la visionneuse. Cela est totalement distinct du jeu de couleurs et de polices OOXML d'une présentation (la galerie « Thèmes » de l'onglet Création modifie le document `.pptx` lui-même) ; `ViewerTheme` n'affecte que l'interface de l'application.

Le système de thèmes est indépendant du framework. Les types, les valeurs par défaut, les préréglages et les fonctions d'aide ci-dessous sont implémentés une seule fois (dans l'API publique `ooxml-ui/pptx`) et réexportés à l'identique par chaque liaison publiée : `pptx-react-viewer`, `pptx-vue-viewer`, `pptx-angular-viewer`, `pptx-svelte-viewer` et `pptx-vanilla-viewer`.

## Apparence Office partagée {#shared-office-appearance}

Les cinq liaisons héritent des mêmes jetons `--office-*` que Word, Excel et Visio. Placez `data-office-theme="light"`, `"dark"` ou `"auto"` sur la racine du document ou sur un élément qui la contient. Les modes délimités remplacent la palette héritée ; `auto` suit le système d'exploitation lorsqu'il change. Les jetons personnalisés posés sur la même portée remplacent la palette choisie :

```html
<section data-office-theme="dark" style="--office-accent: #a78bfa; --office-font: Georgia, serif">
	<!-- Mount a viewer here, without an explicit ViewerTheme palette. -->
</section>
```

La barre d'outils, le chrome natif et les contrôles en shadow DOM héritent de la portée la plus proche. Parmi les jetons utiles figurent `--office-background`, `--office-surface`, `--office-foreground`, `--office-border`, `--office-accent`, `--office-ring` et `--office-font`. Les sorties existantes de `ViewerTheme`, `themeToCssVars()` et `defaultCssVars()` restent compatibles. Les anciennes entrées explicites continuent de l'emporter sur les valeurs de repli des jetons canoniques, y compris les palettes partielles et les dimensions de contrôles telles que `--pptx-switch-width`.

L'apparence de l'interface ne modifie pas les couleurs, les polices ni le thème du document d'une présentation.

## La forme de `ViewerTheme` {#the-viewertheme-shape}

```ts
interface ViewerTheme {
	/** Semantic UI colors. Each key maps to a `--pptx-<key>` custom property. */
	colors?: Partial<ViewerThemeColors>;
	/** Base border-radius value (e.g. "0.5rem", "8px"). */
	radius?: string;
	/** Escape hatch: arbitrary CSS custom properties set on the viewer root. Keys include the `--` prefix. */
	cssVars?: Record<string, string>;
}
```

Chaque champ est facultatif ; les valeurs non définies héritent des jetons d'interface Office partagés au niveau de la visionneuse. Les couleurs explicites de `ViewerTheme` et les anciennes variables `--pptx-*` sont prioritaires. `ViewerThemeColors` compte 19 jetons sémantiques, nommés selon la convention shadcn/ui. Toute chaîne de couleur CSS valide est acceptée (hexadécimal, `rgb()`, `hsl()`, `oklch()`, couleurs nommées) :

| Jetons                                 | Rôle                                          |
| -------------------------------------- | --------------------------------------------- |
| `background`, `foreground`             | Fond de la racine et texte par défaut         |
| `card`, `cardForeground`               | Surfaces de cartes et de panneaux             |
| `popover`, `popoverForeground`         | Fenêtres contextuelles et menus déroulants    |
| `primary`, `primaryForeground`         | Actions principales (boutons, états actifs)   |
| `secondary`, `secondaryForeground`     | Actions secondaires                           |
| `muted`, `mutedForeground`             | Surfaces discrètes et texte secondaire        |
| `accent`, `accentForeground`           | Surfaces de surbrillance au survol            |
| `destructive`, `destructiveForeground` | Actions dangereuses / de suppression          |
| `border`, `input`, `ring`              | Bordures, bordures de champs, anneau de focus |

Consultez la [page de thème React](/react/theming#viewertheme-and-viewerthemecolors) pour le tableau complet, jeton par jeton, avec le nom exact de chaque variable CSS.

## Comment `themeToCssVars` produit les variables `--pptx-*` {#how-themetocssvars-produces-pptx-variables}

`themeToCssVars(theme, omitDefaults = false)` convertit un `ViewerTheme` en un `Record<string, string>` plat de propriétés personnalisées CSS, prêt à être appliqué à la racine de la visionneuse comme style en ligne. Chaque liaison l'appelle en interne lorsque vous passez un thème ; elle est aussi exportée pour construire vos propres outils.

- Chaque clé de `colors` devient `--pptx-<clé-en-kebab-case>` : `primaryForeground: '#fff'` produit `--pptx-primary-foreground: #fff`.
- Chaque couleur est aussi recopiée dans le jeton sémantique Tailwind correspondant (`--color-primary-foreground`), pour la compatibilité avec les hôtes Tailwind CSS v4. La visionneuse utilise `@theme inline` afin que les jetons Office hérités se résolvent au niveau du composant qui les consomme.
- `radius` devient `--pptx-radius`, avec les valeurs dérivées `--radius-sm` / `--radius-md` / `--radius-lg` / `--radius-xl` (`calc(r - 4px)` à `calc(r + 4px)`).
- Les entrées de `cssVars` passent telles quelles.
- Avec `omitDefaults: true`, les valeurs égales aux valeurs par défaut intégrées sont ignorées.

```ts
import { themeToCssVars, defaultCssVars } from 'pptx-react-viewer';

themeToCssVars({ colors: { primary: '#6366f1' }, radius: '0.75rem' });
// {
//   '--pptx-primary': '#6366f1', '--color-primary': '#6366f1',
//   '--pptx-radius': '0.75rem',
//   '--radius-sm': 'calc(0.75rem - 4px)', ... '--radius-xl': 'calc(0.75rem + 4px)',
// }

defaultCssVars();
// The complete set of --pptx-* properties with the built-in dark defaults,
// for generating a full fallback stylesheet.
```

## Thèmes et préréglages intégrés {#built-in-themes-and-presets}

| Export                                         | Palette                                                                                                                                   |
| ---------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `defaultThemeColors` + `defaultRadius`         | L'interface sombre intégrée (échelle de gris Tailwind, primaire indigo, rayon `0.5rem`). Disponible explicitement via `defaultCssVars()`. |
| `vermilionLightTheme` / `vermilionLightColors` | Palette claire chaude « papier » avec l'accent vermillon utilisé par ce site de documentation.                                            |
| `vermilionDarkTheme` / `vermilionDarkColors`   | Palette sombre atténuée « salle du présentateur » avec le même accent.                                                                    |
| `vermilionRadius`                              | `'0.375rem'`, le rayon utilisé par les deux préréglages vermillon.                                                                        |

Les deux préréglages vermillon sont des objets `ViewerTheme` complets (les 19 jetons plus le rayon), ils remplacent donc entièrement les valeurs sombres par défaut. Les palettes brutes `*Colors` sont exportées à côté pour dériver des variantes :

```ts
import { vermilionDarkColors, vermilionRadius } from 'pptx-react-viewer';
import type { ViewerTheme } from 'pptx-react-viewer';

const custom: ViewerTheme = {
	colors: { ...vermilionDarkColors, primary: '#38bdf8' },
	radius: vermilionRadius,
};
```

Une palette claire non vermillon existe aussi, mais elle n'est pas un export nommé des liaisons ; on y accède par le catalogue de thèmes : `resolveThemeCatalogEntry('light')`.

## Appliquer un thème par liaison {#applying-a-theme-per-binding}

Chaque liaison accepte le même objet `ViewerTheme` ; seul le mécanisme de transmission diffère.

::: code-group

```tsx [React]
import { PowerPointViewer, vermilionDarkTheme } from 'pptx-react-viewer';

<PowerPointViewer content={bytes} theme={vermilionDarkTheme} />;
```

```vue [Vue]
<script setup lang="ts">
import { PowerPointViewer, vermilionDarkTheme } from 'pptx-vue-viewer';
</script>

<template>
	<PowerPointViewer :content="bytes" :theme="vermilionDarkTheme" />
</template>
```

```ts [Angular]
// <pptx-viewer [content]="bytes" [theme]="theme" />
import { vermilionDarkTheme } from 'pptx-angular-viewer';

export class DeckComponent {
	theme = vermilionDarkTheme;
}
// Or share one theme across a subtree without the input:
// providers: [provideViewerTheme(vermilionDarkTheme)]
```

```svelte [Svelte]
<script lang="ts">
	import { PowerPointViewer, vermilionDarkTheme } from 'pptx-svelte-viewer';
</script>

<PowerPointViewer content={bytes} theme={vermilionDarkTheme} />
```

```ts [Vanilla]
import { createPptxViewer, vermilionDarkTheme } from 'pptx-vanilla-viewer';

const viewer = createPptxViewer(host, {
	source: bytes,
	theme: vermilionDarkTheme,
});
// Change later at runtime:
viewer.setTheme({ colors: { primary: '#38bdf8' } });
```

:::

Pour partager un même thème entre plusieurs visionneuses, React exporte `ViewerThemeProvider` / `useViewerTheme`, Vue exporte `provideViewerTheme` / `useViewerTheme`, et Angular exporte `provideViewerTheme` / le jeton d'injection `VIEWER_THEME`.

## Modes de style {#styling-modes}

L'interface de la visionneuse référence les propriétés personnalisées `--pptx-*` pour chaque jeton visuel, ce qui permet trois configurations de style (par ordre croissant de contrôle) :

| Mode                      | Configuration                                                                                                                                                                                                                                                    |
| ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Hôte Tailwind CSS v4      | Aucun import CSS requis ; les classes de la visionneuse se résolvent par votre configuration existante. Remplacez les valeurs avec la prop/l'entrée `theme`.                                                                                                     |
| Feuille de style intégrée | `import 'pptx-react-viewer/styles'` (les packages Vue et Angular exposent aussi `/styles` et `/styles.css`). Le package Svelte compile ses styles dans ses composants. Fournit toutes les classes utilitaires requises ainsi que les valeurs sombres par défaut. |
| Propriétés CSS brutes     | Définissez vous-même les propriétés `--pptx-*` (voir `defaultCssVars()` pour la liste complète) et n'utilisez ni l'un ni l'autre.                                                                                                                                |

La liaison vanilla fait exception : `createPptxViewer` injecte automatiquement sa propre feuille de style délimitée (idempotente, `#pptx-vanilla-viewer-styles`). Les hôtes à CSP stricte peuvent plutôt pré-générer la chaîne depuis `getViewerCss()`.

## Habillage des contrôles de sélection {#selection-control-artwork}

L'habillage de la sélection utilise le champ existant `ViewerTheme.cssVars`, et non un objet de thème distinct ni une prop de composant, dans les cinq liaisons. Ces jetons facultatifs ne modifient pas l'apparence par défaut d'une liaison lorsqu'ils sont omis. Ils affectent les contrôles de l'éditeur, non le contenu des diapositives, la taille des polices, la géométrie définie ni les documents exportés.

| Propriété CSS personnalisée            | Contrôle                                                        |
| -------------------------------------- | --------------------------------------------------------------- |
| `--pptx-selection-corner-size`         | Largeur et hauteur de l'habillage des coins                     |
| `--pptx-selection-corner-radius`       | Rayon de bordure de l'habillage des coins                       |
| `--pptx-selection-edge-length`         | Longueur de l'habillage des bords, parallèlement au bord        |
| `--pptx-selection-edge-thickness`      | Épaisseur de l'habillage des bords, perpendiculairement au bord |
| `--pptx-selection-edge-radius`         | Rayon de bordure de l'habillage des bords                       |
| `--pptx-selection-handle-fill`         | Remplissage de l'habillage de redimensionnement                 |
| `--pptx-selection-handle-border-color` | Couleur de bordure de l'habillage Redimensionner et Pivoter     |
| `--pptx-selection-outline-color`       | Contour de sélection et tige existante de Pivoter               |
| `--pptx-selection-rotate-size`         | Diamètre de l'habillage circulaire Pivoter                      |
| `--pptx-selection-rotate-fill`         | Remplissage de l'habillage Pivoter                              |
| `--pptx-selection-rotate-foreground`   | Couleur du glyphe Pivoter, lorsqu'un glyphe est dessiné         |

Utilisez des longueurs en pixels positives comme `6px` pour les tailles, des longueurs CSS ou des pourcentages pour les rayons, et des couleurs CSS valides pour les couleurs. Un repli `var()` gère un jeton omis, mais non une valeur CSS arbitraire invalide. La visionneuse n'analyse ni n'assainit ces chaînes CSS. Supprimez une surcharge pour restaurer la valeur par défaut de la liaison. Les valeurs de rayon utilisent les coordonnées CSS locales du contrôle, comme sa bordure et son ombre existantes ; elles ne sont pas compensées séparément par le zoom de la scène. Utilisez `0px` pour des coins carrés ou `50%` pour une forme ronde proportionnelle dans toutes les liaisons.

```ts
const theme = {
	cssVars: {
		'--pptx-selection-corner-size': '6px',
		'--pptx-selection-corner-radius': '0px',
		'--pptx-selection-edge-length': '6px',
		'--pptx-selection-edge-thickness': '6px',
		'--pptx-selection-edge-radius': '0px',
		'--pptx-selection-handle-fill': '#ffffff',
		'--pptx-selection-handle-border-color': '#6366f1',
		'--pptx-selection-outline-color': '#6366f1',
		'--pptx-selection-rotate-fill': '#ffffff',
		'--pptx-selection-rotate-foreground': '#6366f1',
	},
};
```

Passez ce même thème dans les exemples propres à chaque liaison ci-dessus. Pour une coque personnalisée, appliquez `themeToCssVars(theme)` à l'ancêtre commun de la diapositive et de ses contrôles de sélection. Ne ciblez pas les classes enfants privées et ne supposez pas que les poignées sont imbriquées dans la forme sélectionnée.

L'habillage reste centré sur les ancres existantes. Un habillage plus petit ne réduit pas la cible invisible de la souris ou du tactile d'origine. Un habillage plus grand agrandit son cadre, mais la répartition existante des zones de survol avec les voisins s'applique encore sur les formes minuscules ; des visuels surdimensionnés peuvent se chevaucher sans prendre la zone d'entrée d'une autre poignée. Ce contrat n'expose ni les tailles de zone de survol ni les décalages de Pivoter. Gardez un habillage personnalisé volumineux raisonnable pour les formes éditées. Les losanges d'ajustement et les indicateurs d'extrémité de connecteur conservent leur signification et leur apparence distinctes.

## Fichier > Options > Apparence {#file-options-appearance}

La boîte de dialogue Paramètres de chaque liaison comporte un onglet **Apparence** : une petite galerie de préréglages de thème intégrés (Par défaut, Clair, Vermillon clair, Vermillon sombre) qu'un utilisateur peut parcourir à l'exécution, définie par l'export `THEME_CATALOG` :

```ts
interface ThemeCatalogEntry {
	key: string; // 'default' | 'light' | 'vermilionLight' | 'vermilionDark'
	labelKey: string; // pptx.* translation key for the entry's label
	theme: ViewerTheme | undefined; // undefined = reset to the built-in default
}
```

Il s'agit volontairement d'une liste courte et sélectionnée, pas d'une galerie complète ; passez `availableThemes` (ci-dessous) pour proposer davantage ou moins de choix. `resolveThemeCatalogEntry(key, catalog?)` recherche une entrée par sa clé.

### Priorité : un `theme` explicite l'emporte toujours {#precedence-an-explicit-theme-always-wins}

Si vous passez une prop/entrée `theme` à la visionneuse, elle reste prioritaire sur tout choix fait dans l'onglet Apparence ; le sélecteur est inactif tant que `theme` est défini. C'est voulu : un hôte qui gère son propre thème (par exemple en le synchronisant avec un interrupteur de mode sombre de toute l'application) ne doit pas être remplacé par un clic dans la visionneuse.

L'onglet Apparence n'a d'effet que si vous ne passez **pas** de `theme` explicite. Dans ce mode autonome, l'ordre de résolution est :

1. la prop `defaultThemeKey` (pour un point de départ autre que `'default'`) ;
2. un choix précédemment enregistré (`localStorage`, clé `pptx-viewer-prefs`) ;
3. l'entrée `'default'` du catalogue.

### Props du catalogue (toutes facultatives) {#catalog-props-all-optional}

| Prop               | Type                     | Rôle                                                                                                                                                                                                                                                                                                         |
| ------------------ | ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `defaultThemeKey`  | `string`                 | Clé initiale de `THEME_CATALOG`, utilisée seulement lorsque rien n'est encore enregistré.                                                                                                                                                                                                                    |
| `availableThemes`  | `ThemeCatalogEntry[]`    | Remplace le catalogue proposé dans Apparence (ajoutez vos propres préréglages, ou réduisez la liste).                                                                                                                                                                                                        |
| `onThemeChange`    | `(key: string) => void`  | Lorsqu'elle est fournie, **c'est vous** qui gérez l'enregistrement du choix ; la visionneuse cesse d'écrire dans `localStorage` et se contente de vous appeler.                                                                                                                                              |
| `defaultLocale`    | `string`                 | Même principe pour la Langue : code de langue initial lorsque rien n'est enregistré. React : ignorée à l'intérieur d'un `PptxI18nProvider`, dont `locale` détermine la langue.                                                                                                                               |
| `availableLocales` | `LocaleCatalogEntry[]`   | Remplace les langues proposées dans Options > Langue. React : par défaut, les codes enregistrés sur l'instance `react-i18next` de l'hôte ; à l'intérieur d'un `PptxI18nProvider`, il n'y a aucune valeur par défaut : passez deux langues ou plus, avec `onLocaleChange`, pour afficher la catégorie Langue. |
| `onLocaleChange`   | `(code: string) => void` | Lorsqu'elle est fournie, la visionneuse ne touche jamais à votre instance i18n ; elle se contente de vous appeler. React : à l'intérieur d'un `PptxI18nProvider`, l'hôte gère toujours la langue, et sans ce rappel la catégorie Langue est omise.                                                           |

React, Vue, Angular et Svelte utilisent tous exactement cette même forme. Vanilla fait exception : il disposait déjà des options de constructeur publiques `theme`/`locale` et des méthodes `setTheme()`/`setLocale()`, qui servent donc de valeur initiale à la place des options séparées `defaultThemeKey`/`defaultLocale` ; `availableThemes`, `availableLocales`, `onThemeChange` et `onLocaleChange` sont identiques partout ailleurs.

```tsx
// React - host owns theme persistence (e.g. syncing with app-wide dark mode)
<PowerPointViewer
	content={bytes}
	defaultThemeKey={systemPrefersDark ? 'vermilionDark' : 'vermilionLight'}
	onThemeChange={(key) => saveUserPreference('viewerTheme', key)}
/>
```

### Ajouter vos propres préréglages {#adding-your-own-presets}

Passez `availableThemes` avec vos propres `ThemeCatalogEntry[]` ; étendez les préréglages intégrés plutôt que de les remplacer, en utilisant la décomposition de `THEME_CATALOG` :

```ts
import { THEME_CATALOG } from 'pptx-react-viewer';

const availableThemes = [
	...THEME_CATALOG,
	{ key: 'brand', labelKey: 'app.theme.brand', theme: { colors: { primary: '#7c3aed' } } },
];
```

`labelKey` est recherché dans le dictionnaire i18n fourni par votre application (voir [Localisation](/fr/guide/localization)) ; pour une entrée personnalisée, enregistrez vous-même cette clé. Les clés non enregistrées reviennent à une étiquette lisible dérivée du dernier segment de la clé.

## Détails par liaison {#per-binding-details}

- [Thème React](/react/theming) - la prop `theme`, `ViewerThemeProvider`, et le tableau complet des jetons
- [Thème Vue](/vue/theming) - la prop `theme` et `provideViewerTheme`
- [Thème Angular](/angular/theming) - l'entrée `theme`, `provideViewerTheme` et le jeton `VIEWER_THEME`
- [Thème Svelte](/svelte/theming) - la prop `theme`
- [Thème Vanilla](/vanilla/theming) - l'option `theme`, `setTheme()` et `getViewerCss()`

## Étapes suivantes {#next-steps}

- [Localisation (i18n)](/fr/guide/localization) - l'onglet Langue fonctionne de la même manière, et c'est là que proviennent les clés de traduction `pptx.*` qui alimentent les étiquettes de thème.
- [Compte et connexion](/fr/guide/account) - l'éditeur de profil de Fichier > Compte enregistre aussi dans la même clé `localStorage` que le repli de thème/langue.
