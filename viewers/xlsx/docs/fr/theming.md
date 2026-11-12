# Thèmes

L’habillage d’Excel et les contrôles natifs des frameworks partagent le système de design `--office-*` avec Word,
PowerPoint et Visio. Installez le thème commun une seule fois, puis utilisez un ancêtre commun pour habiller
à la fois le canevas et les contrôles facultatifs. Les contrôles natifs et les éditeurs montés l’installent aussi
de façon idempotente lorsqu’ils sont utilisés seuls.

```ts
import { installOfficeUiTheme } from 'ooxml-ui/theme';
installOfficeUiTheme();
```

```html
<section data-office-theme="dark" class="workbook-host">
	<!-- Render the native Toolbar, FormulaBar, SpreadsheetCanvas and SheetTabs here. -->
</section>
```

`data-office-theme` accepte `light`, `dark` et `auto`. Il peut être défini sur la racine du document ou
sur n’importe quel ancêtre à portée limitée. `auto` suit le système d’exploitation. Un canevas dont
`theme="auto"` est la valeur par défaut hérite de la palette de son ancêtre. Un `theme="light"` ou `theme="dark"`
explicite crée une portée locale sur cet éditeur, en conservant l’API existante de propriétés et d’attributs.
Cela ne modifie ni les contrôles frères de l’hôte ni l’application parente.

Les jetons canoniques peuvent être remplacés au même niveau :

```css
.workbook-host {
	--office-accent: #1f9d63;
	--office-ring: #1f9d63;
	--office-background: #0f172a;
	--office-surface: #1e293b;
	--office-foreground: #f8fafc;
	--office-border: #64748b;
	--office-radius: 6px;
	--office-font: 'Segoe UI', sans-serif;
}
```

Les contrôles natifs facultatifs adoptent les styles communs grâce aux marqueurs `data-office-block`. Leurs
champs, boutons, états sélectionnés, états désactivés et anneaux de focus suivent les mêmes jetons.
La feuille de style ne cible pas les contrôles personnalisés de l’hôte non marqués. Les changements de thème mettent à jour le CSS
sur place, en préservant le moteur courant, la cellule active, le brouillon, la sélection et le champ ayant le focus.
Les couleurs des cellules du classeur et le thème du document restent inchangés ; le fond de la feuille est blanc par défaut.
`--office-paper` et l’ancien `--xve-sheet-bg` sont des surcharges de présentation.

## Surcharges de jetons {#token-overrides}

`themeColors` reste pris en charge pour les surcharges par éditeur. Chaque clé devient une propriété personnalisée `--xve-<clé-en-kebab-case>` sur l’élément hôte, avec le même ensemble de jetons que celui de l’éditeur Word (`--dve-*` là-bas), afin qu’un seul système de design puisse piloter les deux :

| Clé                                    | Utilisée pour                                    |
| -------------------------------------- | ------------------------------------------------ |
| `background`, `foreground`             | Canevas et texte de l’éditeur                    |
| `card`, `cardForeground`               | Ruban et panneaux                                |
| `popover`, `popoverForeground`         | Menus et listes déroulantes                      |
| `primary`, `primaryForeground`         | Accent (barre de titre, sélection, onglet actif) |
| `secondary`, `secondaryForeground`     | Boutons secondaires                              |
| `muted`, `mutedForeground`             | En-têtes, indications, texte désactivé           |
| `accent`, `accentForeground`           | États au survol et enfoncés                      |
| `destructive`, `destructiveForeground` | Erreurs                                          |
| `border`, `input`, `ring`              | Lignes, champs et anneaux de focus               |
| `radius`                               | Rayon des coins                                  |

```ts
editor.themeColors = { primary: '#1f9d63', ring: '#1f9d63', radius: '6px' };
```

Vous pouvez aussi définir les propriétés personnalisées en CSS :

```css
xlsx-editor {
	--xve-primary: #1f9d63;
}
```

## Suivre le thème d’un site {#following-a-site-theme}

La démo suit la clé `vitepress-theme-appearance` de VitePress dans `localStorage` et définit
`data-office-theme` sur la racine du document. Liez ce même attribut au réglage de votre application afin que les contrôles composés et les éditeurs
changent ensemble. L’éditeur ne lit jamais `localStorage` lui-même.

Les surcharges CSS `--xve-*` héritées restent prises en charge sur l’éditeur ou sur un ancêtre. Elles l’emportent sur le
jeton canonique correspondant lorsqu’elles sont fournies. Les primitives communes internes
reçoivent ces valeurs résolues sans redéclarer la palette canonique sur l’hôte de l’éditeur. Les objets
prédéfinis clair/sombre existants et les clés de `themeColors` conservent leur forme publique.
