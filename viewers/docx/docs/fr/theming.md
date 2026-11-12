# Thèmes

L'éditeur applique son propre thème à son habillage (barre de titre, ruban, panneaux, barre d'état, boîtes de dialogue). Le papier du document reste toujours blanc, comme la page de Word.

## Clair, sombre et automatique {#light-dark-and-auto}

Définissez la propriété `theme` (attribut `theme`) sur `light`, `dark` ou `auto`. Les valeurs inconnues reviennent à `auto`, qui hérite du scope de thème Office partagé le plus proche, puis suit le schéma de couleurs du système d'exploitation lorsqu'aucun scope n'est défini.

```ts
editor.theme = 'dark';
```

## Jetons Office partagés {#shared-office-tokens}

Word utilise les mêmes jetons `--office-*` qu'Excel, Visio, PowerPoint et les contrôles partagés. Définissez `data-office-theme="light"`, `"dark"` ou `"auto"` sur un ancêtre commun pour thématiser une composition complète. Son mode et ses jetons personnalisés se mettent à jour en direct, y compris à travers la shadow root de l'éditeur. Les composants natifs React, Vue, Angular, Svelte et Solid adoptent les mêmes styles partagés.

```html
<section data-office-theme="dark" class="office-host">
	<!-- native canvas and toolbar components -->
</section>
<style>
	.office-host {
		--office-accent: #9b59b6;
		--office-ring: #b37cd0;
		--office-radius: 6px;
		--office-font: system-ui, sans-serif;
	}
</style>
```

Utilisez `--office-background` pour l'habillage de la fenêtre, `--office-surface` pour les panneaux et l'espace de travail, `--office-foreground` pour le texte de l'interface, et `--office-border` pour les contours. Les contrôles intégrés partagent aussi la typographie Office, les espacements, le focus et les jetons d'état désactivé. Les hôtes peuvent remplacer les valeurs par défaut de faible spécificité par leurs propres styles. La mise en forme du texte du document et le thème Office sérialisé restent inchangés. Le papier reste blanc par défaut en mode sombre ; `--office-paper` et la variable existante `--dve-page-color` fournissent des surcharges de présentation explicites.

Un `theme="light"` ou `theme="dark"` explicite sur l'éditeur crée un scope local pour son canevas. Utilisez l'ancêtre commun pour thématiser aussi les contrôles natifs voisins ; un éditeur ne modifie pas le thème de son application hôte. Repasser l'éditeur sur `auto` rétablit le comportement de scope hérité.

## Surcharges de jetons historiques {#legacy-token-overrides}

Les surcharges `--dve-*` restent prises en charge et l'emportent sur leur solution de repli canonique. `themeColors` accepte un `EditorTheme` partiel et applique chaque clé comme propriété personnalisée `--dve-*` sur l'hôte, en clair comme en sombre. Les valeurs peuvent être n'importe quelle couleur CSS.

```ts
editor.themeColors = {
	primary: '#c2431f',
	primaryForeground: '#ffffff',
	radius: '6px',
};
```

| Jeton                                  | Rôle                                                    |
| -------------------------------------- | ------------------------------------------------------- |
| `background`, `foreground`             | Chrome de l'application (barre de titre, fenêtre)       |
| `card`, `cardForeground`               | Surface du ruban, des panneaux et de la barre d'état    |
| `popover`, `popoverForeground`         | Menus, boîtes de dialogue et autres surfaces flottantes |
| `primary`, `primaryForeground`         | Couleur de marque et de l'action principale             |
| `secondary`, `secondaryForeground`     | Surfaces secondaires                                    |
| `muted`, `mutedForeground`             | Surface atténuée, également la zone derrière les pages  |
| `accent`, `accentForeground`           | Surbrillance au survol et à la sélection                |
| `destructive`, `destructiveForeground` | Actions destructives                                    |
| `border`, `input`, `ring`              | Bordures, contours des champs, anneau de focus          |
| `radius`                               | Rayon d'arrondi de base, par exemple `4px`              |

Passez `undefined` pour effacer les surcharges.

## Suivre le thème du site de documentation {#following-the-docs-site-theme}

Ce site enregistre le choix d'apparence du visiteur dans `localStorage` sous la clé VitePress `vitepress-theme-appearance` (`light`, `dark` ou `auto`). Tous les sites des visualiseurs partagent l'origine `christophervr.github.io`, et les applications de démo lisent cette même clé au démarrage et se mettent à jour en direct sur l'événement `storage` : changer le thème sur ce site restyle donc un éditeur intégré sans le recharger. En bref :

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

editor.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) editor.theme = resolve(event.newValue);
});
```

Une application hôte est libre d'utiliser sa propre clé : l'élément n'a besoin que de la propriété `theme`.
