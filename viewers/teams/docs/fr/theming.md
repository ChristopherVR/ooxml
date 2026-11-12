# Thèmes

`<teams-app>` prend chaque couleur, valeur d’espacement et rayon dans les jetons partagés de
`ooxml-ui` (`var(--office-...)`). Les propriétés personnalisées traversent les frontières du shadow
DOM : une seule règle sur `:root` (ou sur un conteneur) thématise toute l’application.

## Clair et sombre {#light-and-dark}

Les jetons suivent par défaut le thème de couleurs du système d’exploitation. Pour forcer un thème,
définissez `data-office-theme` sur la racine :

```ts
document.documentElement.setAttribute('data-office-theme', 'dark'); // or 'light'
document.documentElement.removeAttribute('data-office-theme'); // follow the OS again
```

## Surcharges de jetons {#token-overrides}

Surchargez n’importe quel jeton là où vous voulez qu’il s’applique :

```css
:root {
	--office-accent: #0e8f8f;
	--office-teams-brand: #0e8f8f;
	--office-teams-rail: #14171b;
}
```

Les jetons propres à Teams portent le préfixe `--office-teams-` (marque, barre, panneau, bulle,
séparateur et texte secondaire). L’application n’a pas de propriété `theme` : le thème se règle avec
ces jetons. L’accent de la documentation de ce site est un sarcelle neutre, voulu ainsi : OpenTeams
n’emprunte pas les couleurs de marque de Microsoft.

## Suivre le thème du site de documentation {#following-the-docs-site-theme}

Ce site conserve le choix d’apparence du visiteur dans `localStorage`, sous la clé VitePress
`vitepress-theme-appearance` (`light`, `dark` ou `auto`). Toutes les visionneuses partagent l’origine
`christophervr.github.io`, et les applications de démonstration de ce dépôt (`demos/teams/theme-sync.ts`
à la racine du dépôt) l’appliquent au démarrage et sur l’événement `storage` : basculer le thème sur ce
site restyle donc une démo intégrée sans rechargement.

```ts
const key = 'vitepress-theme-appearance';
const apply = (value: string | null) => {
	const root = document.documentElement;
	if (value === 'light' || value === 'dark') root.setAttribute('data-office-theme', value);
	else root.removeAttribute('data-office-theme');
};
apply(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) apply(event.newValue);
});
```
