# Thèmes

Le visualiseur dessine sa propre interface (ruban, coulisses, barre d'état, fenêtre des formes) à partir des commandes partagées de `ooxml-ui`. L'élément lit ses couleurs dans des propriétés personnalisées CSS, de sorte qu'un hôte peut le restyler depuis sa propre feuille de style.

## Jetons `--office-*` partagés {#shared-office-tokens}

L'élément suit le thème Office partagé. À sa connexion, il installe le thème `ooxml-ui` (`installOfficeUiTheme`, idempotent), et chaque couleur qu'il dessine se résout à partir des propriétés personnalisées `--office-*` : les jetons `--office-*` sont donc le moyen principal de le thématiser. Définissez-les sur `:root`, sur un conteneur ou sur l'élément lui-même :

| Jeton                        | Rôle                                                   |
| ---------------------------- | ------------------------------------------------------ |
| `--office-background`        | Surface du ruban, des volets et des boîtes de dialogue |
| `--office-surface`           | Arrière-plan de la fenêtre, surface secondaire         |
| `--office-foreground`        | Texte principal                                        |
| `--office-muted-foreground`  | Texte secondaire                                       |
| `--office-border`            | Bordures                                               |
| `--office-accent`            | Couleur d'accentuation                                 |
| `--office-accent-foreground` | Texte sur la couleur d'accentuation                    |
| `--office-ring`              | Anneau de focus                                        |
| `--office-danger`            | États destructifs et d'erreur                          |
| `--office-shadow`            | Ombre de la page                                       |

```css
visio-viewer {
	--office-accent: #c2431f;
	--office-accent-foreground: #ffffff;
}
```

Le mode clair et le mode sombre suivent le thème de la suite : le visualiseur est clair par défaut, sombre lorsque le système le préfère, et l'attribut `data-office-theme="light" | "dark"` sur la racine du document en impose un. L'élément n'a pas de propriété `theme` ; le même commutateur qui thématise les autres éléments Office de la page thématise aussi le visualiseur.

La couleur d'accentuation par défaut a changé à cette occasion. Le visualiseur codait auparavant `#3955a3` en dur ; il prend désormais `--office-accent` (`#2563eb` en clair, `#6366f1` en sombre). Un hôte qui veut l'ancienne couleur définit `--office-accent: #3955a3` ou le remplacement `--vv-accent` ci-dessous.

## Remplacements `--vv-*` facultatifs {#optional-vv-overrides}

Les anciens jetons `--vv-*` fonctionnent toujours. Chacun, lorsqu'il est défini, l'emporte sur le jeton `--office-*` qu'il alias ; lorsqu'il n'est pas défini, le visualiseur revient au jeton partagé. Utilisez-les pour restyler ce seul visualiseur sans toucher au reste de la page.

| Jeton              | Revient à                    | Rôle                                             |
| ------------------ | ---------------------------- | ------------------------------------------------ |
| `--vv-font`        | `--office-font`              | Famille de police                                |
| `--vv-background`  | `--office-surface`           | Arrière-plan de la fenêtre                       |
| `--vv-surface`     | `--office-background`        | Surface du ruban et des volets                   |
| `--vv-secondary`   | `--office-surface`           | Surface secondaire                               |
| `--vv-shadow`      | `--office-shadow`            | Ombre de la page                                 |
| `--vv-ink`         | `--office-foreground`        | Texte principal                                  |
| `--vv-muted`       | `--office-muted-foreground`  | Texte secondaire                                 |
| `--vv-border`      | `--office-border`            | Bordures                                         |
| `--vv-accent`      | `--office-accent`            | Couleur d'accentuation                           |
| `--vv-accent-soft` | 10 % de l'accentuation       | Surface d'accentuation douce (survol, sélection) |
| `--vv-accent-ink`  | `--office-accent-foreground` | Texte sur la couleur d'accentuation              |
| `--vv-focus`       | `--office-ring`              | Anneau de focus                                  |
| `--vv-danger`      | `--office-danger`            | États destructifs et d'erreur                    |

La démo définit `data-theme="light" | "dark"` sur la racine du document et fournit des valeurs `--vv-*` correspondantes dans `demos/visio/demo/workspace.css` pour sa propre interface de page (écran d'accueil, enveloppe de l'espace de travail) ; ce fichier continue de fonctionner, car les remplacements `--vv-*` restent appliqués.

## Suivre le thème du site de documentation {#following-the-docs-site-theme}

Ce site stocke le choix d'apparence du visiteur dans `localStorage`, sous la clé VitePress `vitepress-theme-appearance` (`light`, `dark` ou aucune valeur, qui suit le système). Tous les sites des visualiseurs partagent l'origine `christophervr.github.io`, et la démo (`demos/visio/demo/suite-theme.js`) lit la même clé au démarrage et se met à jour en direct sur l'événement `storage` ; changer le thème sur ce site restyle donc un visualiseur intégré sans le recharger ni remplacer son diagramme.

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

document.documentElement.dataset.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null)
		document.documentElement.dataset.theme = resolve(event.newValue);
});
```

## Thème de la barre d'état {#status-bar-theme}

La barre d'état est neutre par défaut, comme les autres éditeurs de la suite. Le thème Office « Coloré » de Visio pour ordinateur la remplit avec la couleur du produit. Il existe trois façons de choisir :

- **L'utilisateur** : Fichier > Options > Général > Thème Office (Colorful, White ou Black, qui est le thème sombre). Le choix est conservé sur l'appareil.
- **L'hôte** : l'attribut `status-bar` (`colorful` ou `neutral`), ou la propriété `statusBar`, sur `<visio-viewer>`. L'attribut d'un hôte l'emporte sur le choix enregistré.
- **Des couleurs quelconques** : `--vv-status-background` et `--vv-status-ink` recolorent la barre dans l'un ou l'autre thème.

```html
<visio-viewer status-bar="colorful"></visio-viewer>
<visio-viewer style="--vv-status-background: #1b5e20; --vv-status-ink: #fff"></visio-viewer>
```
