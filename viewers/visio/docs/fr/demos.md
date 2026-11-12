# Les démos en direct

Le même terrain de jeu est compilé une fois par adaptateur de framework et publié à côté de cette documentation. Chacun est le vrai `<visio-viewer>` qui s'exécute entièrement dans votre navigateur : les dessins sont analysés et rendus localement, rien n'est envoyé et aucun serveur ne se trouve derrière. Il s'agit d'une version bêta publique ; une démo qui ouvre un fichier n'est pas une preuve de fidélité à Visio.

| Démo                                         | Adaptateur             | Route            |
| -------------------------------------------- | ---------------------- | ---------------- |
| [JavaScript vanilla](/demo/){target="_self"} | `visio-vanilla-viewer` | `/demo/`         |
| [React](/demo-react/){target="_self"}        | `visio-react-viewer`   | `/demo-react/`   |
| [Vue](/demo-vue/){target="_self"}            | `visio-vue-viewer`     | `/demo-vue/`     |
| [Angular](/demo-angular/){target="_self"}    | `visio-angular-viewer` | `/demo-angular/` |
| [Svelte](/demo-svelte/){target="_self"}      | `visio-svelte-viewer`  | `/demo-svelte/`  |
| [Solid](/demo-solid/){target="_self"}        | `visio-solid-viewer`   | `/demo-solid/`   |

Le terrain de jeu vanilla est aussi servi à [/demo-vanilla/](/demo-vanilla/){target="_self"}, le nom qu'utilisent les sites des autres visualiseurs. La page d'accueil intègre les démos avec un sélecteur de framework : [ouvrir la section de démo en direct](/fr/#live-demo).

## Ce que vous pouvez faire dans une démo {#what-you-can-do-in-a-demo}

- Ouvrir le dessin d'exemple original, ou un fichier `.vsdx` local que vous choisissez.
- Ajouter `?sample=1` à l'URL d'une démo pour ouvrir l'exemple immédiatement, et `?embed=1` pour la disposition compacte intégrée utilisée par la page d'accueil.
- Utiliser le bouton de thème de la démo ou le sélecteur de thème de ce site : toutes les démos suivent en direct la préférence `vitepress-theme-appearance` (voir [thèmes](/fr/theming)).

## Démo de collaboration {#collaboration-demo}

Ajoutez `?share=<session>` à n'importe quelle URL de démo pour rejoindre cette session de partage au chargement. Deux démos construites avec des frameworks différents qui utilisent le même nom de session partagent un seul dessin, par exemple `/demo-vue/?sample=1&share=my-room` et `/demo-angular/?share=my-room` dans deux onglets d'un même navigateur. La page d'accueil affiche deux volets avec un sélecteur de framework pour chacun. Voir [collaboration](/fr/collaboration) pour ce qui est et n'est pas pris en charge.

## Intégrée ici {#embedded-here}

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1"
	title="visio-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Les exécuter en local {#run-them-locally}

```bash
bun install            # once, at the repository root
cd viewers/visio
bun run dev
```

Les sources des démos se trouvent dans `demos/visio/` à la racine du dépôt (`demo` est le terrain de jeu vanilla partagé, `demo-react`, `demo-vue`, `demo-angular`, `demo-svelte` et `demo-solid` montent le même espace de travail). Les scripts du visualiseur et sa configuration Vite les compilent et les servent. `bun run docs:build` compile la documentation et toutes les démos exactement comme le fait la construction Pages du dépôt (`node scripts/build-pages.mjs` à la racine construit le site de chaque visualiseur).
