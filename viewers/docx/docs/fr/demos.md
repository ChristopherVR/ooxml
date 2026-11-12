# Les démos en direct

La même application de démonstration est compilée une fois pour chaque adaptateur de framework et publiée à côté de cette documentation. Chacune est le vrai `<docx-editor>` s'exécutant entièrement dans votre navigateur : rien n'est envoyé, et aucun serveur n'est derrière.

| Démo                                               | Adaptateur            | Source                                                             |
| -------------------------------------------------- | --------------------- | ------------------------------------------------------------------ |
| [React](/demo/){target="_self"}                    | `docx-react-viewer`   | `demos/docx/demo-vanilla` compilé avec `VITE_DEMO_FRAMEWORK=react` |
| [Vue](/demo-vue/){target="_self"}                  | `docx-vue-viewer`     | même application, `vue`                                            |
| [Angular](/demo-angular/){target="_self"}          | `docx-angular-viewer` | même application, `angular`                                        |
| [Svelte](/demo-svelte/){target="_self"}            | `docx-svelte-viewer`  | même application, `svelte`                                         |
| [Solid](/demo-solid/){target="_self"}              | `docx-solid-viewer`   | même application, `solid`                                          |
| [JavaScript natif](/demo-vanilla/){target="_self"} | `docx-vanilla-viewer` | même application, `vanilla`                                        |

La page d'accueil les intègre avec un sélecteur de framework : [ouvrir la section de démo en direct](/fr/#live-demo).

## Ce que vous pouvez faire dans une démo {#what-you-can-do-in-a-demo}

- Choisissez **Ouvrir le document d'exemple**, ou déposez un fichier DOCX ou DOC ancien de votre choix. Les fichiers restent dans l'onglet.
- Ajoutez `?sample=1` à l'URL d'une démo pour ouvrir directement le document d'exemple, et `?locale=de` (également `fr`, `es`, `zh-CN`) pour démarrer dans une autre langue d'interface.
- Utilisez le bouton de thème de la démo, ou le sélecteur de thème de ce site : toutes les démos suivent en direct la préférence `vitepress-theme-appearance` (voir [thèmes](/fr/theming)).

## Démo de collaboration {#collaboration-demo}

Ajoutez `?room=<session>` à l'URL d'une démo pour rejoindre une session partagée du même navigateur. La fenêtre qui a aussi `sample=1` héberge la session ; les autres la rejoignent par son nom, quel que soit le framework : par exemple `/demo-vue/?sample=1&room=my-room` et `/demo-solid/?room=my-room` dans deux onglets. La page d'accueil affiche deux volets avec un sélecteur de framework pour chacun. La [page à deux pairs](/demo/collaboration.html){target="_self"} ajoute un bouton Pause de la diffusion pour les modifications simultanées. Consultez [collaboration](/fr/collaboration) pour savoir ce qui est pris en charge ou non.

## Intégré ici {#embedded-here}

<iframe
	src="/ooxml/docx/demo/?sample=1"
	title="docx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Les exécuter en local {#run-them-locally}

```bash
bun install          # from the repository root
cd viewers/docx
bun run demo
```

`bun run docs:build` (depuis `docs/`) compile la documentation et les six démos exactement comme le fait le workflow Pages.
