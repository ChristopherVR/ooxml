# Les démos en direct

La même application de démonstration est compilée une fois par adaptateur de framework et publiée à côté de cette documentation. Chacune est le vrai `<xlsx-editor>` qui fonctionne entièrement dans votre navigateur : les classeurs sont analysés, modifiés et enregistrés dans l’onglet, rien n’est envoyé et il n’existe aucun serveur derrière.

| Démo                                               | Adaptateur                         |
| -------------------------------------------------- | ---------------------------------- |
| [React](/demo/){target="_self"}                    | `@christophervr/xlsx-react-viewer` |
| [Vue](/demo-vue/){target="_self"}                  | `xlsx-vue-viewer`                  |
| [Angular](/demo-angular/){target="_self"}          | `xlsx-angular-viewer`              |
| [Svelte](/demo-svelte/){target="_self"}            | `xlsx-svelte-viewer`               |
| [Solid](/demo-solid/){target="_self"}              | `xlsx-solid-viewer`                |
| [JavaScript natif](/demo-vanilla/){target="_self"} | `xlsx-vanilla-viewer`              |

La page d’accueil les intègre avec un sélecteur de framework : [ouvrir la section de démo en direct](/fr/#live-demo).

## Ce que vous pouvez faire dans une démo {#what-you-can-do-in-a-demo}

- Déposer un fichier `.xlsx`, `.xlsm`, `.xls` ancien format ou `.csv`, démarrer un nouveau classeur ou ouvrir le classeur d’exemple.
- Ajouter `?sample=1` pour ouvrir directement le classeur d’exemple, et `?locale=de` (toute balise prise en charge) pour démarrer dans une autre langue d’interface.
- Utiliser le bouton de thème de la démo ou le bascule de thème de ce site : les démos suivent en direct la préférence `vitepress-theme-appearance` (voir [thèmes](/fr/theming)).

## Collaboration {#collaboration}

La démo n’a pas de page de collaboration distincte : la coédition est intégrée à l’éditeur. Choisissez **Partager**, puis démarrez une session ; une autre fenêtre ou un autre onglet du même navigateur la rejoint à partir du lien d’invitation ou du nom de la session. Consultez la [collaboration](/fr/collaboration) pour les options de salle et pour savoir ce qui est et n’est pas partagé.

## Intégrée ici {#embedded-here}

<iframe
	src="/ooxml/xlsx/demo/?sample=1"
	title="xlsx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Les exécuter localement {#run-them-locally}

```bash
bun install   # at the repository root
cd viewers/xlsx
bun run demo
```

`bun run docs:build` (depuis `docs/`) compile la documentation et les six démos exactement comme le fait le workflow Pages.
