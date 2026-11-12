# Publication des paquets

Sept paquets sans portée sont publiés sur npm sous le nom du produit **OpenTeams** (`openteams`
appartient à un projet sans rapport, donc chaque nom est plus long). Chacun a **sa propre ligne de
version**, incrémentée uniquement lorsqu’il change réellement, et son propre tag git
`<nom-npm>@<version>` (par exemple `openteams-server@0.4.0`).

Les publications se lancent depuis la **racine du dépôt**, et non depuis ce dossier. Le planificateur
(`scripts/release-plan.mjs`), la table des paquets de visionneuses (`scripts/viewer-packages.mjs`), le
workflow (`.github/workflows/release.yml`), la publication de confiance et la configuration initiale
sont décrits dans la page racine [`docs/releasing.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md).
Ce dossier ne contient ni script de publication, ni workflow propre (son `.github/` est inactif), ni
commande `release:plan` ou `publish:local`.

| Paquet                     | Dossier            | Dépend de (registre)                                        | Intégré dans             |
| -------------------------- | ------------------ | ----------------------------------------------------------- | ------------------------ |
| `openteams-react-viewer`   | `packages/react`   | `ooxml-core`, `ooxml-ui`, `lit` ; homologue `react`         | `packages/web-component` |
| `openteams-vue-viewer`     | `packages/vue`     | `ooxml-core`, `ooxml-ui`, `lit` ; homologue `vue`           | `packages/web-component` |
| `openteams-angular-viewer` | `packages/angular` | `ooxml-core`, `ooxml-ui`, `lit` ; homologue `@angular/core` | `packages/web-component` |
| `openteams-svelte-viewer`  | `packages/svelte`  | `ooxml-core`, `ooxml-ui`, `lit` ; homologue `svelte`        | `packages/web-component` |
| `openteams-solid-viewer`   | `packages/solid`   | `ooxml-core`, `ooxml-ui`, `lit` ; homologue `solid-js`      | `packages/web-component` |
| `openteams-vanilla-viewer` | `packages/vanilla` | `ooxml-core`, `ooxml-ui`, `lit`                             | `packages/web-component` |
| `openteams-server`         | `server`           | `ws`, `yjs`, `y-protocols`, `lib0`                          | rien (livre son source)  |

`teams-viewer` (`packages/web-component`, un réexport de `ooxml-ui/teams`, qui contient `<teams-app>`
et le câblage des liaisons) et les démos sont `private` et ne sont jamais publiés.
`scripts/build-packages.mjs` intègre le composant web, CSS compris, dans chaque paquet de framework
(compilation en bibliothèque avec vite) et aplatit ses déclarations dans le `dist/index.d.ts` de
chaque paquet (rollup-plugin-dts), afin qu’une archive n’importe que ce que son manifeste déclare.
Svelte livre `dist/Teams.svelte` sous forme de source, à côté d’un `dist/runtime.js` compilé (le store
et les helpers, sous `openteams-svelte-viewer/runtime`) ; Angular est un module ES regroupé dont les
décorateurs sont appliqués à l’exécution (il faut le compilateur JIT dans l’application qui l’utilise).
Le serveur est en ESM simple et n’est pas compilé ; sa commande `openteams-server` est `server/bin.mjs`.

Dans le monorepo, les manifestes des frameworks dépendent de `ooxml-core` et `ooxml-ui` avec `*` ; le
flux de publication racine réécrit ces plages en plages caret sur les versions publiées avec eux. Un
changement qui a besoin à la fois du cœur et de cette visionneuse est livré dans un seul envoi et
publié ensemble.

## Ce qui décide d’une publication ici {#what-decides-a-release-here}

Le planificateur racine lit `bundled` et `globals` dans l’entrée `teams` de
`scripts/viewer-packages.mjs` :

- **Fichiers propres.** Tout fichier non-test sous un dossier de paquet publie ce paquet. Les tests,
  `CHANGELOG.md` et les modifications de manifeste qui ne touchent que `version`, `scripts` ou
  `devDependencies` ne publient jamais rien.
- **Paquet interne intégré.** Une modification non-test sous `packages/web-component` publie chaque
  paquet de framework.
- **Chaîne de compilation partagée.** `scripts/build-packages.mjs` et `tsconfig.release.json`
  publient chaque paquet de framework. Le serveur n’est pas concerné (`global: false`) : seul
  `server/**` le publie.

Le niveau de version est le plus élevé parmi les commits Conventional Commits qui touchent les
fichiers publiés dans la portée du paquet : `!` ou `BREAKING CHANGE:` donne majeur, `feat` mineur, le
reste correctif. Utilisez `docs(teams): ...`, `fix(server): ...` et des portées semblables (voir
`AGENTS.md` à la racine).

## Vérifications exécutées pour cette visionneuse {#checks-that-run-for-this-viewer}

`bun run check:published` (aucune plage `workspace:`, `file:` ou `link:` et aucun paquet privé dans un
manifeste publié ; il analyse aussi les fichiers `.mjs` et le dossier du serveur) et `bun run pack:smoke`
(installe et importe les archives). `bun run test:scripts` exécute les tests de scripts propres à la
visionneuse (`scripts/*.test.mjs`).

La CI lance `bun run pack:smoke --workspace-runtime` : les archives sont installées avec les
`ooxml-core` et `ooxml-ui` de ce dépôt (compilés au préalable), la combinaison qu’une publication
livre ensemble. Sans cet indicateur, `pack:smoke` installe le cœur et l’interface du registre, ce
qu’obtient un utilisateur extérieur. Le `verify` racine n’exécute ces étapes de paquet qu’avec
`--full`.

Ne lancez jamais `npm publish` ni ne poussez de tags de publication à la main, et ne modifiez jamais
les versions ni les fichiers `CHANGELOG.md` : c’est le commit de publication qui s’en charge. Un
nouveau nom de paquet nécessite l’amorçage manuel et la configuration de l’éditeur de confiance
décrits sur la page racine.
