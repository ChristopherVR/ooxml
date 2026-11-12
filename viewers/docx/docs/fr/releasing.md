# Publication des paquets

Les paquets du visualiseur Word sont publiés par le flux propre au dépôt `ooxml`, en même temps que `ooxml-core` et `ooxml-ui`. Il n'existe aucun workflow de publication distinct dans `viewers/docx`. L'ensemble du flux (le planificateur, `release.yml` exécuté chaque heure, la publication de confiance, les journaux de modifications et la configuration initiale) est décrit dans le [guide de publication](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) du dépôt. Ne lancez jamais `npm publish` ni n'envoyez de balises de publication à la main.

## Ce que ce visualiseur publie {#what-this-viewer-publishes}

Chaque paquet a sa propre ligne de versions et sa propre balise git `<npm-name>@<version>`. La correspondance entre le répertoire et le nom npm se trouve dans `scripts/viewer-packages.mjs` à la racine du dépôt.

| Paquet                | Répertoire         | Remarques                                                               |
| --------------------- | ------------------ | ----------------------------------------------------------------------- |
| `docx-core`           | `packages/core`    | Réexportation directe de `ooxml-core/docx` et de `/docx/embedded`.      |
| `docx-react-viewer`   | `packages/react`   | Point d'entrée React (`WordEditor`).                                    |
| `docx-vue-viewer`     | `packages/vue`     | Point d'entrée Vue 3.                                                   |
| `docx-angular-viewer` | `packages/angular` | Point d'entrée Angular (`WordEditorComponent`).                         |
| `docx-svelte-viewer`  | `packages/svelte`  | Point d'entrée Svelte 5.                                                |
| `docx-solid-viewer`   | `packages/solid`   | Point d'entrée Solid.                                                   |
| `docx-vanilla-viewer` | `packages/vanilla` | Enregistrement de `<docx-editor>` en JavaScript natif et `mountEditor`. |
| `docx-viewer-mcp`     | `mcp`              | Le serveur MCP Word (voir son README et le `mcp.md` racine).            |

`web-component` et `bindings` sont `private: true` et ne sont jamais publiés. `scripts/build-packages.mjs` les intègre dans chaque paquet de framework : un tarball n'importe donc que `docx-core`, `ooxml-core`, `ooxml-ui`, les bibliothèques ProseMirror et sa dépendance peer du framework. `ooxml-ui` et `docx-core` restent de vraies dépendances du registre (jamais intégrées), et les codecs `.doc` hérités de `ole2` sont intégrés dans `ooxml-core`, de sorte que `ole2` n'est une dépendance d'aucun élément ici.

## Ce qui déclenche une publication {#what-triggers-a-release}

Le planificateur racine (`scripts/release-plan.mjs`, avec la table de `scripts/viewer-packages.mjs`) examine les commits conventionnels depuis la dernière balise de chaque paquet :

- Une modification hors tests sous `packages/<name>/` publie ce paquet.
- Une modification hors tests sous `packages/web-component` ou `packages/bindings` publie chaque paquet de framework, car chacun embarque ce code. `scripts/build-packages.mjs` et `tsconfig.release.json` publient tout ce visualiseur.
- Une publication de `docx-core` republie les paquets de framework qui en dépendent.
- Les tests, la documentation, les démos et la CI ne publient jamais rien. Une modification sous `demos/docx/` ou `e2e/docx/` lance seulement les vérifications de ce visualiseur dans la CI.

L'éditeur lui-même vit dans `src/ui/src/docx` et est livré dans `ooxml-ui` ; une modification à cet endroit publie `ooxml-ui`, et les paquets du visualiseur la reprennent grâce à leur plage `ooxml-ui`.

## Vérifications avant l'envoi {#checks-before-pushing}

Lancez `bun run verify` à la racine du dépôt. Pour ce visualiseur, les vérifications locales équivalentes, depuis `viewers/docx`, sont :

```sh
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run check:published && bun run pack:smoke
```

`check:published` échoue si un tarball importe un paquet interne ou `ole2`, et `pack:smoke` installe les tarballs dans un consommateur propre et charge les fichiers DOCX et `.doc` hérités via chaque paquet de framework.
