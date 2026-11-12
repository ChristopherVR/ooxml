# Publication des paquets

Les paquets Excel sont publiés par le flux automatisé unique du dépôt, et non par un flux qui leur serait propre. Ce viewer est `viewers/xlsx` dans le dépôt [ooxml](https://github.com/ChristopherVR/ooxml), avec `ooxml-core` (`src/core`) et `ooxml-ui` (`src/ui`). L’ensemble du processus (comment une version est décidée, le niveau d’incrémentation, le workflow horaire, la publication de confiance et la configuration unique du dépôt) est décrit dans [docs/releasing.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) à la racine du dépôt.

## Ce qui est publié depuis ici {#what-is-published-from-here}

Chaque paquet a sa propre ligne de version, sa balise git `<npm-name>@<version>` et son journal des modifications. Ils sont déclarés dans `scripts/viewer-packages.mjs` à la racine du dépôt.

| Paquet                             | Dossier            | Remarques                                                                               |
| ---------------------------------- | ------------------ | --------------------------------------------------------------------------------------- |
| `@christophervr/xlsx-core`         | `packages/core`    | Point d’entrée du modèle de classeur au-dessus de `ooxml-core` ; aucune logique propre. |
| `@christophervr/xlsx-react-viewer` | `packages/react`   | Paquet d’éditeur autonome ; intègre le composant web et les liaisons.                   |
| `xlsx-vue-viewer`                  | `packages/vue`     | Même structure.                                                                         |
| `xlsx-angular-viewer`              | `packages/angular` | Même structure.                                                                         |
| `xlsx-svelte-viewer`               | `packages/svelte`  | Même structure.                                                                         |
| `xlsx-solid-viewer`                | `packages/solid`   | Même structure.                                                                         |
| `xlsx-vanilla-viewer`              | `packages/vanilla` | Même structure ; possède aussi le point d’entrée `<xlsx-editor>` simple.                |
| `xlsx-viewer-mcp`                  | `mcp`              | Serveur MCP autonome, voir son README.                                                  |

Les paquets internes `packages/web-component` (`xlsx-web-component`) et `packages/bindings` (`xlsx-bindings`) sont `private` et ne sont jamais publiés : `scripts/build-packages.mjs` les intègre à chaque paquet de framework, si bien qu’une modification hors test dans l’un ou l’autre publie les six. Une archive publiée n’importe que `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` et sa dépendance de pair de framework ; `ole2` est intégré dans `ooxml-core`. `ooxml-core` et `ooxml-ui` sont de vraies dépendances du registre, jamais intégrées au paquet.

## Vérifications depuis ce dossier {#checks-from-this-folder}

Exécutez-les depuis `viewers/xlsx` après un `bun install` à la racine :

```sh
bun run build:packages
bun run check:published   # no tarball imports an internal package or ole2
bun run pack:smoke        # install the tarballs into a clean consumer and load .xlsx and .xls
```

Le planificateur de publication, les garde-fous de publication et le workflow s’exécutent depuis la racine du dépôt (`bun run release:plan`, `node scripts/publish-released.mjs --dry-run`, `.github/workflows/release.yml`). Ne lancez jamais `npm publish` ni ne poussez de balises de publication à la main.

## Conventions de commit {#commit-conventions}

Le niveau d’incrémentation vient du type de commit (`!` ou `BREAKING CHANGE:` correspond à une version majeure, `feat` à une version mineure, tout le reste à un correctif), et les fichiers touchés par un commit déterminent si un paquet est publié. Les règles et exemples figurent dans [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md#commit-conventions).
