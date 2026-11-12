# Pakete veröffentlichen

Die Excel-Pakete werden nicht über einen eigenen Ablauf veröffentlicht, sondern über den einen automatisierten Ablauf des Repositorys. Dieser Viewer ist `viewers/xlsx` im Repository [ooxml](https://github.com/ChristopherVR/ooxml), zusammen mit `ooxml-core` (`src/core`) und `ooxml-ui` (`src/ui`). Der gesamte Ablauf (wie über eine Veröffentlichung entschieden wird, die Stufe der Versionserhöhung, der stündliche Workflow, Trusted Publishing und die einmalige Einrichtung des Repositorys) ist in [docs/releasing.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) im Wurzelverzeichnis des Repositorys beschrieben.

## Was von hier veröffentlicht wird {#what-is-published-from-here}

Jedes Paket hat seine eigene Versionslinie, sein eigenes Git-Tag `<npm-name>@<version>` und sein eigenes Changelog. Sie sind in `scripts/viewer-packages.mjs` im Wurzelverzeichnis des Repositorys deklariert.

| Paket                              | Verzeichnis        | Hinweise                                                                           |
| ---------------------------------- | ------------------ | ---------------------------------------------------------------------------------- |
| `@christophervr/xlsx-core`         | `packages/core`    | Einstiegspunkt des Arbeitsmappenmodells über `ooxml-core`; hat keine eigene Logik. |
| `@christophervr/xlsx-react-viewer` | `packages/react`   | Eigenständiges Editor-Paket; bündelt die Web Component und die Bindings.           |
| `xlsx-vue-viewer`                  | `packages/vue`     | Gleiche Struktur.                                                                  |
| `xlsx-angular-viewer`              | `packages/angular` | Gleiche Struktur.                                                                  |
| `xlsx-svelte-viewer`               | `packages/svelte`  | Gleiche Struktur.                                                                  |
| `xlsx-solid-viewer`                | `packages/solid`   | Gleiche Struktur.                                                                  |
| `xlsx-vanilla-viewer`              | `packages/vanilla` | Gleiche Struktur; enthält außerdem den einfachen Einstiegspunkt `<xlsx-editor>`.   |
| `xlsx-viewer-mcp`                  | `mcp`              | Eigenständiger MCP-Server, siehe seine README.                                     |

Die internen Pakete `packages/web-component` (`xlsx-web-component`) und `packages/bindings` (`xlsx-bindings`) sind `private` und werden nie veröffentlicht: `scripts/build-packages.mjs` bündelt sie in jedes Framework-Paket, daher betrifft eine Änderung außerhalb der Tests in einem von beiden alle sechs Pakete. Ein veröffentlichtes Tarball importiert nur `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` und seine Framework-Peer-Abhängigkeit; `ole2` ist in `ooxml-core` eingebettet. `ooxml-core` und `ooxml-ui` sind echte Registry-Abhängigkeiten und werden nie gebündelt.

## Prüfungen aus diesem Ordner {#checks-from-this-folder}

Führen Sie diese nach einem `bun install` im Wurzelverzeichnis aus `viewers/xlsx` aus:

```sh
bun run build:packages
bun run check:published   # no tarball imports an internal package or ole2
bun run pack:smoke        # install the tarballs into a clean consumer and load .xlsx and .xls
```

Der Release-Planer, die Veröffentlichungsprüfungen und der Workflow laufen aus dem Wurzelverzeichnis des Repositorys (`bun run release:plan`, `node scripts/publish-released.mjs --dry-run`, `.github/workflows/release.yml`). Führen Sie niemals `npm publish` aus und setzen Sie Release-Tags nie von Hand.

## Commit-Konventionen {#commit-conventions}

Die Stufe der Versionserhöhung ergibt sich aus dem Commit-Typ (`!` oder `BREAKING CHANGE:` bedeutet major, `feat` bedeutet minor, alles andere patch), und die Dateien, die ein Commit berührt, entscheiden, ob ein Paket veröffentlicht wird. Regeln und Beispiele stehen in [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md#commit-conventions).
