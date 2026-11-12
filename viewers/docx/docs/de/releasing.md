# Pakete veröffentlichen

Die Word-Viewer-Pakete werden über den eigenen Ablauf des Repositorys `ooxml` veröffentlicht, zusammen mit `ooxml-core` und `ooxml-ui`. In `viewers/docx` gibt es keinen eigenen Release-Workflow. Der gesamte Ablauf (der Planer, der stündliche Workflow `release.yml`, vertrauenswürdiges Veröffentlichen, Changelogs und die einmalige Einrichtung) ist im [Leitfaden zum Veröffentlichen](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) des Repositorys beschrieben. Führen Sie niemals `npm publish` aus und pushen Sie niemals Release-Tags von Hand.

## Was dieser Viewer veröffentlicht {#what-this-viewer-publishes}

Jedes Paket hat seine eigene Versionslinie und seinen eigenen Git-Tag `<npm-name>@<version>`. Die Zuordnung von Verzeichnis zu npm-Name steht in `scripts/viewer-packages.mjs` im Root des Repositorys.

| Paket                 | Verzeichnis        | Hinweise                                                             |
| --------------------- | ------------------ | -------------------------------------------------------------------- |
| `docx-core`           | `packages/core`    | Schlanke Neuexportierung von `ooxml-core/docx` und `/docx/embedded`. |
| `docx-react-viewer`   | `packages/react`   | React-Einstiegspunkt (`WordEditor`).                                 |
| `docx-vue-viewer`     | `packages/vue`     | Vue-3-Einstiegspunkt.                                                |
| `docx-angular-viewer` | `packages/angular` | Angular-Einstiegspunkt (`WordEditorComponent`).                      |
| `docx-svelte-viewer`  | `packages/svelte`  | Svelte-5-Einstiegspunkt.                                             |
| `docx-solid-viewer`   | `packages/solid`   | Solid-Einstiegspunkt.                                                |
| `docx-vanilla-viewer` | `packages/vanilla` | Einfache Registrierung von `<docx-editor>` und `mountEditor`.        |
| `docx-viewer-mcp`     | `mcp`              | Der Word-MCP-Server (siehe sein README und das Root-`mcp.md`).       |

`web-component` und `bindings` sind `private: true` und werden nie veröffentlicht. `scripts/build-packages.mjs` bündelt sie in jedes Framework-Paket. Ein Tarball importiert daher nur `docx-core`, `ooxml-core`, `ooxml-ui`, die ProseMirror-Bibliotheken und seine Framework-Peer-Abhängigkeit. `ooxml-ui` und `docx-core` bleiben echte Abhängigkeiten aus der Registry (nie gebündelt), und die älteren `.doc`-Codecs aus `ole2` sind in `ooxml-core` eingebettet. Deshalb ist `ole2` für nichts hier eine Abhängigkeit.

## Was eine Veröffentlichung auslöst {#what-triggers-a-release}

Der Planer des Root-Verzeichnisses (`scripts/release-plan.mjs`, mit der Tabelle aus `scripts/viewer-packages.mjs`) betrachtet die konventionellen Commits seit dem letzten Tag jedes Pakets:

- Eine Änderung außerhalb von Tests unter `packages/<name>/` veröffentlicht dieses Paket.
- Eine Änderung außerhalb von Tests unter `packages/web-component` oder `packages/bindings` veröffentlicht jedes Framework-Paket, weil jedes diesen Code ausliefert. `scripts/build-packages.mjs` und `tsconfig.release.json` veröffentlichen in diesem Viewer alles.
- Eine Veröffentlichung von `docx-core` veröffentlicht die Framework-Pakete erneut, die von ihm abhängen.
- Tests, Dokumentation, Demos und CI veröffentlichen nie etwas. Eine Änderung unter `demos/docx/` oder `e2e/docx/` führt nur die Prüfungen dieses Viewers in CI aus.

Der eigene Code des Editors liegt in `src/ui/src/docx` und wird in `ooxml-ui` ausgeliefert. Eine Änderung dort veröffentlicht `ooxml-ui`, und die Viewer-Pakete übernehmen sie über ihre Abhängigkeit von `ooxml-ui`.

## Prüfungen vor dem Push {#checks-before-pushing}

Führen Sie `bun run verify` im Root des Repositorys aus. Für diesen Viewer entsprechen die lokalen Prüfungen aus `viewers/docx` diesen Befehlen:

```sh
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run check:published && bun run pack:smoke
```

`check:published` schlägt fehl, wenn ein Tarball ein internes Paket oder `ole2` importiert. `pack:smoke` installiert die Tarballs in einen sauberen Verbraucher und lädt DOCX- und ältere `.doc`-Dateien über jedes Framework-Paket.
