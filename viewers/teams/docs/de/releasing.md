# Pakete veröffentlichen

Sieben Pakete ohne Scope werden unter dem Produktnamen **OpenTeams** auf npm veröffentlicht (`openteams` selbst gehört
einem unabhängigen Projekt, daher ist jeder Name länger). Jedes hat **eine eigene Versionslinie**, die nur erhöht wird,
wenn sich das Paket tatsächlich ändert, und einen eigenen Git-Tag `<npm-name>@<version>` (zum Beispiel
`openteams-server@0.4.0`).

Veröffentlichungen laufen aus dem **Repository-Root**, nicht aus diesem Ordner. Der Planer
(`scripts/release-plan.mjs`), die Tabelle der Viewer-Pakete (`scripts/viewer-packages.mjs`), der Workflow
(`.github/workflows/release.yml`), das vertrauenswürdige Veröffentlichen und die einmalige Einrichtung sind in der
Root-Seite [`docs/releasing.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) beschrieben.
Dieser Ordner hat keine eigenen Veröffentlichungsskripte, keinen eigenen Workflow (sein `.github/` ist inaktiv) und
keine Befehle `release:plan` oder `publish:local`.

| Paket                      | Verzeichnis        | Hängt ab von (Registry)                               | Gebündelt in                          |
| -------------------------- | ------------------ | ----------------------------------------------------- | ------------------------------------- |
| `openteams-react-viewer`   | `packages/react`   | `ooxml-core`, `ooxml-ui`, `lit`; Peer `react`         | `packages/web-component`              |
| `openteams-vue-viewer`     | `packages/vue`     | `ooxml-core`, `ooxml-ui`, `lit`; Peer `vue`           | `packages/web-component`              |
| `openteams-angular-viewer` | `packages/angular` | `ooxml-core`, `ooxml-ui`, `lit`; Peer `@angular/core` | `packages/web-component`              |
| `openteams-svelte-viewer`  | `packages/svelte`  | `ooxml-core`, `ooxml-ui`, `lit`; Peer `svelte`        | `packages/web-component`              |
| `openteams-solid-viewer`   | `packages/solid`   | `ooxml-core`, `ooxml-ui`, `lit`; Peer `solid-js`      | `packages/web-component`              |
| `openteams-vanilla-viewer` | `packages/vanilla` | `ooxml-core`, `ooxml-ui`, `lit`                       | `packages/web-component`              |
| `openteams-server`         | `server`           | `ws`, `yjs`, `y-protocols`, `lib0`                    | nichts (liefert seinen Quellcode aus) |

`teams-viewer` (`packages/web-component`, eine Neuexportierung von `ooxml-ui/teams`, die `<teams-app>` und die
Verdrahtung der Bindings enthält) sowie die Demos sind `private` und werden nie veröffentlicht. `scripts/build-packages.mjs`
bettet die Web-Komponente einschließlich CSS in jedes Framework-Bundle ein (Bibliotheksbuild mit Vite) und flacht ihre
Deklarationen in die `dist/index.d.ts` jedes Pakets ab (rollup-plugin-dts). So importiert ein Tarball nur, was sein
Manifest deklariert. Svelte liefert `dist/Teams.svelte` als Quelltext neben einer gebündelten `dist/runtime.js` (der
Store und die Hilfsfunktionen, unter `openteams-svelte-viewer/runtime`). Angular ist ein gebündeltes ES-Modul, dessen
Decorators zur Laufzeit angewendet werden (die verwendende App braucht dafür den JIT-Compiler). Der Server ist reines
ESM und wird nicht gebaut; sein Befehl `openteams-server` ist `server/bin.mjs`.

Im Monorepo hängen die Framework-Manifeste mit `*` von `ooxml-core` und `ooxml-ui` ab. Der Veröffentlichungsablauf
im Root schreibt diese Bereiche in Caret-Bereiche auf die Versionen um, die mit ihnen veröffentlicht werden. Eine
Änderung, die sowohl den Core als auch diesen Viewer braucht, landet in einem einzigen Push und wird gemeinsam
veröffentlicht.

## Was hier über eine Veröffentlichung entscheidet {#what-decides-a-release-here}

Der Planer im Root liest `bundled` und `globals` aus dem Eintrag `teams` von `scripts/viewer-packages.mjs`:

- **Eigene Dateien.** Jede Nicht-Test-Datei unter einem Paketverzeichnis veröffentlicht das Paket. Tests, `CHANGELOG.md`
  und Manifest-Änderungen, die nur `version`, `scripts` oder `devDependencies` betreffen, tun dies nie.
- **Gebündeltes internes Paket.** Eine Nicht-Test-Änderung unter `packages/web-component` veröffentlicht jedes
  Framework-Paket.
- **Gemeinsame Build-Pipeline.** `scripts/build-packages.mjs` und `tsconfig.release.json` veröffentlichen jedes
  Framework-Paket. Der Server ist ausgenommen (`global: false`): Nur `server/**` veröffentlicht ihn.

Die Erhöhungsstufe ist die höchste Conventional-Commit-Stufe unter den Commits, die veröffentlichte Dateien im Geltungsbereich
des Pakets berühren: `!` oder `BREAKING CHANGE:` ist major, `feat` minor, alles andere patch. Verwenden Sie Bereiche wie
`docs(teams): ...`, `fix(server): ...` und ähnliche (siehe `AGENTS.md` im Root).

## Prüfungen für diesen Viewer {#checks-that-run-for-this-viewer}

`bun run check:published` (kein Bereich `workspace:`, `file:` oder `link:` und kein privates Paket in einem
veröffentlichten Manifest; es prüft auch `.mjs`-Dateien und das Serververzeichnis) und `bun run pack:smoke`
(Tarballs installieren und importieren). `bun run test:scripts` führt die eigenen Skripttests des Viewers aus
(`scripts/*.test.mjs`).

Die CI führt `bun run pack:smoke --workspace-runtime` aus: Die Tarballs werden mit dem `ooxml-core` und `ooxml-ui`
dieses Checkouts installiert (zuerst gebaut), also der Kombination, die gemeinsam veröffentlicht wird. Ohne das Flag
installiert `pack:smoke` Core und UI aus der Registry, so wie es ein externer Nutzer täte. Das `verify` im Root führt
diese Paketschritte nur mit `--full` aus.

Führen Sie nie `npm publish` oder Release-Tags von Hand aus und bearbeiten Sie nie Versionen oder `CHANGELOG.md`-Dateien;
das Release-Commit erledigt das. Für einen neuen Paketnamen sind die manuelle Erstinstallation und die Einrichtung des
vertrauenswürdigen Publishers nötig, die auf der Root-Seite beschrieben sind.
