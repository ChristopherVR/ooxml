# Die Live-Demos

Dieselbe Demo-App wird für jeden Framework-Adapter einzeln gebaut und neben dieser Dokumentation veröffentlicht. Jede davon ist der echte `<xlsx-editor>`, der vollständig in Ihrem Browser läuft: Arbeitsmappen werden im Tab geparst, bearbeitet und gespeichert, nichts wird hochgeladen, und dahinter steht kein Server.

| Demo                                         | Adapter                            |
| -------------------------------------------- | ---------------------------------- |
| [React](/demo/){target="_self"}              | `@christophervr/xlsx-react-viewer` |
| [Vue](/demo-vue/){target="_self"}            | `xlsx-vue-viewer`                  |
| [Angular](/demo-angular/){target="_self"}    | `xlsx-angular-viewer`              |
| [Svelte](/demo-svelte/){target="_self"}      | `xlsx-svelte-viewer`               |
| [Solid](/demo-solid/){target="_self"}        | `xlsx-solid-viewer`                |
| [Vanilla JS](/demo-vanilla/){target="_self"} | `xlsx-vanilla-viewer`              |

Die Startseite bettet sie mit einem Framework-Umschalter ein: [zum Abschnitt der Live-Demo](/de/#live-demo).

## Was Sie in einer Demo tun können {#what-you-can-do-in-a-demo}

- Eine Datei `.xlsx`, `.xlsm`, ältere `.xls` oder `.csv` hineinziehen, eine neue Arbeitsmappe anlegen oder die Beispielarbeitsmappe öffnen.
- Mit `?sample=1` öffnen Sie sofort die Beispielarbeitsmappe, und mit `?locale=de` (jedes unterstützte Tag) starten Sie in einer anderen Oberflächensprache.
- Nutzen Sie den Design-Schalter der Demo oder den Design-Umschalter dieser Website: Die Demos folgen der Einstellung `vitepress-theme-appearance` live (siehe [Themes](/de/theming)).

## Zusammenarbeit {#collaboration}

Die Demo hat keine eigene Seite für Zusammenarbeit, denn gemeinsames Bearbeiten ist im Editor bereits eingebaut. Wählen Sie **Freigeben** und starten Sie dann eine Sitzung; ein anderes Fenster oder ein anderer Tab desselben Browsers tritt ihr über den Einladungslink oder den Sitzungsnamen bei. Die Optionen des Raums sowie das, was geteilt wird und was nicht, finden Sie unter [Zusammenarbeit](/de/collaboration).

## Hier eingebettet {#embedded-here}

<iframe
	src="/ooxml/xlsx/demo/?sample=1"
	title="xlsx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Lokal ausführen {#run-them-locally}

```bash
bun install   # at the repository root
cd viewers/xlsx
bun run demo
```

`bun run docs:build` (im Ordner `docs/`) baut die Dokumentation und alle sechs Demos genau so, wie es der Pages-Workflow tut.
