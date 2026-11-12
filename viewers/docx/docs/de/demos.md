# Die Live-Demos

Dieselbe Demo-App wird für jeden Framework-Adapter einmal erstellt und neben dieser Dokumentation veröffentlicht. Jede ist der echte `<docx-editor>`, der vollständig in Ihrem Browser läuft: Es wird nichts hochgeladen, und dahinter steht kein Server.

| Demo                                         | Adapter               | Quelle                                                              |
| -------------------------------------------- | --------------------- | ------------------------------------------------------------------- |
| [React](/demo/){target="_self"}              | `docx-react-viewer`   | `demos/docx/demo-vanilla`, erstellt mit `VITE_DEMO_FRAMEWORK=react` |
| [Vue](/demo-vue/){target="_self"}            | `docx-vue-viewer`     | dieselbe App, `vue`                                                 |
| [Angular](/demo-angular/){target="_self"}    | `docx-angular-viewer` | dieselbe App, `angular`                                             |
| [Svelte](/demo-svelte/){target="_self"}      | `docx-svelte-viewer`  | dieselbe App, `svelte`                                              |
| [Solid](/demo-solid/){target="_self"}        | `docx-solid-viewer`   | dieselbe App, `solid`                                               |
| [Vanilla JS](/demo-vanilla/){target="_self"} | `docx-vanilla-viewer` | dieselbe App, `vanilla`                                             |

Die Startseite bettet sie mit einem Framework-Umschalter ein: [zum Abschnitt der Live-Demo](/de/#live-demo).

## Was Sie in einer Demo tun können {#what-you-can-do-in-a-demo}

- Wählen Sie **Beispieldokument öffnen** oder ziehen Sie eine eigene DOCX- oder ältere DOC-Datei hinein. Die Dateien bleiben im Tab.
- Fügen Sie `?sample=1` an die URL einer Demo an, um das Beispieldokument sofort zu öffnen, und `?locale=de` (auch `fr`, `es`, `zh-CN`), um in einer anderen Oberflächensprache zu starten.
- Verwenden Sie die Designschaltfläche der Demo oder den Umschalter für Hell und Dunkel auf dieser Website: Alle Demos folgen der Einstellung `vitepress-theme-appearance` live (siehe [Theming](/de/theming)).

## Kollaborations-Demo {#collaboration-demo}

Fügen Sie `?room=<session>` an die URL einer Demo an, um einer gemeinsamen Sitzung im selben Browser beizutreten. Das Fenster, das außerdem `sample=1` hat, hostet die Sitzung. Die anderen treten über den Namen bei, in jedem beliebigen Framework: zum Beispiel `/demo-vue/?sample=1&room=my-room` und `/demo-solid/?room=my-room` in zwei Tabs. Die Startseite zeigt zwei Bereiche mit je einer Framework-Auswahl. Die [Seite mit zwei Teilnehmern](/demo/collaboration.html){target="_self"} fügt eine Schaltfläche Pause der Zustellung für gleichzeitige Bearbeitungen hinzu. Was unterstützt wird und was nicht, steht unter [Zusammenarbeit](/de/collaboration).

## Hier eingebettet {#embedded-here}

<iframe
	src="/ooxml/docx/demo/?sample=1"
	title="docx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Lokal ausführen {#run-them-locally}

```bash
bun install          # from the repository root
cd viewers/docx
bun run demo
```

`bun run docs:build` (im Ordner `docs/`) erstellt die Dokumentation und alle sechs Demos genau so, wie es der Pages-Workflow tut.
