# Die Live-Demos

Derselbe Playground wird je Framework-Adapter einmal gebaut und neben dieser Dokumentation veröffentlicht. Jeder ist das echte `<visio-viewer>`, das vollständig in Ihrem Browser läuft: Zeichnungen werden lokal geparst und gerendert, nichts wird hochgeladen, und es gibt keinen Server dahinter. Dies ist eine öffentliche Beta; dass eine Demo eine Datei öffnet, ist kein Beleg für die Treue zu Visio.

| Demo                                      | Adapter                | Route            |
| ----------------------------------------- | ---------------------- | ---------------- |
| [Vanilla JS](/demo/){target="_self"}      | `visio-vanilla-viewer` | `/demo/`         |
| [React](/demo-react/){target="_self"}     | `visio-react-viewer`   | `/demo-react/`   |
| [Vue](/demo-vue/){target="_self"}         | `visio-vue-viewer`     | `/demo-vue/`     |
| [Angular](/demo-angular/){target="_self"} | `visio-angular-viewer` | `/demo-angular/` |
| [Svelte](/demo-svelte/){target="_self"}   | `visio-svelte-viewer`  | `/demo-svelte/`  |
| [Solid](/demo-solid/){target="_self"}     | `visio-solid-viewer`   | `/demo-solid/`   |

Der Vanilla-Playground ist außerdem unter [/demo-vanilla/](/demo-vanilla/){target="_self"} erreichbar, dem Namen, den die Seiten der anderen Viewer verwenden. Die Startseite bettet die Demos mit einem Framework-Umschalter ein: [Abschnitt mit der Live-Demo öffnen](/de/#live-demo).

## Was Sie in einer Demo tun können {#what-you-can-do-in-a-demo}

- Öffnen Sie die originale Beispielzeichnung oder eine lokale `.vsdx`-Datei Ihrer Wahl.
- Fügen Sie einer Demo-URL `?sample=1` hinzu, um die Beispielzeichnung sofort zu öffnen, und `?embed=1` für das kompakte eingebettete Layout, das die Startseite verwendet.
- Verwenden Sie die Design-Schaltfläche der Demo oder den Design-Umschalter dieser Seite: Alle Demos folgen live der Einstellung `vitepress-theme-appearance` (siehe [Design](/de/theming)).

## Demo zur Zusammenarbeit {#collaboration-demo}

Fügen Sie einer Demo-URL `?share=<session>` hinzu, um beim Laden dieser Freigabesitzung beizutreten. Zwei Demos aus unterschiedlichen Frameworks, die denselben Sitzungsnamen verwenden, teilen sich eine Zeichnung, zum Beispiel `/demo-vue/?sample=1&share=my-room` und `/demo-angular/?share=my-room` in zwei Tabs desselben Browsers. Die Startseite zeigt zwei Bereiche mit einer Framework-Auswahl für jeden. Was unterstützt wird und was nicht, steht unter [Zusammenarbeit](/de/collaboration).

## Hier eingebettet {#embedded-here}

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1"
	title="visio-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Lokal ausführen {#run-them-locally}

```bash
bun install            # once, at the repository root
cd viewers/visio
bun run dev
```

Die Quellen der Demos liegen im Repository-Stamm in `demos/visio/` (`demo` ist der gemeinsame Vanilla-Playground; `demo-react`, `demo-vue`, `demo-angular`, `demo-svelte` und `demo-solid` binden denselben Workspace ein). Die Skripte und die Vite-Konfiguration des Viewers bauen und stellen sie bereit. `bun run docs:build` baut die Dokumentation und alle Demos genau so, wie es der Pages-Build des Repositorys tut (`node scripts/build-pages.mjs` im Stamm baut die Seite jedes Viewers).
