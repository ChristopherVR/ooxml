# Design

Der Viewer zeichnet seine eigene Oberfläche (Multifunktionsleiste, Backstage, Statusleiste, Formenfenster) aus gemeinsamen `ooxml-ui`-Steuerelementen. Das Element liest seine Farben aus CSS-Custom-Properties, sodass ein Host es aus seinem eigenen Stylesheet umgestalten kann.

## Gemeinsame Tokens `--office-*` {#shared-office-tokens}

Das Element folgt dem gemeinsamen Office-Design. Beim Verbinden installiert es das Design von `ooxml-ui` (`installOfficeUiTheme`, idempotent), und jede Farbe, die es zeichnet, wird aus den Custom Properties `--office-*` aufgelöst. Die Tokens `--office-*` sind daher der wichtigste Weg, den Viewer zu gestalten. Setzen Sie sie auf `:root`, auf einem Container oder direkt auf dem Element:

| Token                        | Rolle                                               |
| ---------------------------- | --------------------------------------------------- |
| `--office-background`        | Fläche der Multifunktionsleiste, Panels und Dialoge |
| `--office-surface`           | Hintergrund des Fensters, sekundäre Fläche          |
| `--office-foreground`        | Primärer Text                                       |
| `--office-muted-foreground`  | Sekundärer Text                                     |
| `--office-border`            | Rahmen                                              |
| `--office-accent`            | Akzentfarbe                                         |
| `--office-accent-foreground` | Text auf der Akzentfarbe                            |
| `--office-ring`              | Fokusring                                           |
| `--office-danger`            | Zerstörerische und Fehlerzustände                   |
| `--office-shadow`            | Seitenschatten                                      |

```css
visio-viewer {
	--office-accent: #c2431f;
	--office-accent-foreground: #ffffff;
}
```

Hell und dunkel folgen dem Design der Suite: Der Viewer ist standardmäßig hell, dunkel, wenn das System dies bevorzugt, und das Attribut `data-office-theme="light" | "dark"` am Dokumentstamm erzwingt eines davon. Das Element hat keine Eigenschaft `theme`; derselbe Schalter, der die anderen Office-Elemente auf der Seite gestaltet, gestaltet auch den Viewer.

Die Standardakzentfarbe hat sich dadurch geändert. Der Viewer hat früher `#3955a3` fest vorgegeben; jetzt verwendet er `--office-accent` (`#2563eb` im hellen, `#6366f1` im dunklen Design). Ein Host, der die bisherige Farbe möchte, setzt `--office-accent: #3955a3` oder die Überschreibung `--vv-accent` unten.

## Optionale Überschreibungen `--vv-*` {#optional-vv-overrides}

Die früheren Tokens `--vv-*` funktionieren weiterhin. Jedes davon hat, wenn es gesetzt ist, Vorrang vor dem Token `--office-*`, das es nachbildet; ist es nicht gesetzt, greift der Viewer auf das gemeinsame Token zurück. Verwenden Sie sie, um nur diesen Viewer umzugestalten, ohne den Rest der Seite zu berühren.

| Token              | Fällt zurück auf             | Rolle                                      |
| ------------------ | ---------------------------- | ------------------------------------------ |
| `--vv-font`        | `--office-font`              | Schriftfamilie                             |
| `--vv-background`  | `--office-surface`           | Hintergrund des Fensters                   |
| `--vv-surface`     | `--office-background`        | Fläche der Multifunktionsleiste und Panels |
| `--vv-secondary`   | `--office-surface`           | Sekundäre Fläche                           |
| `--vv-shadow`      | `--office-shadow`            | Seitenschatten                             |
| `--vv-ink`         | `--office-foreground`        | Primärer Text                              |
| `--vv-muted`       | `--office-muted-foreground`  | Sekundärer Text                            |
| `--vv-border`      | `--office-border`            | Rahmen                                     |
| `--vv-accent`      | `--office-accent`            | Akzentfarbe                                |
| `--vv-accent-soft` | 10 % der Akzentfarbe         | Weiche Akzentfläche (Hover, Auswahl)       |
| `--vv-accent-ink`  | `--office-accent-foreground` | Text auf der Akzentfarbe                   |
| `--vv-focus`       | `--office-ring`              | Fokusring                                  |
| `--vv-danger`      | `--office-danger`            | Zerstörerische und Fehlerzustände          |

Der Playground setzt `data-theme="light" | "dark"` am Dokumentstamm und liefert passende Werte `--vv-*` in `demos/visio/demo/workspace.css` für die eigene Seitengestaltung (Startbildschirm, Workspace-Hülle). Diese Datei funktioniert weiterhin, weil Überschreibungen `--vv-*` noch wirken.

## Dem Design der Dokumentationsseite folgen {#following-the-docs-site-theme}

Diese Seite speichert die Erscheinungsoption des Besuchers im `localStorage` unter dem VitePress-Schlüssel `vitepress-theme-appearance` (`light`, `dark` oder kein Wert, der dem System folgt). Alle Viewer-Seiten teilen den Ursprung `christophervr.github.io`, und der Playground (`demos/visio/demo/suite-theme.js`) liest beim Start denselben Schlüssel und aktualisiert sich live über das Ereignis `storage`. Wenn Sie das Design auf dieser Seite umschalten, wird ein eingebetteter Viewer neu gestaltet, ohne neu geladen zu werden und ohne seine Zeichnung zu ersetzen.

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

document.documentElement.dataset.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null)
		document.documentElement.dataset.theme = resolve(event.newValue);
});
```

## Design der Statusleiste {#status-bar-theme}

Die Statusleiste ist standardmäßig neutral, wie die anderen Editoren der Suite. Das „farbige“ Office-Design der Desktop-Version von Visio füllt sie dagegen mit der Produktfarbe. Es gibt drei Wege, dies zu wählen:

- **Der Benutzer**: Datei > Optionen > Allgemein > Office-Design (Farbig, Weiß oder Schwarz; Schwarz ist das dunkle Design). Die Auswahl wird auf dem Gerät gespeichert.
- **Der Host**: das Attribut `status-bar` (`colorful` oder `neutral`) bzw. die Eigenschaft `statusBar` auf `<visio-viewer>`. Das Attribut des Hosts hat Vorrang vor der gespeicherten Auswahl.
- **Beliebige Farben**: `--vv-status-background` und `--vv-status-ink` färben die Leiste in beiden Designs neu.

```html
<visio-viewer status-bar="colorful"></visio-viewer>
<visio-viewer style="--vv-status-background: #1b5e20; --vv-status-ink: #fff"></visio-viewer>
```
