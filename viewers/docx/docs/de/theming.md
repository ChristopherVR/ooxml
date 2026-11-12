# Theming

Der Editor gestaltet seine eigene Oberfläche selbst (Titelleiste, Multifunktionsleiste, Bereiche, Statusleiste, Dialoge). Die Dokumentseite bleibt immer weiß, wie die Seite in Word.

## Hell, dunkel und automatisch {#light-dark-and-auto}

Setzen Sie die Eigenschaft `theme` (Attribut `theme`) auf `light`, `dark` oder `auto`. Unbekannte Werte fallen auf `auto` zurück. `auto` übernimmt den nächstgelegenen gemeinsamen Office-Theme-Bereich und folgt dem Farbschema des Betriebssystems, wenn kein Bereich gesetzt ist.

```ts
editor.theme = 'dark';
```

## Gemeinsame Office-Tokens {#shared-office-tokens}

Word verwendet dieselben Tokens `--office-*` wie Excel, Visio, PowerPoint und die gemeinsamen Steuerelemente. Setzen Sie `data-office-theme="light"`, `"dark"` oder `"auto"` auf einen gemeinsamen Vorfahren, um eine vollständige Komposition zu gestalten. Modus und eigene Tokens aktualisieren sich live, auch über die Shadow Root des Editors hinweg. Native Komponenten für React, Vue, Angular, Svelte und Solid übernehmen dieselben gemeinsamen Styles.

```html
<section data-office-theme="dark" class="office-host">
	<!-- native canvas and toolbar components -->
</section>
<style>
	.office-host {
		--office-accent: #9b59b6;
		--office-ring: #b37cd0;
		--office-radius: 6px;
		--office-font: system-ui, sans-serif;
	}
</style>
```

Verwenden Sie `--office-background` für den Fensterrahmen, `--office-surface` für Bereiche und den Arbeitsbereich, `--office-foreground` für Text der Oberfläche und `--office-border` für Umrisse. Eingebaute Steuerelemente teilen sich außerdem die Tokens für Office-Typografie, Abstände, Fokus und deaktivierte Zustände. Hosts können die Standardwerte mit geringer Spezifität mit eigenen Styles überschreiben. Die Textformatierung des Dokuments und das serialisierte Office-Design bleiben unverändert. Das Papier bleibt im Dunkelmodus standardmäßig weiß; `--office-paper` und die vorhandene Variable `--dve-page-color` sind ausdrückliche Überschreibungen der Darstellung.

Ein explizites `theme="light"` oder `theme="dark"` am Editor erzeugt einen lokalen Bereich für seine Zeichenfläche. Verwenden Sie den gemeinsamen Vorfahren, um auch benachbarte native Steuerelemente zu gestalten. Ein Editor ändert nicht das Theme seiner Host-Anwendung. Wird der Editor wieder auf `auto` gesetzt, greift wieder das Verhalten des geerbten Bereichs.

## Ältere Token-Überschreibungen {#legacy-token-overrides}

`--dve-*`-Überschreibungen werden weiterhin unterstützt und haben Vorrang vor ihrem kanonischen Fallback. `themeColors` nimmt einen Teil von `EditorTheme` und wendet jeden Schlüssel als eigene Eigenschaft `--dve-*` auf dem Host an, sowohl im hellen als auch im dunklen Modus. Die Werte können beliebige CSS-Farben sein.

```ts
editor.themeColors = {
	primary: '#c2431f',
	primaryForeground: '#ffffff',
	radius: '6px',
};
```

| Token                                  | Rolle                                                       |
| -------------------------------------- | ----------------------------------------------------------- |
| `background`, `foreground`             | Anwendungsrahmen (Titelleiste, Fenster)                     |
| `card`, `cardForeground`               | Fläche von Multifunktionsleiste, Bereichen und Statusleiste |
| `popover`, `popoverForeground`         | Menüs, Dialoge und andere schwebende Flächen                |
| `primary`, `primaryForeground`         | Markenfarbe und Farbe der Hauptaktion                       |
| `secondary`, `secondaryForeground`     | Sekundäre Flächen                                           |
| `muted`, `mutedForeground`             | Gedämpfte Fläche, auch der Bereich hinter den Seiten        |
| `accent`, `accentForeground`           | Hervorhebung bei Hover und Auswahl                          |
| `destructive`, `destructiveForeground` | Aktionen, die Daten löschen                                 |
| `border`, `input`, `ring`              | Rahmen, Umrisse von Eingabefeldern, Fokusring               |
| `radius`                               | Grundlegender Eckenradius, zum Beispiel `4px`               |

Übergeben Sie `undefined`, um die Überschreibungen zu entfernen.

## Dem Design der Dokumentationsseite folgen {#following-the-docs-site-theme}

Diese Seite speichert die Darstellungswahl der Besucherin oder des Besuchers in `localStorage` unter dem VitePress-Schlüssel `vitepress-theme-appearance` (`light`, `dark` oder `auto`). Alle Viewer-Seiten teilen sich den Ursprung `christophervr.github.io`, und die Demo-Apps lesen beim Start denselben Schlüssel und aktualisieren sich live über das Ereignis `storage`. Wenn Sie das Design auf dieser Seite umschalten, wird ein eingebetteter Editor also neu gestaltet, ohne neu geladen zu werden. Kurz gefasst:

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

editor.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) editor.theme = resolve(event.newValue);
});
```

Eine Host-Anwendung darf einen eigenen Schlüssel verwenden. Das Element benötigt nur die Eigenschaft `theme`.
