---
title: Design und Farben
description: Das bindungsübergreifende ViewerTheme-System - die Form des Theme-Objekts, wie themeToCssVars die --pptx-*-Custom-Properties erzeugt, die eingebauten Voreinstellungen, die Anwendung je Bindung und die Auswahl unter Datei > Optionen > Darstellung.
---

# Design und Farben

Die Oberfläche des Viewers (Werkzeugleiste, Menüband, Dialoge, Backstage) wird über ein `ViewerTheme`-Objekt gestaltet: eine Menge von CSS-Custom-Properties (`--pptx-*`), die auf das Viewer-Wurzelelement angewendet werden. Das ist vollständig getrennt von den eigenen OOXML-Farben und Schriften einer Präsentation (die Galerie "Designs" auf der Registerkarte Entwurf bearbeitet die `.pptx`-Datei selbst). `ViewerTheme` wirkt nur auf die Oberfläche der App.

Das Theme-System ist framework-unabhängig. Die Typen, Standardwerte, Voreinstellungen und Hilfsfunktionen sind einmal in der öffentlichen API von `ooxml-ui/pptx` implementiert und werden von jeder veröffentlichten Bindung identisch erneut exportiert: `pptx-react-viewer`, `pptx-vue-viewer`, `pptx-angular-viewer`, `pptx-svelte-viewer` und `pptx-vanilla-viewer`.

## Gemeinsame Office-Darstellung {#shared-office-appearance}

Alle fünf Bindings übernehmen dieselben `--office-*`-Tokens wie Word, Excel und Visio. Setzen Sie `data-office-theme="light"`, `"dark"` oder `"auto"` auf das Dokument-Wurzelelement oder ein umschließendes Element. Bereichsbezogene Modi überschreiben die geerbte Palette; `auto` folgt dem Betriebssystem, wenn es sich ändert. Eigene Tokens im selben Bereich überschreiben die gewählte Palette:

```html
<section data-office-theme="dark" style="--office-accent: #a78bfa; --office-font: Georgia, serif">
	<!-- Mount a viewer here, without an explicit ViewerTheme palette. -->
</section>
```

Werkzeugleiste, native Oberfläche und Steuerelemente im Shadow DOM übernehmen den nächstgelegenen Bereich. Nützliche Tokens sind `--office-background`, `--office-surface`, `--office-foreground`, `--office-border`, `--office-accent`, `--office-ring` und `--office-font`. Bestehende Ausgaben von `ViewerTheme`, `themeToCssVars()` und `defaultCssVars()` bleiben kompatibel. Explizite ältere Eingaben haben weiterhin Vorrang vor den kanonischen Token-Ersatzwerten, einschließlich unvollständiger Paletten und Maße von Steuerelementen wie `--pptx-switch-width`.

Die Darstellung der Oberfläche ändert keine Farben, Schriften oder das Dokumentdesign einer Präsentation.

## Die Form von `ViewerTheme` {#the-viewertheme-shape}

```ts
interface ViewerTheme {
	/** Semantic UI colors. Each key maps to a `--pptx-<key>` custom property. */
	colors?: Partial<ViewerThemeColors>;
	/** Base border-radius value (e.g. "0.5rem", "8px"). */
	radius?: string;
	/** Escape hatch: arbitrary CSS custom properties set on the viewer root. Keys include the `--` prefix. */
	cssVars?: Record<string, string>;
}
```

Jedes Feld ist optional; nicht gesetzte Werte übernehmen die gemeinsamen Office-UI-Tokens auf der Viewer-Ebene. Explizite Farben in `ViewerTheme` und ältere `--pptx-*`-Variablen haben Vorrang. `ViewerThemeColors` enthält 19 semantische Tokens, benannt nach der shadcn/ui-Konvention. Jede gültige CSS-Farbangabe wird akzeptiert (Hex, `rgb()`, `hsl()`, `oklch()`, benannte Farben):

| Tokens                                 | Funktion                                          |
| -------------------------------------- | ------------------------------------------------- |
| `background`, `foreground`             | Hintergrund des Wurzelelements und Standardtext   |
| `card`, `cardForeground`               | Flächen für Karten und Panels                     |
| `popover`, `popoverForeground`         | Popover und Dropdowns                             |
| `primary`, `primaryForeground`         | Primäre Aktionen (Schaltflächen, aktive Zustände) |
| `secondary`, `secondaryForeground`     | Dezente Aktionen                                  |
| `muted`, `mutedForeground`             | Gedämpfte Flächen und sekundärer Text             |
| `accent`, `accentForeground`           | Flächen für die Hervorhebung beim Überfahren      |
| `destructive`, `destructiveForeground` | Gefährliche Aktionen und Löschen                  |
| `border`, `input`, `ring`              | Rahmen, Eingaberahmen, Fokusring                  |

Die vollständige Tabelle mit jedem Token und seinem exakten CSS-Variablennamen finden Sie auf der [React-Theming-Seite](/react/theming#viewertheme-and-viewerthemecolors).

## Wie `themeToCssVars` die `--pptx-*`-Variablen erzeugt {#how-themetocssvars-produces-pptx-variables}

`themeToCssVars(theme, omitDefaults = false)` wandelt ein `ViewerTheme` in ein flaches `Record<string, string>` aus CSS-Custom-Properties um, das als Inline-Stil auf das Viewer-Wurzelelement angewendet werden kann. Jede Bindung ruft es intern auf, wenn Sie ein Theme übergeben; es ist außerdem exportiert, um eigene Werkzeuge zu bauen.

- Jeder Schlüssel in `colors` wird zu `--pptx-<kebab-case-key>`: `primaryForeground: '#fff'` ergibt `--pptx-primary-foreground: #fff`.
- Jede Farbe wird außerdem in das passende Tailwind-Semantik-Token gespiegelt (`--color-primary-foreground`), für die Kompatibilität mit Tailwind-CSS-v4-Hosts. Der Viewer verwendet `@theme inline`, damit geerbte Office-Tokens an der Komponente aufgelöst werden, die sie nutzt.
- `radius` wird zu `--pptx-radius` sowie den abgeleiteten Werten `--radius-sm` / `--radius-md` / `--radius-lg` / `--radius-xl` (`calc(r - 4px)` bis `calc(r + 4px)`).
- Einträge in `cssVars` werden unverändert übernommen.
- Mit `omitDefaults: true` werden Werte übersprungen, die den eingebauten Standardwerten entsprechen.

```ts
import { themeToCssVars, defaultCssVars } from 'pptx-react-viewer';

themeToCssVars({ colors: { primary: '#6366f1' }, radius: '0.75rem' });
// {
//   '--pptx-primary': '#6366f1', '--color-primary': '#6366f1',
//   '--pptx-radius': '0.75rem',
//   '--radius-sm': 'calc(0.75rem - 4px)', ... '--radius-xl': 'calc(0.75rem + 4px)',
// }

defaultCssVars();
// The complete set of --pptx-* properties with the built-in dark defaults,
// for generating a full fallback stylesheet.
```

## Eingebaute Designs und Voreinstellungen {#built-in-themes-and-presets}

| Export                                         | Palette                                                                                                                                          |
| ---------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `defaultThemeColors` + `defaultRadius`         | Die eingebaute dunkle Oberfläche (Tailwind-Graustufen, Indigo als Primärfarbe, Radius `0.5rem`). Ausdrücklich verfügbar über `defaultCssVars()`. |
| `vermilionLightTheme` / `vermilionLightColors` | Warme helle "Papier"-Palette mit dem Zinnoberrot als Akzent, das auf dieser Dokumentationsseite verwendet wird.                                  |
| `vermilionDarkTheme` / `vermilionDarkColors`   | Gedämpfte dunkle "Vortragsraum"-Palette mit demselben Akzent.                                                                                    |
| `vermilionRadius`                              | `'0.375rem'`, der Radius, den beide Zinnoberrot-Voreinstellungen verwenden.                                                                      |

Beide Zinnoberrot-Voreinstellungen sind vollständige `ViewerTheme`-Objekte (alle 19 Tokens plus Radius) und ersetzen die dunklen Standardwerte komplett. Die reinen `*Colors`-Paletten werden daneben exportiert, um Varianten abzuleiten:

```ts
import { vermilionDarkColors, vermilionRadius } from 'pptx-react-viewer';
import type { ViewerTheme } from 'pptx-react-viewer';

const custom: ViewerTheme = {
	colors: { ...vermilionDarkColors, primary: '#38bdf8' },
	radius: vermilionRadius,
};
```

Eine einfache helle Palette (nicht Zinnoberrot) gibt es ebenfalls, aber sie ist kein benannter Export der Bindings; Sie erreichen sie über den Theme-Katalog: `resolveThemeCatalogEntry('light')`.

## Ein Theme je Bindung anwenden {#applying-a-theme-per-binding}

Jede Bindung nimmt dasselbe `ViewerTheme`-Objekt; nur die Art der Übergabe unterscheidet sich.

::: code-group

```tsx [React]
import { PowerPointViewer, vermilionDarkTheme } from 'pptx-react-viewer';

<PowerPointViewer content={bytes} theme={vermilionDarkTheme} />;
```

```vue [Vue]
<script setup lang="ts">
import { PowerPointViewer, vermilionDarkTheme } from 'pptx-vue-viewer';
</script>

<template>
	<PowerPointViewer :content="bytes" :theme="vermilionDarkTheme" />
</template>
```

```ts [Angular]
// <pptx-viewer [content]="bytes" [theme]="theme" />
import { vermilionDarkTheme } from 'pptx-angular-viewer';

export class DeckComponent {
	theme = vermilionDarkTheme;
}
// Or share one theme across a subtree without the input:
// providers: [provideViewerTheme(vermilionDarkTheme)]
```

```svelte [Svelte]
<script lang="ts">
	import { PowerPointViewer, vermilionDarkTheme } from 'pptx-svelte-viewer';
</script>

<PowerPointViewer content={bytes} theme={vermilionDarkTheme} />
```

```ts [Vanilla]
import { createPptxViewer, vermilionDarkTheme } from 'pptx-vanilla-viewer';

const viewer = createPptxViewer(host, {
	source: bytes,
	theme: vermilionDarkTheme,
});
// Change later at runtime:
viewer.setTheme({ colors: { primary: '#38bdf8' } });
```

:::

Um ein Theme über mehrere Viewer hinweg zu teilen, exportiert React `ViewerThemeProvider` / `useViewerTheme`, Vue exportiert `provideViewerTheme` / `useViewerTheme`, und Angular exportiert `provideViewerTheme` sowie das Injection-Token `VIEWER_THEME`.

## Stilmodi {#styling-modes}

Die Oberfläche des Viewers verweist für jedes visuelle Token auf `--pptx-*`-Custom-Properties. Dadurch sind drei Einrichtungen möglich, in aufsteigender Reihenfolge der Kontrolle:

| Modus                      | Einrichtung                                                                                                                                                                                                                                                 |
| -------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Tailwind-CSS-v4-Host       | Kein CSS-Import nötig; die Klassen des Viewers werden über Ihre bestehende Konfiguration aufgelöst. Überschreiben Sie Werte mit dem Prop/Input `theme`.                                                                                                     |
| Mitgelieferte Stylesheet   | `import 'pptx-react-viewer/styles'` (die Pakete für Vue und Angular stellen außerdem `/styles` und `/styles.css` bereit). Das Svelte-Paket kompiliert seine Stile in die Komponenten. Enthält alle benötigten Hilfsklassen sowie die dunklen Standardwerte. |
| Rohe CSS-Custom-Properties | Definieren Sie die `--pptx-*`-Properties selbst (die vollständige Liste finden Sie in `defaultCssVars()`) und verzichten Sie auf beides oben.                                                                                                               |

Die Vanilla-Bindung bildet die Ausnahme: `createPptxViewer` fügt automatisch sein eigenes bereichsbezogenes Stylesheet ein (idempotent, `#pptx-vanilla-viewer-styles`). Hosts mit einer strikten CSP können stattdessen den String aus `getViewerCss()` vorab rendern.

## Artwork für Auswahlsteuerelemente {#selection-control-artwork}

Das Artwork der Auswahl verwendet das bestehende Feld `ViewerTheme.cssVars`, nicht ein eigenes Theme-Objekt oder eine eigene Komponenten-Prop, und zwar in allen fünf Bindings. Diese optionalen Tokens ändern das Standardaussehen einer Bindung nicht, wenn sie weggelassen werden. Sie wirken auf die Bearbeitungssteuerelemente, nicht auf Folieninhalte, Schriftgrößen, festgelegte Geometrie oder exportierte Dokumente.

| CSS-Custom-Property                    | Steuert                                              |
| -------------------------------------- | ---------------------------------------------------- |
| `--pptx-selection-corner-size`         | Breite und Höhe des Eckartworks                      |
| `--pptx-selection-corner-radius`       | Eckenradius des Eckartworks                          |
| `--pptx-selection-edge-length`         | Länge des Kantenartworks parallel zur Kante          |
| `--pptx-selection-edge-thickness`      | Dicke des Kantenartworks quer zur Kante              |
| `--pptx-selection-edge-radius`         | Eckenradius des Kantenartworks                       |
| `--pptx-selection-handle-fill`         | Füllung des Größenänderungs-Artworks                 |
| `--pptx-selection-handle-border-color` | Rahmenfarbe des Größenänderungs- und Drehartworks    |
| `--pptx-selection-outline-color`       | Auswahlumriss und bestehender Drehstiel              |
| `--pptx-selection-rotate-size`         | Durchmesser des kreisförmigen Drehartworks           |
| `--pptx-selection-rotate-fill`         | Füllung des Drehartworks                             |
| `--pptx-selection-rotate-foreground`   | Farbe des Drehsymbols, wo ein Symbol gezeichnet wird |

Verwenden Sie positive Pixelwerte wie `6px` für Größen, CSS-Längen oder Prozentwerte für Radien und gültige CSS-Farben für Farben. Ein `var()`-Fallback greift bei einem fehlenden Token, nicht bei einem beliebigen ungültigen CSS-Wert. Der Viewer parst oder bereinigt diese CSS-Strings nicht. Entfernen Sie eine Überschreibung, um den eigenen Standardwert der Bindung wiederherzustellen. Radiuswerte verwenden die lokalen CSS-Koordinaten des Steuerelements, wie seine bestehende Rahmen- und Schattenangabe; sie werden nicht separat an die Zoomstufe der Bühne angepasst. Verwenden Sie `0px` für eckige Ecken oder `50%` für eine proportional runde Form in allen Bindings.

```ts
const theme = {
	cssVars: {
		'--pptx-selection-corner-size': '6px',
		'--pptx-selection-corner-radius': '0px',
		'--pptx-selection-edge-length': '6px',
		'--pptx-selection-edge-thickness': '6px',
		'--pptx-selection-edge-radius': '0px',
		'--pptx-selection-handle-fill': '#ffffff',
		'--pptx-selection-handle-border-color': '#6366f1',
		'--pptx-selection-outline-color': '#6366f1',
		'--pptx-selection-rotate-fill': '#ffffff',
		'--pptx-selection-rotate-foreground': '#6366f1',
	},
};
```

Übergeben Sie dasselbe Theme durch die oben gezeigten bindungsspezifischen Beispiele. Wenden Sie für eine eigene Oberfläche `themeToCssVars(theme)` auf den gemeinsamen Vorfahren von Folie und ihren Auswahlsteuerelementen an. Zielen Sie nicht auf private untergeordnete Klassen und nehmen Sie nicht an, dass Griffe in der ausgewählten Form verschachtelt sind.

Das Artwork bleibt auf den bestehenden Ankerpunkten zentriert. Ein kleineres Artwork verkleinert nicht das ursprüngliche, unsichtbare Ziel für Maus oder Berührung. Ein größeres Artwork vergrößert seinen Rahmen, und die bestehende Aufteilung der Trefferbereiche zu Nachbargriffen gilt auf sehr kleinen Formen weiterhin; übergroße Darstellungen können sich überlappen, ohne den Eingabebereich eines anderen Griffs zu übernehmen. Dieser Vertrag legt weder Trefferbereiche noch Versätze des Drehgriffs offen. Halten Sie großes eigenes Artwork für die bearbeiteten Formen praktikabel. Anpassungsrauten und Endpunktindikatoren für Verbinder behalten ihre eigene Bedeutung und ihr eigenes Aussehen.

## Datei > Optionen > Darstellung {#file-options-appearance}

Der Dialog "Einstellungen" jeder Bindung hat eine Registerkarte **Darstellung**: eine kleine Galerie eingebauter Theme-Voreinstellungen (Standard, Hell, Zinnoberrot Hell, Zinnoberrot Dunkel), die ein Benutzer zur Laufzeit durchklicken kann. Sie wird durch den Export `THEME_CATALOG` definiert:

```ts
interface ThemeCatalogEntry {
	key: string; // 'default' | 'light' | 'vermilionLight' | 'vermilionDark'
	labelKey: string; // pptx.* translation key for the entry's label
	theme: ViewerTheme | undefined; // undefined = reset to the built-in default
}
```

Das ist bewusst eine kurze, kuratierte Liste und keine vollständige Galerie. Übergeben Sie `availableThemes` (unten), um mehr oder weniger Auswahlmöglichkeiten anzubieten. `resolveThemeCatalogEntry(key, catalog?)` sucht einen Eintrag über seinen Schlüssel.

### Vorrang: Ein explizites `theme` gewinnt immer {#precedence-an-explicit-theme-always-wins}

Wenn Sie dem Viewer eine Prop/einen Input `theme` übergeben, gewinnt sie weiterhin gegenüber allem, was auf der Registerkarte Darstellung ausgewählt wird; die Auswahl ist wirkungslos, solange `theme` gesetzt ist. Das ist beabsichtigt: Ein Host, der sein eigenes Theme verwaltet (etwa die Synchronisation mit einem anwendungsweiten Umschalter für den Dunkelmodus), soll nicht durch einen Klick im Viewer überschrieben werden.

Die Registerkarte Darstellung wirkt nur, wenn Sie **kein** explizites `theme` übergeben. In diesem eigenständigen Modus gilt diese Reihenfolge:

1. Prop `defaultThemeKey` (für einen anderen Startpunkt als `'default'`)
2. Eine zuvor gespeicherte Auswahl (`localStorage`, Schlüssel `pptx-viewer-prefs`)
3. Der Eintrag `'default'` des Katalogs

### Katalog-Props (alle optional) {#catalog-props-all-optional}

| Prop               | Typ                      | Zweck                                                                                                                                                                                                                                                                                                                             |
| ------------------ | ------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `defaultThemeKey`  | `string`                 | Anfänglicher Schlüssel von `THEME_CATALOG`, nur verwendet, solange noch nichts gespeichert ist.                                                                                                                                                                                                                                   |
| `availableThemes`  | `ThemeCatalogEntry[]`    | Ersetzt den Katalog, der unter Darstellung angeboten wird (eigene Voreinstellungen hinzufügen oder die Liste eingrenzen).                                                                                                                                                                                                         |
| `onThemeChange`    | `(key: string) => void`  | Wenn angegeben, **verwalten Sie** das Speichern der Auswahl selbst; der Viewer schreibt nicht mehr in `localStorage` und ruft nur diesen Callback auf.                                                                                                                                                                            |
| `defaultLocale`    | `string`                 | Dasselbe Prinzip für die Sprache: anfänglicher Sprachcode, wenn nichts gespeichert ist. React: wird innerhalb eines `PptxI18nProvider` ignoriert, dessen `locale` die Sprache bestimmt.                                                                                                                                           |
| `availableLocales` | `LocaleCatalogEntry[]`   | Ersetzt die Sprachen, die unter Optionen > Sprache angeboten werden. React: Standard sind die Codes, die auf der Instanz `react-i18next` des Hosts registriert sind; innerhalb eines `PptxI18nProvider` gibt es keinen Standard: Übergeben Sie zwei oder mehr zusammen mit `onLocaleChange`, um die Kategorie Sprache anzuzeigen. |
| `onLocaleChange`   | `(code: string) => void` | Wenn angegeben, berührt der Viewer Ihre i18n-Instanz nie selbst; er ruft nur diesen Callback auf. React: innerhalb eines `PptxI18nProvider` bestimmt immer der Host die Sprache, und ohne diesen Callback fehlt die Kategorie Sprache.                                                                                            |

React, Vue, Angular und Svelte verwenden alle genau diese Form. Vanilla ist die einzige Ausnahme: Es hatte bereits die öffentlichen Konstruktoroptionen `theme`/`locale` sowie die Methoden `setTheme()`/`setLocale()`, daher dienen diese als Anfangswert anstelle der eigenen Optionen `defaultThemeKey`/`defaultLocale`. `availableThemes`/`availableLocales`/`onThemeChange`/`onLocaleChange` sind überall gleich.

```tsx
// React - host owns theme persistence (e.g. syncing with app-wide dark mode)
<PowerPointViewer
	content={bytes}
	defaultThemeKey={systemPrefersDark ? 'vermilionDark' : 'vermilionLight'}
	onThemeChange={(key) => saveUserPreference('viewerTheme', key)}
/>
```

### Eigene Voreinstellungen hinzufügen {#adding-your-own-presets}

Übergeben Sie `availableThemes` mit eigenen `ThemeCatalogEntry[]`; erweitern Sie die eingebauten Einträge, statt sie zu ersetzen, indem Sie `THEME_CATALOG` ausbreiten:

```ts
import { THEME_CATALOG } from 'pptx-react-viewer';

const availableThemes = [
	...THEME_CATALOG,
	{ key: 'brand', labelKey: 'app.theme.brand', theme: { colors: { primary: '#7c3aed' } } },
];
```

`labelKey` wird über das i18n-Wörterbuch nachgeschlagen, das Ihre App bereitstellt (siehe [Lokalisierung](/de/guide/localization)). Für einen eigenen Eintrag registrieren Sie diesen Schlüssel selbst. Nicht registrierte Schlüssel fallen auf eine lesbare Bezeichnung zurück, die aus dem letzten Segment des Schlüssels abgeleitet wird.

## Details je Bindung {#per-binding-details}

- [React-Theming](/react/theming) - die Prop `theme`, `ViewerThemeProvider` und die vollständige Token-Tabelle
- [Vue-Theming](/vue/theming) - die Prop `theme` und `provideViewerTheme`
- [Angular-Theming](/angular/theming) - der Input `theme`, `provideViewerTheme` und das Token `VIEWER_THEME`
- [Svelte-Theming](/svelte/theming) - die Prop `theme`
- [Vanilla-Theming](/vanilla/theming) - die Option `theme`, `setTheme()` und `getViewerCss()`

## Nächste Schritte {#next-steps}

- [Lokalisierung (i18n)](/de/guide/localization) - die Registerkarte Sprache funktioniert genauso, und hier stammen die `pptx.*`-Übersetzungsschlüssel her, die die Bezeichnungen des Themes tragen.
- [Konto und Anmeldung](/de/guide/account) - der Profil-Editor unter Datei > Konto speichert außerdem unter demselben `localStorage`-Schlüssel wie der Fallback für Theme und Sprache.
