# Theming

`<teams-app>` bezieht jede Farbe, jeden Abstandswert und jeden Radius aus den gemeinsamen Tokens von `ooxml-ui`
(`var(--office-...)`). Benutzerdefinierte Eigenschaften überschreiten Shadow-Grenzen. Eine Regel auf `:root` (oder auf
einem Container) gestaltet daher die ganze App.

## Hell und dunkel {#light-and-dark}

Die Tokens folgen standardmäßig dem Farbschema des Betriebssystems. Sie erzwingen ein Schema, indem Sie
`data-office-theme` am Root setzen:

```ts
document.documentElement.setAttribute('data-office-theme', 'dark'); // or 'light'
document.documentElement.removeAttribute('data-office-theme'); // follow the OS again
```

## Token überschreiben {#token-overrides}

Überschreiben Sie jedes Token dort, wo es gelten soll:

```css
:root {
	--office-accent: #0e8f8f;
	--office-teams-brand: #0e8f8f;
	--office-teams-rail: #14171b;
}
```

Die Teams-spezifischen Tokens tragen das Präfix `--office-teams-` (Marke, Leiste, Feld, Blase, Trennlinie und
gedämpfter Text). Die App hat keine eigene Eigenschaft `theme`: Gestaltet wird mit diesen Tokens. Die Akzentfarbe der
Dokumentation auf dieser Website ist bewusst ein neutrales Blaugrün; OpenTeams übernimmt keine Markenfarben von Microsoft.

## Dem Design der Dokumentationsseite folgen {#following-the-docs-site-theme}

Diese Website speichert die Erscheinungsauswahl des Besuchers im `localStorage` unter dem VitePress-Schlüssel
`vitepress-theme-appearance` (`light`, `dark` oder `auto`). Alle Viewer-Websites teilen den Ursprung
`christophervr.github.io`. Die Demo-Apps in diesem Repository (`demos/teams/theme-sync.ts` im Root des Repositorys)
übernehmen sie beim Start und beim Ereignis `storage`. Wenn Sie also das Design auf dieser Website umschalten, wird
eine eingebettete Demo ohne Neuladen neu gestaltet:

```ts
const key = 'vitepress-theme-appearance';
const apply = (value: string | null) => {
	const root = document.documentElement;
	if (value === 'light' || value === 'dark') root.setAttribute('data-office-theme', value);
	else root.removeAttribute('data-office-theme');
};
apply(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) apply(event.newValue);
});
```
