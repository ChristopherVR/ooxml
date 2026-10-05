# Theming

`<teams-app>` takes every colour, spacing value and radius from the shared `ooxml-ui` tokens (`var(--office-...)`). Custom properties cross shadow boundaries, so one rule on `:root` (or on a container) themes the whole app.

## Light and dark

The tokens follow the operating system colour scheme by default. Force a scheme by setting `data-office-theme` on the root:

```ts
document.documentElement.setAttribute('data-office-theme', 'dark'); // or 'light'
document.documentElement.removeAttribute('data-office-theme'); // follow the OS again
```

## Token overrides

Override any token where you want it to apply:

```css
:root {
	--office-accent: #0e8f8f;
	--office-teams-brand: #0e8f8f;
	--office-teams-rail: #14171b;
}
```

The Teams-specific tokens are prefixed `--office-teams-` (brand, rail, panel, bubble, divider and subtle text). The app has no `theme` property of its own: theming is done with these tokens. The docs accent on this site is a neutral teal on purpose; OpenTeams does not borrow Microsoft's brand colours.

## Following the docs site theme

This site keeps the visitor's appearance choice in `localStorage` under the VitePress key `vitepress-theme-appearance` (`light`, `dark` or `auto`). All the viewer sites share the `christophervr.github.io` origin, and the demo apps in this repository (`demos/theme-sync.ts`) apply it at start-up and on the `storage` event, so toggling the theme on this site restyles an embedded demo without a reload:

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
