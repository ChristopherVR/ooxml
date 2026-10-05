# Theming

The viewer draws its own chrome (ribbon, backstage, status bar, Shapes window) from shared `ooxml-ui` controls. The element reads its colours from CSS custom properties, so a host can restyle it from its own stylesheet.

## Custom properties

The element maps these `--vv-*` tokens onto the shared `--office-*` tokens its controls use:

| Token              | Role                                  |
| ------------------ | ------------------------------------- |
| `--vv-background`  | Window background                     |
| `--vv-surface`     | Ribbon and panel surface              |
| `--vv-secondary`   | Secondary surface                     |
| `--vv-shadow`      | Page shadow                           |
| `--vv-ink`         | Primary text                          |
| `--vv-muted`       | Secondary text                        |
| `--vv-border`      | Borders                               |
| `--vv-accent`      | Accent colour                         |
| `--vv-accent-soft` | Soft accent surface (hover, selected) |
| `--vv-accent-ink`  | Text on the accent colour             |
| `--vv-focus`       | Focus ring                            |
| `--vv-danger`      | Destructive and error states          |

```css
visio-viewer {
	--vv-accent: #c2431f;
	--vv-accent-ink: #ffffff;
}
```

There is no `theme` property on the element in this version: light and dark come from the tokens the host provides. The playground sets `data-theme="light" | "dark"` on the document root and supplies matching token values in `demo/workspace.css`.

## Following the docs site theme

This site stores the visitor's appearance choice in `localStorage` under the VitePress key `vitepress-theme-appearance` (`light`, `dark` or no value, which follows the system). All the viewer sites share the `christophervr.github.io` origin, and the playground (`demo/suite-theme.js`) reads the same key at start-up and updates live on the `storage` event, so toggling the theme on this site restyles an embedded viewer without reloading it or replacing its diagram.

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
