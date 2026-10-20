# Theming

The editor themes its own chrome (title bar, ribbon, panels, status bar, dialogs). The document paper always stays white, like Word's page.

## Light, dark and auto

Set the `theme` property (attribute `theme`) to `light`, `dark` or `auto`. Unknown values fall back to `auto`, which inherits the nearest shared Office theme scope, then follows the operating system colour scheme when no scope is set.

```ts
editor.theme = 'dark';
```

## Shared Office tokens

Word uses the same `--office-*` tokens as Excel, Visio, PowerPoint and the shared controls.
Set `data-office-theme="light"`, `"dark"` or `"auto"` on a common ancestor to theme a complete
composition. Its mode and custom tokens update live, including across the editor's shadow root.
Native React, Vue, Angular, Svelte and Solid toolbar components opt into the same shared styles.

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

Use `--office-background` for window chrome, `--office-surface` for panels and the workspace,
`--office-foreground` for UI text, and `--office-border` for outlines. Built-in controls also
share Office typography, spacing, focus and disabled-state tokens. Hosts can override the low
specificity defaults with their own styles. The document's text formatting and serialized
Office theme remain unchanged. Paper stays white by default in dark mode; `--office-paper`
and the existing `--dve-page-color` provide explicit presentation overrides.

An explicit editor `theme="light"` or `theme="dark"` creates a local scope for its canvas.
Use the common ancestor to theme sibling native controls too; an editor does not change its
host application's theme. Returning the editor to `auto` restores inherited scope behavior.

## Legacy token overrides

`--dve-*` overrides remain supported and take precedence over their canonical fallback.
`themeColors` takes a partial `EditorTheme` and applies each key as a `--dve-*` custom property on the host, in both light and dark. Values can be any CSS colour.

```ts
editor.themeColors = {
	primary: '#c2431f',
	primaryForeground: '#ffffff',
	radius: '6px',
};
```

| Token                                  | Role                                          |
| -------------------------------------- | --------------------------------------------- |
| `background`, `foreground`             | Application chrome (title bar, window)        |
| `card`, `cardForeground`               | Ribbon, panel and status-bar surface          |
| `popover`, `popoverForeground`         | Menus, dialogs and other floating surfaces    |
| `primary`, `primaryForeground`         | Brand and primary action colour               |
| `secondary`, `secondaryForeground`     | Secondary surfaces                            |
| `muted`, `mutedForeground`             | Muted surface, also the area behind the pages |
| `accent`, `accentForeground`           | Hover and selected highlight                  |
| `destructive`, `destructiveForeground` | Destructive actions                           |
| `border`, `input`, `ring`              | Borders, input outlines, focus ring           |
| `radius`                               | Base corner radius, for example `4px`         |

Pass `undefined` to clear the overrides.

## Following the docs site theme

This site stores the visitor's appearance choice in `localStorage` under the VitePress key `vitepress-theme-appearance` (`light`, `dark` or `auto`). All the viewer sites share the `christophervr.github.io` origin, and the demo apps read the same key at start-up and update live on the `storage` event, so toggling the theme on this site restyles an embedded editor without reloading it. In short:

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

A host application is free to use its own key: the element only needs the `theme` property.
