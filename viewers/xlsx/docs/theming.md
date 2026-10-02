# Theming

The editor ships a light and a dark theme for its chrome (ribbon, formula bar, headers, tabs, dialogs). The `theme` attribute or property selects one: `light`, `dark`, or `auto` (the default) to follow the operating system's `prefers-color-scheme`. Cell content is always drawn with the workbook's own colours on a white sheet, like Excel.

```html
<xlsx-editor theme="dark"></xlsx-editor>
```

## Token overrides

`themeColors` overrides individual tokens. Each key becomes a `--xve-<kebab-key>` custom property on the host element, with the same token set the Word editor uses (`--dve-*` there), so one design system can drive both:

| Key                                    | Used for                                  |
| -------------------------------------- | ----------------------------------------- |
| `background`, `foreground`             | Editor canvas and text                    |
| `card`, `cardForeground`               | Ribbon and panels                         |
| `popover`, `popoverForeground`         | Menus and dropdowns                       |
| `primary`, `primaryForeground`         | Accent (title bar, selection, active tab) |
| `secondary`, `secondaryForeground`     | Secondary buttons                         |
| `muted`, `mutedForeground`             | Headers, hints, disabled text             |
| `accent`, `accentForeground`           | Hover and pressed states                  |
| `destructive`, `destructiveForeground` | Errors                                    |
| `border`, `input`, `ring`              | Lines, fields and focus rings             |
| `radius`                               | Corner radius                             |

```ts
editor.themeColors = { primary: '#1f9d63', ring: '#1f9d63', radius: '6px' };
```

You can also set the custom properties in CSS:

```css
xlsx-editor {
	--xve-primary: #1f9d63;
}
```

## Following a site theme

The demo follows the VitePress `vitepress-theme-appearance` key in `localStorage` and sets `theme` on every editor when it changes (see `demos/demo-vanilla/theme.ts`). Do the same with your own site setting; the editor never reads `localStorage` itself.
